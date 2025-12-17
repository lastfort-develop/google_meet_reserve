// スプレッドシートID (ここを実際のものに書き換えてください)
const SS_ID = '1kkprVtF4Iog7lGKP3-CJQqsjtwpWpJdr-n_O82YGy3Y';

/**
 * ウェブアプリへのアクセス時に実行される関数
 * URLパラメータによって表示モードを切り替えます
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  
  const id = e.parameter.id;
  const mode = e.parameter.mode; // 'confirm' など
  
  // 初期状態（新規作成）
  if (!id) {
    template.viewMode = 'create';
    template.bookingData = {};
  } else {
    // IDがある場合はデータを取得
    const bookingData = getBookingById(id);
    template.bookingData = bookingData;
    
    if (mode === 'confirm') {
      template.viewMode = 'confirm'; // 主催者確定モード
    } else {
      template.viewMode = 'guest'; // ゲスト調整モード
    }
  }

  return template.evaluate()
    .setTitle('Google Meet 会議調整アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 必要なHTMLファイルをインクルードするためのヘルパー関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 空き時間シートからタイムテーブル設定を取得
 */
function getAvailabilityConfig() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('空き時間');
  const data = sheet.getDataRange().getValues();
  // 1行目はヘッダー(曜日)、A列は時間帯
  return data; 
}

/**
 * 新規会議予定を保存し、下書きメールを作成する
 */
function createDraftAndSave(formData) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  
  // UUID生成
  const id = Utilities.getUuid();
  
  // 保存データ構築
  const newRow = [
    id,
    formData.meetingName,
    JSON.stringify(formData.host), // 主催者情報
    JSON.stringify(formData.attendees), // 出席者情報
    '', // 予約日時(未定)
    '調整中', // 状態
    '', // 予約者名
    ''  // 予約者メール
  ];
  
  sheet.appendRow(newRow);
  
  // アプリのURL取得
  const appUrl = ScriptApp.getService().getUrl();
  const guestUrl = `${appUrl}?id=${id}`;
  
  // メール送信先
  const attendeesEmails = formData.attendees.map(a => a.email).join(',');
  const hostEmail = formData.host.email;
  const subject = `お打合せ時間調整のお願い（${formData.meetingName}）`;

  // 1. プレーンテキスト版本文（HTML非対応メーラー用）
  let body = `いつもお世話になっております、${formData.host.company} ${formData.host.name}です。\n\n`;
  body += `掲題のお打合せについて、ご調整をお願いできますでしょうか？\n`;
  body += `下記アドレスからご都合日の指定を頂くか、本メールのご返信にて調整日をお知らせください。\n\n`;
  body += `${guestUrl}\n\n`; 
  body += `よろしくお願いいたします。`;
  
  // 2. HTML版本文（こちらが優先表示され、リンクになります）
  let htmlBody = `いつもお世話になっております、${formData.host.company} ${formData.host.name}です。<br><br>`;
  htmlBody += `掲題のお打合せについて、ご調整をお願いできますでしょうか？<br>`;
  htmlBody += `下記アドレスからご都合日の指定を頂くか、本メールのご返信にて調整日をお知らせください。<br><br>`;
  htmlBody += `<a href="${guestUrl}">会議調整URL</a><br><br>`; // ★アンカータグに変更
  htmlBody += `よろしくお願いいたします。`;
  
  GmailApp.createDraft(attendeesEmails, subject, body, {
    cc: hostEmail,
    htmlBody: htmlBody
  });
  
  return { success: true, message: '下書きを作成し、会議を登録しました。' };
}

/**
 * IDから予約情報を取得する
 */
function getBookingById(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  
  // ヘッダーを除いて検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return {
        id: data[i][0],
        meetingName: data[i][1],
        host: JSON.parse(data[i][2]),
        attendees: JSON.parse(data[i][3]),
        confirmedDate: data[i][4],
        status: data[i][5],
        reserverName: data[i][6],
        reserverEmail: data[i][7],
        rowIndex: i + 1 // 更新用に行番号を保持
      };
    }
  }
  return null;
}

/**
 * ゲストからの返信処理
 */
function handleGuestReply(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const booking = getBookingById(data.id);
  
  if (!booking) throw new Error('Booking not found');
  
  // DB更新
  sheet.getRange(booking.rowIndex, 6).setValue('相手返信済');
  sheet.getRange(booking.rowIndex, 7).setValue(data.reserverName);
  sheet.getRange(booking.rowIndex, 8).setValue(data.reserverEmail);
  
  if (data.selectedDate) {
    sheet.getRange(booking.rowIndex, 5).setValue(data.selectedDate); 
  }

  // 主催者へ通知メール
  const appUrl = ScriptApp.getService().getUrl();
  const confirmUrl = `${appUrl}?id=${data.id}&mode=confirm`; // 確定用URL
  
  const subject = `【連絡】日程調整について（${booking.meetingName}）`;
  let body = `相手先（${data.reserverName}）より連絡が届いています。\n\n`;
  body += `希望日時 : ${data.selectedDate || '日時指定なし'}\n`;
  body += `メッセージ内容：${data.message || 'なし'}\n\n`;
  body += `▼日程を確定するには以下のURLへアクセスしてください\n`;
  body += `${confirmUrl}`;
  
  MailApp.sendEmail(booking.host.email, subject, body);
  
  return { success: true };
}

/**
 * 主催者による確定処理
 */
function confirmBooking(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const booking = getBookingById(data.id);
  
  if (!booking) throw new Error('Booking not found');

  // 1. DB更新
  sheet.getRange(booking.rowIndex, 6).setValue('予約確定');
  sheet.getRange(booking.rowIndex, 5).setValue(data.finalDate);
  
  // 2. カレンダー登録
  const dateParts = parseDateTimeString(data.finalDate);
  if (dateParts) {
    const calendar = CalendarApp.getDefaultCalendar(); // 主催者のカレンダー
    calendar.createEvent(
      `${booking.meetingName}（${booking.reserverName} 様）`,
      dateParts.start,
      dateParts.end,
      {
        description: '会議調整アプリより登録',
        guests: `${booking.reserverEmail},${booking.host.email}`,
        location: 'Google Meet' 
      }
    );
  }
  
  // 3. 完了メール送信
  const subject = `【確定】会議日程のお知らせ（${booking.meetingName}）`;
  let body = `会議日時が確定しました。\n\n`;
  body += `日時: ${data.finalDate}\n`;
  body += `場所: Google Meet\n`;
  
  // 予約者と主催者に送信
  MailApp.sendEmail([booking.reserverEmail, booking.host.email].join(','), subject, body);
  
  return { success: true };
}

// 日時文字列パース用ヘルパー
function parseDateTimeString(dateStr) {
  try {
    const match = dateStr.match(/(\d{4})\/(\d{1,2})\/(\d{1,2}).*?(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/);
    if (match) {
      const year = parseInt(match[1]);
      const month = parseInt(match[2]) - 1;
      const day = parseInt(match[3]);
      
      const start = new Date(year, month, day, parseInt(match[4]), parseInt(match[5]));
      const end = new Date(year, month, day, parseInt(match[6]), parseInt(match[7]));
      
      return { start: start, end: end };
    }
  } catch (e) {
    console.error(e);
  }
  return null;
}