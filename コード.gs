// スプレッドシートID (ここを実際のものに書き換えてください)
const SS_ID = '1kkprVtF4Iog7lGKP3-CJQqsjtwpWpJdr-n_O82YGy3Y';

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  const id = e.parameter.id;
  const mode = e.parameter.mode;
  
  if (!id) {
    template.viewMode = 'create';
    template.bookingData = {};
  } else {
    const bookingData = getBookingById(id);
    template.bookingData = bookingData;
    
    if (mode === 'confirm') {
      template.viewMode = 'confirm';
    } else {
      template.viewMode = 'guest';
    }
  }

  return template.evaluate()
    .setTitle('超楽調整（会議調整アプリ）') // タイトルも合わせて変更
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAvailabilityConfig() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('空き時間');
  const data = sheet.getDataRange().getValues();
  return data;
}

function createDraftAndSave(formData) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const id = Utilities.getUuid();
  
  const newRow = [
    id,
    formData.meetingName,
    JSON.stringify(formData.host), 
    JSON.stringify(formData.attendees), 
    '', 
    '調整中', 
    '', 
    ''  
  ];
  sheet.appendRow(newRow);
  
  const appUrl = ScriptApp.getService().getUrl();
  const guestUrl = `${appUrl}?id=${id}`;
  const attendeesEmails = formData.attendees.map(a => a.email).join(',');
  const hostEmail = formData.host.email;
  const subject = `お打合せ時間調整のお願い（${formData.meetingName}）`;
  
  let body = `いつもお世話になっております、${formData.host.company} ${formData.host.name}です。\n\n`;
  body += `掲題のお打合せについて、ご調整をお願いできますでしょうか？\n`;
  body += `下記アドレスからご都合日の指定を頂くか、本メールのご返信にて調整日をお知らせください。\n\n`;
  body += `${guestUrl}\n\n`;
  body += `よろしくお願いいたします。`;
  
  let htmlBody = `いつもお世話になっております、${formData.host.company} ${formData.host.name}です。<br><br>`;
  htmlBody += `掲題のお打合せについて、ご調整をお願いできますでしょうか？<br>`;
  htmlBody += `下記アドレスからご都合日の指定を頂くか、本メールのご返信にて調整日をお知らせください。<br><br>`;
  htmlBody += `<a href="${guestUrl}">会議調整URL</a><br><br>`;
  htmlBody += `よろしくお願いいたします。`;
  
  GmailApp.createDraft(attendeesEmails, subject, body, {
    cc: hostEmail,
    htmlBody: htmlBody
  });
  
  // ★修正3: メッセージ文言変更
  return { success: true, message: '会議開催通知の下書きを作成しました。' };
}

function getBookingById(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  
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
        rowIndex: i + 1
      };
    }
  }
  return null;
}

function handleGuestReply(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const booking = getBookingById(data.id);
  
  if (!booking) throw new Error('Booking not found');
  
  sheet.getRange(booking.rowIndex, 6).setValue('相手返信済');
  sheet.getRange(booking.rowIndex, 7).setValue(data.reserverName);
  sheet.getRange(booking.rowIndex, 8).setValue(data.reserverEmail);
  
  if (data.selectedDate) {
    sheet.getRange(booking.rowIndex, 5).setValue(data.selectedDate);
  }

  const appUrl = ScriptApp.getService().getUrl();
  const confirmUrl = `${appUrl}?id=${data.id}&mode=confirm`;
  
  const subject = `【連絡】日程調整について（${booking.meetingName}）`;
  let body = `相手先（${data.reserverName}）より連絡が届いています。\n\n`;
  body += `希望日時 : ${data.selectedDate || '日時指定なし'}\n`;
  body += `メッセージ内容：${data.message || 'なし'}\n\n`;
  body += `▼日程を確定するには以下のURLへアクセスしてください\n`;
  body += `${confirmUrl}`;
  
  MailApp.sendEmail(booking.host.email, subject, body);
  
  return { success: true };
}

function confirmBooking(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Bookings');
  const booking = getBookingById(data.id);
  if (!booking) throw new Error('Booking not found');

  sheet.getRange(booking.rowIndex, 6).setValue('予約確定');
  sheet.getRange(booking.rowIndex, 5).setValue(data.finalDate);

  // カレンダー登録
  let meetUrl = '';
  const dateParts = parseDateTimeString(data.finalDate);
  if (dateParts) {
    const calendar = CalendarApp.getDefaultCalendar();
    
    // イベント作成
    const event = calendar.createEvent(
      `${booking.meetingName}（${booking.reserverName} 様）`,
      dateParts.start,
      dateParts.end,
      {
        description: '会議調整アプリより登録',
        guests: `${booking.reserverEmail},${booking.host.email}`,
        location: 'Google Meet' 
      }
    );
    
    // ★修正6: Meet URLの取得
    meetUrl = event.getHangoutLink();
  }
  
  // 完了メール送信
  const subject = `【確定】会議日程のお知らせ（${booking.meetingName}）`;
  let body = `会議日時が確定しました。\n\n`;
  body += `日時: ${data.finalDate}\n`;
  body += `場所: Google Meet\n`;
  
  // ★修正6: 会議URLを追記
  if (meetUrl) {
    body += `会議URL: ${meetUrl}\n`;
  } else {
    // 取得できなかった場合のフォールバック（または何も表示しない）
    // body += `会議URL: カレンダーをご確認ください\n`; 
  }
  
  MailApp.sendEmail([booking.reserverEmail, booking.host.email].join(','), subject, body);
  
  return { success: true };
}

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