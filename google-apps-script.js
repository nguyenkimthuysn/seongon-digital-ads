/**
 * SEONGON – Google Apps Script Backend
 * ─────────────────────────────────────
 * Google Sheet: https://docs.google.com/spreadsheets/d/1_2XveLa7-xbsK7kI3QWr98mtL4vn6yTRvqnurcB8dUw/edit
 *
 * HƯỚNG DẪN DEPLOY (3 bước):
 *
 * 1. Mở Google Sheet trên → Extensions → Apps Script
 * 2. Xóa code mặc định, paste toàn bộ file này
 * 3. Deploy → New deployment:
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    → Authorize → Copy Web App URL
 * 4. Dán URL vào index.html tại biến APPS_SCRIPT_URL
 *
 * LƯU Ý: Sheet cần có 2 tab tên chính xác:
 *   - "Kiểm tra tài khoản"
 *   - "Liên hệ"
 */

const SHEET_ID = '1_2XveLa7-xbsK7kI3QWr98mtL4vn6yTRvqnurcB8dUw';

function doPost(e) {
  try {
    const ss   = SpreadsheetApp.openById(SHEET_ID);
    const data = JSON.parse(e.postData.contents);

    if (data.formType === 'healthcheck') {
      const sheet = ss.getSheetByName('Kiểm tra tài khoản');

      // Kiểm tra cell A1 có header chưa (tránh lỗi do sheet có empty rows)
      if (sheet.getRange('A1').getValue() === '') {
        sheet.clearContents();
        sheet.getRange('A1').setValue('Thời gian');
        sheet.getRange(1, 1, 1, 6).setValues([[
          'Thời gian', 'Họ và tên', 'Số điện thoại',
          'Kênh cần kiểm tra', 'Link tài sản số', 'CID / ID tài khoản'
        ]]);
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold')
          .setBackground('#0A2AE0').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
        data.name,
        data.phone,
        data.channels,
        data.link,
        data.cid
      ]);

    } else if (data.formType === 'contact') {
      const sheet = ss.getSheetByName('Liên hệ');

      if (sheet.getRange('A1').getValue() === '') {
        sheet.clearContents();
        sheet.getRange(1, 1, 1, 5).setValues([[
          'Thời gian', 'Họ và tên', 'Số điện thoại', 'Email', 'Dịch vụ cần tư vấn'
        ]]);
        sheet.getRange(1, 1, 1, 5).setFontWeight('bold')
          .setBackground('#0A2AE0').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
        data.name,
        data.phone,
        data.email,
        data.service
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test bằng cách chạy hàm này trong Apps Script editor
function testHealthCheck() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        formType : 'healthcheck',
        name     : 'Nguyễn Test',
        phone    : '0901234567',
        channels : 'Google Ads, Facebook Ads',
        link     : 'https://example.com',
        cid      : '123-456-7890'
      })
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}

function testContact() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        formType : 'contact',
        name     : 'Trần Test',
        phone    : '0987654321',
        email    : 'test@example.com',
        service  : 'Google Ads'
      })
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}
