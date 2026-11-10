/**
 * Nhận đăng ký từ web giới thiệu CES Anh – Trung – Kids và ghi vào Google Sheet này.
 * Hai tab tự tạo ở lần đăng ký đầu: "Đối tác" và "Giáo viên".
 * Dán toàn bộ tệp này vào: Sheet > Tiện ích mở rộng > Apps Script, rồi Triển khai > Ứng dụng web.
 */

// true: mỗi đăng ký mới gửi một email báo về tài khoản đã triển khai script này.
const BAO_EMAIL = true;

const COT = {
  'doi-tac':   { tab: 'Đối tác',   tieuDe: ['Thời gian', 'Họ tên', 'SĐT / Zalo', 'Tỉnh / thành', 'Trung tâm / trường', 'Số HS 4–10 tuổi', 'Hỏi thêm', 'Trạng thái', 'Người phụ trách', 'Ghi chú CES'],
                 giaTri: d => [d.name, d.phone, d.city, d.org, d.size, d.note] },
  'giao-vien': { tab: 'Giáo viên', tieuDe: ['Thời gian', 'Họ tên', 'SĐT / Zalo', 'Tỉnh / thành', 'Dạy', 'Kinh nghiệm dạy trẻ', 'Hỏi thêm', 'Trạng thái', 'Người phụ trách', 'Ghi chú CES'],
                 giaTri: d => [d.name, d.phone, d.city, d.lang, d.exp, d.note] },
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    // Chặn máy gửi tự động: ô ẩn "website" bị điền, hoặc gửi trong chưa tới 3 giây sau khi mở trang.
    if (d.website || (d.secs !== undefined && d.secs < 3)) return traLoi({ ok: true });
    const loai = COT[d.who] ? d.who : 'doi-tac';
    if (!d.name || !d.phone) return traLoi({ ok: false, loi: 'thieu_ten_hoac_sdt' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = COT[loai];
    let sh = ss.getSheetByName(cfg.tab);
    if (!sh) {
      sh = ss.insertSheet(cfg.tab);
      sh.appendRow(cfg.tieuDe);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, cfg.tieuDe.length).setFontWeight('bold').setBackground('#FFF4CC');
    }
    const sach = v => {
      const s = String(v == null ? '' : v).slice(0, 500);
      return /^[=+\-@0]/.test(s) ? "'" + s : s;   // giữ số 0 đầu SĐT, chặn công thức
    };
    sh.appendRow([new Date()].concat(cfg.giaTri(d).map(sach), ['Mới', '', '']));

    if (BAO_EMAIL) {
      try {
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          '[CES] Đăng ký mới – ' + cfg.tab + ': ' + d.name,
          cfg.tieuDe.slice(1, 7).map((t, i) => t + ': ' + (cfg.giaTri(d)[i] || '')).join('\n') +
          '\n\nMở Sheet: ' + ss.getUrl());
      } catch (err) { /* không gửi được email thì vẫn lưu đăng ký */ }
    }
    return traLoi({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Mở link /exec trên trình duyệt để kiểm tra: thấy {"ok":true,"v":1} là đã chạy.
function doGet() { return traLoi({ ok: true, v: 1 }); }

function traLoi(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Chạy hàm này một lần trong trình soạn thảo để cấp quyền và thử ghi một dòng mẫu (xóa dòng đó sau).
function thuGhi() {
  doPost({ postData: { contents: JSON.stringify({ who: 'doi-tac', name: 'Thử', phone: '0900000000', city: 'Hà Nội', org: 'Thử', size: 'Dưới 50', note: 'dòng thử', secs: 10 }) } });
}
