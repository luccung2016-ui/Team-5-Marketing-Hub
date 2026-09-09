const SPREADSHEET_ID = '1l07-4LDdOO-nKWfaf8CYEqQV-oIpr4gsTrqLSPjWIn0';
const SHEETS = [
  'Học Tiếng Trung 09.2026',
  'Học Tiếng Nhật 09.2026',
  'Học Tiếng Anh 09.2026',
  'Du Học Nghề Singapore 09.2026',
  'Du Học Úc 09.2026'
];

function doGet(e) {
  try {
    const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_TOKEN');
    const token = e && e.parameter ? e.parameter.token : '';
    if (!expected || token !== expected) return json_({ ok: false, error: 'UNAUTHORIZED' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const posts = [];

    SHEETS.forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return;
      const values = sh.getRange(1, 1, lastRow, 11).getDisplayValues();
      const headers = values[0].map(x => String(x || '').trim());
      const idx = headerMap_(headers);

      values.slice(1).forEach(r => {
        const id = cell_(r, idx.id);
        if (!id) return;
        posts.push({
          id,
          date: normalizeDate_(cell_(r, idx.date)),
          page: cell_(r, idx.page),
          pillar: cell_(r, idx.pillar),
          angle: cell_(r, idx.angle),
          topic: cell_(r, idx.topic),
          format: cell_(r, idx.format),
          audience: cell_(r, idx.audience),
          status: cell_(r, idx.status) || 'Chờ viết',
          contentBrief: cell_(r, idx.contentBrief),
          visualBrief: cell_(r, idx.visualBrief),
          sourceSheet: name
        });
      });
    });

    return json_({
      ok: true,
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetName: ss.getName(),
      syncedAt: new Date().toISOString(),
      count: posts.length,
      posts
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function headerMap_(h) {
  const find = names => {
    for (let i = 0; i < h.length; i++) {
      const v = h[i].toLowerCase();
      if (names.some(n => v === n.toLowerCase())) return i;
    }
    return -1;
  };
  return {
    id: find(['ID','Mã content']),
    date: find(['Ngày đăng','Thứ và Ngày']),
    page: find(['Fanpage']),
    pillar: find(['Pillar','Content Pillar']),
    angle: find(['Angle','Content Angle']),
    topic: find(['Chủ đề','Chủ đề bài viết']),
    format: find(['Định dạng']),
    audience: find(['Nhóm khách hàng']),
    status: find(['Trạng thái']),
    contentBrief: find(['Hướng triển khai content','Hướng triển khai content (áp dụng theo content framework phù hợp)']),
    visualBrief: find(['Hướng triển khai Visual'])
  };
}

function cell_(row, i) { return i >= 0 ? String(row[i] || '').trim() : ''; }

function normalizeDate_(v) {
  if (!v) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  const m = v.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
  if (!m) return v;
  let y = m[3] || '2026';
  if (y.length === 2) y = '20' + y;
  return `${y}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
