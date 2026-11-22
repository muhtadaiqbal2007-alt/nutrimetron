/*****************************************************************
 * NUTRIMETRON - Code.gs  (Google Apps Script, terikat ke Spreadsheet)
 * Spreadsheet aktif dicari otomatis: SpreadsheetApp.getActiveSpreadsheet()
 *
 * LANGKAH PASANG (ringkas):
 * 1. Buka Google Spreadsheet > Extensions (Ekstensi) > Apps Script
 * 2. Hapus isi Code.gs bawaan, tempel seluruh isi file ini, Save.
 * 3. Pilih fungsi setup() lalu Run (izinkan otorisasi). Ini membuat sheet
 *    'guest' dan 'settings' + memeriksa sheet bahan makanan.
 * 4. Deploy > New deployment > jenis "Web app"
 *      Execute as : Me
 *      Who has access : Anyone
 *    Klik Deploy, salin URL yang berakhiran /exec
 * 5. Tempel URL itu ke index.html pada baris:  const API_URL = '...';
 * 6. Upload index.html ke GitHub (repo publik) > Settings > Pages.
 * 7. Setiap mengubah Code.gs: Deploy > Manage deployments > Edit >
 *    Version: New version > Deploy (URL tetap sama).
 *****************************************************************/

var CFG = {
  ADMIN_USER: 'nutrimetron',
  ADMIN_PASS: '12345678',
  TKPI_SHEETS: ['serealia','umbi_berpati','kacang_kacangan','sayuran','buah','daging_unggas',
                'ikan','telur','susu','lemak_minyak','konfeksioneri','bumbu','minuman_nonalkohol'],
  USDA_SHEETS: ['usda1','usda2','usda3'],
  // Nomor kolom (A=1, B=2, C=3, ...). Sesuaikan bila susunan sheet berbeda.
  TKPI: { kode: 2, nama: 3, energi: 6, protein: 7, lemak: 8, karbo: 9 },
  // USDA: isi proteinN / proteinH dengan nomor kolom jika sheet punya kolom terpisah (0 = tidak ada)
  USDA: { kode: 2, nama: 3, energi: 6, protein: 7, lemak: 8, karbo: 9, proteinN: 0, proteinH: 0 }
};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
  return handle_(p);
}

function handle_(p) {
  try {
    switch (p.action) {
      case 'ping':         return out_({ ok: true, name: ss_().getName() });
      case 'login':        return out_(login_(p));
      case 'foodList':     return out_(foodList_(p));
      case 'getSettings':  return out_({ ok: true, settings: getSettings_() });
      case 'saveSettings': return out_(saveSettings_(p));
      default:             return out_({ ok: false, error: 'Aksi tidak dikenal' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function sheet_(name, headers) {
  var s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    if (headers) s.appendRow(headers);
  }
  return s;
}

function isAdmin_(u, pw) { return u === CFG.ADMIN_USER && pw === CFG.ADMIN_PASS; }

function login_(p) {
  if (p.role === 'admin') {
    if (isAdmin_(String(p.username || '').trim(), String(p.password || ''))) return { ok: true, role: 'admin' };
    return { ok: false, error: 'Username atau password salah' };
  }
  var nama = String(p.nama || '').trim(), email = String(p.email || '').trim();
  if (!nama || !email) return { ok: false, error: 'Nama dan e-mail wajib diisi' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Format e-mail tidak valid' };
  sheet_('guest', ['Waktu', 'Nama', 'Email']).appendRow([new Date(), nama, email]);
  return { ok: true, role: 'user' };
}

function num_(x) {
  var n = parseFloat(String(x).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function foodList_(p) {
  var type = p.type === 'USDA' ? 'USDA' : 'TKPI';
  var allowed = type === 'USDA' ? CFG.USDA_SHEETS : CFG.TKPI_SHEETS;
  if (allowed.indexOf(p.sheet) < 0) return { ok: false, error: 'Sheet tidak diizinkan' };
  var sh = ss_().getSheetByName(p.sheet);
  if (!sh) return { ok: false, error: "Sheet '" + p.sheet + "' tidak ditemukan di spreadsheet" };
  var c = CFG[type];
  var data = sh.getDataRange().getValues();
  var items = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var kode = String(r[c.kode - 1] || '').trim();
    var nama = String(r[c.nama - 1] || '').trim();
    var en = String(r[c.energi - 1]).trim();
    if (!kode || !nama || en === '' || isNaN(parseFloat(en.replace(',', '.')))) continue; // lewati judul/baris kosong
    var it = { kode: kode, nama: nama, energi: num_(r[c.energi - 1]), protein: num_(r[c.protein - 1]),
               lemak: num_(r[c.lemak - 1]), karbo: num_(r[c.karbo - 1]) };
    if (type === 'USDA' && c.proteinN > 0) it.pn = num_(r[c.proteinN - 1]);
    if (type === 'USDA' && c.proteinH > 0) it.ph = num_(r[c.proteinH - 1]);
    items.push(it);
  }
  return { ok: true, items: items };
}

function getSettings_() {
  var sh = ss_().getSheetByName('settings');
  var o = {};
  if (!sh) return o;
  sh.getDataRange().getValues().forEach(function (r) {
    if (r[0] === 'key' || !r[0]) return;
    try { o[r[0]] = JSON.parse(r[1]); } catch (e) { o[r[0]] = r[1]; }
  });
  return o;
}

function saveSettings_(p) {
  if (!isAdmin_(p.username, p.password)) return { ok: false, error: 'Hanya admin yang boleh menyimpan' };
  var sh = sheet_('settings', ['key', 'value']);
  var data = sh.getDataRange().getValues();
  var settings = p.settings || {};
  Object.keys(settings).forEach(function (k) {
    var v = JSON.stringify(settings[k]), row = -1;
    for (var i = 0; i < data.length; i++) if (data[i][0] === k) { row = i + 1; break; }
    if (row > 0) sh.getRange(row, 2).setValue(v);
    else { sh.appendRow([k, v]); data.push([k, v]); }
  });
  return { ok: true };
}

/** Jalankan sekali dari editor Apps Script. */
function setup() {
  sheet_('guest', ['Waktu', 'Nama', 'Email']);
  sheet_('settings', ['key', 'value']);
  var missing = CFG.TKPI_SHEETS.concat(CFG.USDA_SHEETS).filter(function (n) { return !ss_().getSheetByName(n); });
  Logger.log(missing.length ? 'Sheet bahan yang belum ada: ' + missing.join(', ') : 'Semua sheet bahan lengkap.');
}
