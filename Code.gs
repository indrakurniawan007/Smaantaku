/**
 * ABSEN SHOLAT BERJAMAAH - Backend Google Apps Script (v2)
 * Tempel seluruh kode ini di Extensions > Apps Script (file Code.gs), lalu
 * Deploy > Manage deployments > (pensil) > Version: New version > Deploy  (URL /exec tetap sama).
 *
 * Perubahan v3:
 *  - PIN per pengawas diganti SATU TOKEN AKSES untuk semua pengguna aplikasi.
 *    Token disimpan di Script Properties; hanya PEMILIK Spreadsheet yang bisa
 *    melihat / mengganti lewat menu "Kelola Token Akses".
 *    Nama pengawas diketik sendiri di aplikasi dan dicatat di kolom Pengawas.
 *
 * Perubahan v2:
 *  - Cache untuk getSiswa & getAbsen (ringan walau banyak HP menarik data bersamaan).
 *  - Pengecekan duplikat tetap benar walau ada kiriman lama (antrean offline berhari-hari).
 *  - Cadangan otomatis harian ke Drive.
 */
const TZ = 'Asia/Makassar';
const SH = {
  SISWA: '1. Data Siswa',
  ABSEN: '2. Absen Mentah',
  BOLOS: '5. Tidak Sholat Hari Ini - MERAH',
  REKAP: '6. Rekap Per Siswa'
};
const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const HIJAU = '#047857';
const MERAH = '#FFCDD2';
const MERAH_KELAS = '#E57373';
const BATAS_HARI = '#455A64';
const BATAS_ABSEN = '#C8E6C9';
const ORANYE = '#FFCC80';

// ---- pengaturan yang boleh Anda ubah ----
const SCAN_ROWS = 5000;
const HARI_LAPORAN = 30;
const JAM_FINAL = 16;
const BATAS_DEFAULT = 2;
const MAKS_BATCH = 60;          // maksimal data per kiriman
const CACHE_ABSEN_DETIK = 30;
const CACHE_SISWA_DETIK = 300;
const PROP_TOKEN = 'TOKEN_AKSES'; // satu token untuk semua pengawas
const TOKEN_MIN = 6;               // panjang minimal token
const PROP_KOTOR = 'LAPORAN_KOTOR'; // penanda: ada absen baru sejak laporan terakhir dibuat

const FORMAT_ABSEN = ['dd/MM/yyyy HH:mm:ss', '@', '@', '@', '@', '@', '@', '@', '0'];
const REKAP_HDR = 6;
const REKAP_AWAL = 7;
const REKAP_KOLOM = 13;

/* ---------- MENU ---------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📿 Menu Absen')
    .addItem('🔑 Kelola Token Akses (khusus pemilik)', 'kelolaToken')
    .addSeparator()
    .addItem('Siapkan Sheet (ensureSheets)', 'ensureSheets')
    .addItem('Perbarui Laporan Tidak Sholat & Rekap Sekarang', 'buatLaporanBolosHarian')
    .addItem('Rapikan Absen Mentah (terbaru di atas)', 'rapikanAbsenMentah')
    .addSeparator()
    .addItem('Aktifkan Pembaruan Otomatis (tiap 10 menit)', 'aktifkanPembaruanOtomatis')
    .addItem('Matikan Pembaruan Otomatis', 'matikanPembaruanOtomatis')
    .addSeparator()
    .addItem('Aktifkan Cadangan Harian (jam 18.00)', 'aktifkanCadangan')
    .addItem('Cadangkan Sekarang', 'buatCadangan')
    .addToUi();
}

/* ---------- SIAPKAN SHEET ---------- */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defs = [
    { name: SH.SISWA, head: ['NIS', 'Nama Lengkap', 'Kelas', 'Wali Kelas'], textCols: [1], auto: true },
    { name: SH.ABSEN, head: ['Timestamp', 'Hari', 'Tanggal', 'NIS', 'Nama', 'Kelas', 'Jenis Sholat', 'Pengawas', 'Minggu Ke'], textCols: [3, 4], auto: true },
    { name: SH.BOLOS, head: [], textCols: [2], auto: false },
    { name: SH.REKAP, head: [], textCols: [1], auto: false }
  ];

  defs.forEach(function (d) {
    let sh = ss.getSheetByName(d.name);
    if (!sh) sh = ss.insertSheet(d.name);
    if (d.auto && sh.getRange(1, 1).getValue() === '') {
      sh.getRange(1, 1, 1, d.head.length).setValues([d.head])
        .setFontWeight('bold').setBackground(HIJAU).setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    d.textCols.forEach(function (c) {
      sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });

  const abs = ss.getSheetByName(SH.ABSEN);
  abs.getRange(2, 1, abs.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  abs.setColumnWidths(1, 9, 130);

  const bolos = ss.getSheetByName(SH.BOLOS);
  if (bolos.getRange(1, 1).getValue() === '') {
    bolos.getRange(1, 1).setValue('Belum ada laporan. Jalankan: 📿 Menu Absen > Perbarui Laporan Tidak Sholat & Rekap Sekarang');
  }
  const rekap = ss.getSheetByName(SH.REKAP);
  if (rekap.getRange(1, 1).getValue() === '') {
    rekap.getRange(1, 1).setValue('Belum ada rekap. Jalankan: 📿 Menu Absen > Perbarui Laporan Tidak Sholat & Rekap Sekarang');
  }
  return 'OK';
}

/* ======================================================================
 *  TOKEN AKSES (satu token untuk semua pengawas, hanya pemilik yang boleh kelola)
 * ====================================================================== */
function bacaToken_() {
  return String(PropertiesService.getScriptProperties().getProperty(PROP_TOKEN) || '');
}
function tokenValid_(t) {
  const asli = bacaToken_();
  t = String(t || '').trim();
  if (!asli || !t || t.length !== asli.length) return false;
  let beda = 0;                       // perbandingan tanpa berhenti di karakter pertama yang beda
  for (let i = 0; i < asli.length; i++) beda |= asli.charCodeAt(i) ^ t.charCodeAt(i);
  return beda === 0;
}
function bersihNama_(s) {
  s = String(s || '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^[=+\-@]+/, '').trim();   // cegah isi sel dibaca sebagai rumus
  return s.slice(0, 40) || 'Tanpa nama';
}
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// Hanya pemilik Spreadsheet. (Editor lain yang diberi akses tidak bisa melihat/mengganti token.)
function pastikanPemilik_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pemilik = ss.getOwner();                 // null bila file ada di Shared Drive
  let saya = '';
  try { saya = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (!pemilik || !saya || pemilik.getEmail().toLowerCase() !== saya.toLowerCase()) {
    throw new Error('Hanya pemilik Spreadsheet yang boleh melihat atau mengganti token.');
  }
}

function tokenAcak_() {
  const huruf = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';   // tanpa O/0/I/1 agar tidak salah baca
  let t = '';
  for (let i = 0; i < 8; i++) t += huruf.charAt(Math.floor(Math.random() * huruf.length));
  return t;
}

function kelolaToken() {
  try { pastikanPemilik_(); }
  catch (e) { SpreadsheetApp.getUi().alert('🔒 ' + e.message); return; }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px">' +
    '<p>Satu token ini dipakai <b>semua pengawas</b> untuk masuk ke aplikasi.<br>' +
    'Mengganti token = semua HP harus memasukkan token baru.</p>' +
    '<p>Token saat ini:</p>' +
    '<input id="t" type="text" autocomplete="off" style="width:100%;font-size:20px;font-family:monospace;padding:8px;box-sizing:border-box" value="' + esc_(bacaToken_()) + '">' +
    '<p style="color:#555">Minimal ' + TOKEN_MIN + ' karakter, tanpa spasi. Kosongkan lalu klik Simpan untuk dibuatkan token acak.</p>' +
    '<p><button onclick="s()" style="padding:8px 16px;font-size:14px">Simpan</button> ' +
    '<button onclick="g()" style="padding:8px 16px;font-size:14px">Buat Acak</button> <span id="h"></span></p></div>' +
    '<script>' +
    'function g(){var h=document.getElementById("h");h.textContent="...";' +
    'google.script.run.withSuccessHandler(function(r){document.getElementById("t").value=r;h.textContent="Klik Simpan untuk memakai token ini."})' +
    '.withFailureHandler(function(e){h.textContent="Gagal: "+e.message}).tokenAcakBaru();}' +
    'function s(){var h=document.getElementById("h");h.textContent="Menyimpan...";' +
    'google.script.run.withSuccessHandler(function(r){document.getElementById("t").value=r;h.textContent="Tersimpan. Bagikan token ini ke pengawas."})' +
    '.withFailureHandler(function(e){h.textContent="Gagal: "+e.message}).simpanToken(document.getElementById("t").value)}' +
    '</script>'
  ).setWidth(420).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'Kelola Token Akses');
}

function tokenAcakBaru() {
  pastikanPemilik_();
  return tokenAcak_();
}

function simpanToken(teks) {
  pastikanPemilik_();
  let t = String(teks || '').trim();
  if (!t) t = tokenAcak_();
  if (t.length < TOKEN_MIN || /\s/.test(t)) throw new Error('Token minimal ' + TOKEN_MIN + ' karakter dan tidak boleh ada spasi');
  PropertiesService.getScriptProperties().setProperty(PROP_TOKEN, t);
  return t;
}

/* ======================================================================
 *  CACHE (dipecah per potongan agar tidak melewati batas 100 KB)
 * ====================================================================== */
function cachePut_(key, str, detik) {
  try {
    const c = CacheService.getScriptCache(), U = 80000;
    const n = Math.ceil(str.length / U), o = {};
    if (n > 20) return;
    for (let i = 0; i < n; i++) o[key + '_' + i] = str.substr(i * U, U);
    o[key + '_n'] = String(n);
    c.putAll(o, detik);
  } catch (e) {}
}
function cacheGet_(key) {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(key + '_n') || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = c.getAll(keys);
    let out = '';
    for (let i = 0; i < n; i++) { const v = got[key + '_' + i]; if (v === undefined || v === null) return null; out += v; }
    return out;
  } catch (e) { return null; }
}
function cacheDel_(key) {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(key + '_n') || 0);
    const keys = [key + '_n'];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    c.removeAll(keys);
  } catch (e) {}
}
function jsonMentah_(str) {
  return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- GET ---------- */
function doGet(e) {
  const q = (e && e.parameter) || {};
  const action = q.action || '';
  try {
    if (!action) return json({ status: 'sukses', pesan: 'API Absen Sholat aktif' });

    if (!tokenValid_(q.token)) {
      return json({ status: 'error', kode: 'TOKEN', pesan: bacaToken_() ? 'Token salah atau sudah diganti' : 'Admin belum membuat token akses' });
    }

    if (action === 'cek') return json({ status: 'sukses' });

    if (action === 'getSiswa') {
      let txt = cacheGet_('siswa');
      if (!txt) {
        const sh = sheet_(SH.SISWA);
        const last = sh.getLastRow();
        let data = [];
        if (last >= 2) {
          data = sh.getRange(2, 1, last - 1, 4).getValues()
            .filter(function (r) { return String(r[0]).trim() !== ''; })
            .map(function (r) {
              return { nis: String(r[0]).trim(), nama: String(r[1]).trim(), kelas: String(r[2]).trim(), waliKelas: String(r[3]).trim() };
            });
        }
        txt = JSON.stringify(data);
        cachePut_('siswa', txt, CACHE_SISWA_DETIK);
      }
      return jsonMentah_('{"status":"sukses","data":' + txt + '}');
    }

    // Absensi hari ini, format ringkas: [nis, jenis, jam, pengawas]
    if (action === 'getAbsen') {
      const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
      let txt = cacheGet_('absen_' + today);
      if (!txt) {
        const data = [];
        bacaBagianAbsen_(sheet_(SH.ABSEN)).forEach(function (r) {
          if (isSep_(r[0]) || String(r[3]).trim() === '') return;
          if (tglStr_(r[2]) !== today) return;
          data.push([
            String(r[3]).trim(),
            String(r[6]).trim(),
            r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'HH:mm') : '',
            String(r[7]).trim()
          ]);
        });
        txt = JSON.stringify(data);
        cachePut_('absen_' + today, txt, CACHE_ABSEN_DETIK);
      }
      return jsonMentah_('{"status":"sukses","tanggal":"' + today + '","data":' + txt + '}');
    }

    return json({ status: 'error', pesan: 'Aksi tidak dikenal' });
  } catch (err) {
    return json({ status: 'error', pesan: String(err) });
  }
}

/* ---------- POST (satu data atau banyak sekaligus) ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let punyaLock = false;
  try {
    const d = JSON.parse(e.postData.contents);
    if (!tokenValid_(d.token)) return json({ status: 'error', kode: 'TOKEN', pesan: 'Token salah atau sudah diganti' });

    const items = Array.isArray(d.batch) ? d.batch : [d];
    if (items.length > MAKS_BATCH) return json({ status: 'error', pesan: 'Kiriman terlalu besar, maksimal ' + MAKS_BATCH });

    lock.waitLock(20000);
    punyaLock = true;
    const hasil = simpanBatch_(items);
    if (Array.isArray(d.batch)) return json({ status: 'sukses', hasil: hasil });
    const h = hasil[0];
    return json({ status: h.status, pesan: h.pesan, jam: h.jam });
  } catch (err) {
    return json({ status: 'error', pesan: String(err) });
  } finally {
    if (punyaLock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

/* ======================================================================
 *  ABSEN MENTAH: terbaru di atas, lama di bawah, dipisah per hari
 * ====================================================================== */
function isSep_(v) { return typeof v === 'string' && v.indexOf('📅') === 0; }
function sepDate_(v) { const m = /(\d{4}-\d{2}-\d{2})/.exec(String(v)); return m ? m[1] : ''; }
function tglStr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).trim(); }
function hariDari_(tgl) {
  return HARI[Number(Utilities.formatDate(Utilities.parseDate(tgl, TZ, 'yyyy-MM-dd'), TZ, 'u')) % 7];
}
function labelHari_(tgl) { return '📅 ' + hariDari_(tgl).toUpperCase() + ', ' + tgl; }
function ensureRows_(sh, sampai) {
  const m = sh.getMaxRows();
  if (sampai > m) sh.insertRowsAfter(m, sampai - m);
}
function tsNum_(v) { return v instanceof Date ? v.getTime() : 0; }

function bacaBagianAbsen_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, SCAN_ROWS);
  const baruFormat = isSep_(sh.getRange(2, 1).getValue());
  const start = baruFormat ? 2 : last - n + 1;
  return sh.getRange(start, 1, n, 8).getValues();
}

function pastikanFormatAbsen_(sh) {
  if (sh.getLastRow() < 2) return;
  if (isSep_(sh.getRange(2, 1).getValue())) return;
  susunUlangAbsen_(sh);
}

function rapikanAbsenMentah() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const n = susunUlangAbsen_(sheet_(SH.ABSEN));
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Selesai. ' + n + ' baris data disusun terbaru di atas.', 'Absen Mentah', 5); } catch (x) {}
  } finally { lock.releaseLock(); }
}

function susunUlangAbsen_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 9).getValues();
  const hari = {};
  let jumlah = 0;
  vals.forEach(function (r) {
    if (isSep_(r[0]) || String(r[3]).trim() === '') return;
    const t = tglStr_(r[2]) || (r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : '');
    if (!t) return;
    (hari[t] = hari[t] || []).push(r);
    jumlah++;
  });
  const tanggal = Object.keys(hari).sort().reverse();
  const out = [], fmts = [], bgs = [], fcs = [], wts = [];
  tanggal.forEach(function (t) {
    out.push([labelHari_(t), '', '', '', '', '', '', '', '']);
    fmts.push(new Array(9).fill('@'));
    bgs.push(new Array(9).fill(BATAS_ABSEN)); fcs.push(new Array(9).fill('#1B5E20')); wts.push(new Array(9).fill('bold'));
    hari[t].sort(function (a, b) { return tsNum_(b[0]) - tsNum_(a[0]); }).forEach(function (r) {
      out.push(r);
      fmts.push(FORMAT_ABSEN);
      bgs.push(new Array(9).fill('#ffffff')); fcs.push(new Array(9).fill('#000000')); wts.push(new Array(9).fill('normal'));
    });
  });
  const tinggi = Math.max(last - 1, out.length);
  ensureRows_(sh, 1 + tinggi);
  sh.getRange(2, 1, tinggi, 9).clear();
  if (out.length) {
    const rg = sh.getRange(2, 1, out.length, 9);
    rg.setNumberFormats(fmts);
    rg.setValues(out);
    rg.setBackgrounds(bgs).setFontColors(fcs).setFontWeights(wts);
  }
  SpreadsheetApp.flush();
  return jumlah;
}

// baca n baris teratas: daftar pembatas hari + kunci NIS|tanggal|jenis yang sudah ada
function pindaiAtas_(sh, n) {
  const vals = n ? sh.getRange(2, 1, n, 7).getValues() : [];
  const blok = [], ada = {};
  vals.forEach(function (r, i) {
    if (isSep_(r[0])) { blok.push({ tanggal: sepDate_(r[0]), row: i + 2 }); return; }
    const k = nisKey_(r[3]);
    if (!k) return;
    ada[k + '|' + tglStr_(r[2]) + '|' + String(r[6]).trim()] = true;
  });
  return { blok: blok, ada: ada };
}

// Menyimpan banyak absen sekaligus (dipanggil di dalam lock).
function simpanBatch_(items) {
  const sh = sheet_(SH.ABSEN);
  pastikanFormatAbsen_(sh);

  // tanggal tertua di kiriman ini (data offline lama bisa jauh di bawah)
  let minT = '9999-99-99';
  items.forEach(function (d) {
    try {
      let ts = d && d.waktu ? new Date(d.waktu) : new Date();
      if (isNaN(ts.getTime())) ts = new Date();
      const t = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd');
      if (t < minT) minT = t;
    } catch (x) {}
  });

  const last = sh.getLastRow();
  const total = last >= 2 ? last - 1 : 0;
  let r = pindaiAtas_(sh, Math.min(total, SCAN_ROWS));
  if (total > SCAN_ROWS) {
    const tua = r.blok.length ? r.blok.map(function (x) { return x.tanggal; }).sort()[0] : '';
    if (!tua || minT <= tua) r = pindaiAtas_(sh, total); // kiriman lama: baca seluruh sheet agar duplikat tetap terdeteksi
  }
  const blok = r.blok, ada = r.ada;

  const hasil = [], baru = [];
  items.forEach(function (d) {
    const id = d && d.id ? d.id : '';
    try {
      if (!d || !d.nis || !d.jenisSholat) { hasil.push({ id: id, status: 'error', pesan: 'Data tidak lengkap' }); return; }
      const jenis = String(d.jenisSholat).trim();
      if (jenis !== 'Dhuhur' && jenis !== 'Ashar') { hasil.push({ id: id, status: 'error', pesan: 'Jenis sholat tidak dikenal' }); return; }
      let ts = d.waktu ? new Date(d.waktu) : new Date();
      if (isNaN(ts.getTime())) ts = new Date();
      const tanggal = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd');
      const nis = String(d.nis).trim();
      const jam = Utilities.formatDate(ts, TZ, 'HH:mm');
      const k = nisKey_(nis) + '|' + tanggal + '|' + jenis;
      if (ada[k]) { hasil.push({ id: id, status: 'duplikat', jam: jam }); return; }
      ada[k] = true;
      baru.push({
        tanggal: tanggal,
        ts: ts,
        row: [
          ts,
          HARI[Number(Utilities.formatDate(ts, TZ, 'u')) % 7],
          tanggal,
          nis,
          d.nama || '',
          d.kelas || '',
          jenis,
          bersihNama_(d.pengawas),
          Number(Utilities.formatDate(ts, TZ, 'w'))
        ]
      });
      hasil.push({ id: id, status: 'sukses', jam: jam });
    } catch (err) {
      hasil.push({ id: id, status: 'error', pesan: String(err) });
    }
  });

  if (baru.length) {
    const grup = {};
    baru.forEach(function (b) { (grup[b.tanggal] = grup[b.tanggal] || []).push(b); });
    Object.keys(grup).sort().reverse().forEach(function (tgl) { sisipHari_(sh, blok, tgl, grup[tgl]); });
    SpreadsheetApp.flush();
    Object.keys(grup).forEach(function (t) { cacheDel_('absen_' + t); });
    try { PropertiesService.getScriptProperties().setProperty(PROP_KOTOR, '1'); } catch (x) {}
  }
  return hasil;
}

function geser_(blok, dariRow, k) {
  blok.forEach(function (x) { if (x.row >= dariRow) x.row += k; });
}
function tulisBaris_(sh, p, data, fmt) {
  const rg = sh.getRange(p, 1, data.length, 9);
  rg.setNumberFormats(fmt);
  rg.setValues(data);
  rg.setBackground('#ffffff').setFontWeight('normal').setFontColor('#000000');
}
function sisipBaris_(sh, p, k) {
  if (p <= sh.getLastRow()) sh.insertRowsBefore(p, k);
  else ensureRows_(sh, p + k - 1);
}

function sisipHari_(sh, blok, tgl, list) {
  list.sort(function (a, b) { return b.ts - a.ts; });
  const k = list.length;
  const data = list.map(function (b) { return b.row; });
  const fmt = data.map(function () { return FORMAT_ABSEN; });
  const b = blok.filter(function (x) { return x.tanggal === tgl; })[0];

  if (b) {
    const p = b.row + 1;
    sisipBaris_(sh, p, k);
    geser_(blok, p, k);
    tulisBaris_(sh, p, data, fmt);
    const bawah = sh.getRange(p + k, 1).getValue();
    if (bawah instanceof Date && bawah.getTime() > list[k - 1].ts.getTime()) {
      const urutRow = blok.map(function (x) { return x.row; }).sort(function (x, y) { return x - y; });
      const berikut = urutRow.filter(function (r) { return r > b.row; })[0];
      const akhir = berikut ? berikut - 1 : sh.getLastRow();
      sh.getRange(p, 1, akhir - p + 1, 9).sort({ column: 1, ascending: false });
    }
    return;
  }

  const lebihLama = blok.slice().sort(function (x, y) { return x.row - y.row; })
    .filter(function (x) { return x.tanggal < tgl; })[0];
  const p = lebihLama ? lebihLama.row : Math.max(2, sh.getLastRow() + 1);
  sisipBaris_(sh, p, k + 1);
  geser_(blok, p, k + 1);
  const sep = sh.getRange(p, 1, 1, 9);
  sep.setNumberFormats([new Array(9).fill('@')]);
  sep.setValues([[labelHari_(tgl), '', '', '', '', '', '', '', '']]);
  sep.setBackground(BATAS_ABSEN).setFontWeight('bold').setFontColor('#1B5E20');
  tulisBaris_(sh, p + 1, data, fmt);
  blok.push({ tanggal: tgl, row: p });
}

/* ======================================================================
 *  LAPORAN TIDAK SHOLAT + REKAP PER SISWA
 * ====================================================================== */
function buatLaporanBolosHarian() {
  const r = bangunSemua_(true);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      r && r.hariLaporan !== undefined ? 'Laporan ' + r.hariLaporan + ' hari dan rekap ' + r.siswaRekap + ' siswa diperbarui.' : 'Sedang diproses, coba lagi sebentar.',
      'Absen Sholat', 5);
  } catch (x) {}
  return r;
}

function perbaruiLaporanOtomatis() {
  const hari = Number(Utilities.formatDate(new Date(), TZ, 'u'));
  const jamNow = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  const dalamJam = !(hari > 4 || jamNow < 11 || jamNow >= 17);
  const p = PropertiesService.getScriptProperties();
  const kotor = p.getProperty(PROP_KOTOR) === '1';
  // di luar jam laporan: hanya bangun ulang bila ada absen terlambat yang baru masuk
  if (!dalamJam && !kotor) return;
  p.deleteProperty(PROP_KOTOR);
  if (!bangunSemua_(false)) p.setProperty(PROP_KOTOR, '1');   // gagal dapat kunci: coba lagi 10 menit lagi
}

function bangunSemua_(aktifkan, sekarang) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return null;
  try {
    sheet_(SH.REKAP); sheet_(SH.BOLOS);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = sekarang || new Date();
    const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const jamNow = Number(Utilities.formatDate(now, TZ, 'H'));
    const p = bacaPengaturan_();
    const siswa = bacaSiswa_();
    const hari = hitungHari_();
    const tanggal = Object.keys(hari).filter(function (t) { return !p.kecuali[t]; }).sort().reverse();
    const hariLaporan = bangunLaporanBolos_(ss, siswa, hari, tanggal, today, jamNow, now);
    const siswaRekap = bangunRekap_(ss, siswa, hari, tanggal, today, jamNow, p);
    if (aktifkan) ss.setActiveSheet(ss.getSheetByName(SH.BOLOS));
    return { hariLaporan: hariLaporan, siswaRekap: siswaRekap };
  } finally {
    lock.releaseLock();
  }
}

function bacaSiswa_() {
  const sh = sheet_(SH.SISWA);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return { nis: String(r[0]).trim(), key: nisKey_(r[0]), nama: String(r[1]).trim(), kelas: String(r[2]).trim() || '(tanpa kelas)', wali: String(r[3]).trim() };
    });
}

function hitungHari_() {
  const sh = sheet_(SH.ABSEN);
  const last = sh.getLastRow();
  const hari = {};
  if (last < 2) return hari;
  sh.getRange(2, 3, last - 1, 5).getValues().forEach(function (r) {
    const k = nisKey_(r[1]);
    if (!k) return;
    const t = tglStr_(r[0]);
    if (!t) return;
    const j = String(r[4]).trim();
    if (j !== 'Dhuhur' && j !== 'Ashar') return;
    const h = hari[t] || (hari[t] = { Dhuhur: {}, Ashar: {} });
    h[j][k] = true;
  });
  return hari;
}

function bacaPengaturan_() {
  const sh = sheet_(SH.REKAP);
  const v2 = sh.getRange(2, 3).getValue();
  let batas = Number(v2);
  if (v2 === '' || !isFinite(batas)) batas = BATAS_DEFAULT;
  const kecuali = {};
  String(tglStr_(sh.getRange(3, 3).getValue())).split(/[,;\s]+/).forEach(function (t) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(t)) kecuali[t] = true;
  });
  return { batas: batas, kecuali: kecuali };
}

function urutAlfa_(a, b) { return String(a).localeCompare(String(b), 'id', { numeric: true }); }
function adaSesi_(h, j) { return Object.keys(h[j]).length > 0; }
function hariFinal_(t, today, jamNow) { return t < today || (t === today && jamNow >= JAM_FINAL); }

function bangunLaporanBolos_(ss, siswa, hari, tanggal, today, jamNow, now) {
  const shB = ss.getSheetByName(SH.BOLOS);
  const tampil = tanggal.slice(0, HARI_LAPORAN);
  const values = [], bgs = [], fcs = [], wts = [];
  const tambah = function (row, bg, fc, wt) {
    values.push(row);
    bgs.push(new Array(8).fill(bg)); fcs.push(new Array(8).fill(fc)); wts.push(new Array(8).fill(wt));
  };

  tampil.forEach(function (t) {
    const h = hari[t];
    const heldD = adaSesi_(h, 'Dhuhur'), heldA = adaSesi_(h, 'Ashar');
    const final = hariFinal_(t, today, jamNow);
    const kata = final ? 'Tidak' : 'Belum';
    const grup = {};
    let n = 0;
    siswa.forEach(function (s) {
      const sd = heldD && !h.Dhuhur[s.key], sa = heldA && !h.Ashar[s.key];
      if (!sd && !sa) return;
      const ket = (sd && sa) ? kata + ' Dhuhur & Ashar' : (sd ? kata + ' Dhuhur' : kata + ' Ashar');
      (grup[s.kelas] = grup[s.kelas] || []).push({
        nis: s.nis, nama: s.nama, kelas: s.kelas, wali: s.wali,
        dc: heldD ? (sd ? '✘' : '✔') : '-', ac: heldA ? (sa ? '✘' : '✔') : '-', ket: ket
      });
      n++;
    });
    const label = '📅 ' + hariDari_(t).toUpperCase() + ', ' + t + '  —  ' + n + ' siswa ' + (final ? 'tidak lengkap' : 'belum lengkap') +
      '  (hadir: Dhuhur ' + Object.keys(h.Dhuhur).length + ', Ashar ' + Object.keys(h.Ashar).length + ')' +
      (final ? '' : '  [SEMENTARA - hari ini belum selesai]');
    tambah([label, '', '', '', '', '', '', ''], BATAS_HARI, '#FFFFFF', 'bold');
    if (!n) { tambah(['', '', '🎉 Semua siswa lengkap pada sesi yang berlangsung', '', '', '', '', ''], '#E8F5E9', '#1B5E20', 'bold'); return; }
    Object.keys(grup).sort(urutAlfa_).forEach(function (k) {
      const list = grup[k].sort(function (x, y) { return urutAlfa_(x.nama, y.nama); });
      tambah(['', '', 'KELAS ' + k + '  —  ' + list.length + ' siswa', '', '', '', '', ''], MERAH_KELAS, '#000000', 'bold');
      list.forEach(function (s, i) { tambah([i + 1, s.nis, s.nama, s.kelas, s.wali, s.dc, s.ac, s.ket], MERAH, '#000000', 'normal'); });
    });
  });

  shB.clear();
  shB.getRange(1, 1).setValue('RIWAYAT SISWA TIDAK SHOLAT BERJAMAAH (hari terbaru di paling atas) - diperbarui ' +
    Utilities.formatDate(now, TZ, 'HH:mm') + ' WITA')
    .setFontWeight('bold').setFontSize(12);
  const head = ['No', 'NIS', 'Nama Lengkap', 'Kelas', 'Wali Kelas', 'Dhuhur', 'Ashar', 'Keterangan'];
  shB.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#B71C1C').setFontColor('#ffffff');
  if (values.length) {
    ensureRows_(shB, 2 + values.length);
    shB.getRange(3, 2, values.length, 1).setNumberFormat('@');
    const rg = shB.getRange(3, 1, values.length, head.length);
    rg.setValues(values).setBackgrounds(bgs).setFontColors(fcs).setFontWeights(wts);
    shB.getRange(3, 6, values.length, 2).setHorizontalAlignment('center');
  } else {
    shB.getRange(3, 1).setValue('Belum ada data absen.');
  }
  shB.setFrozenRows(2);
  shB.setColumnWidths(1, 8, 100);
  shB.setColumnWidth(3, 260);
  shB.setColumnWidth(8, 170);
  return tampil.length;
}

function bangunRekap_(ss, siswa, hari, tanggal, today, jamNow, p) {
  const sh = ss.getSheetByName(SH.REKAP);

  const lama = {};
  const lastOld = sh.getLastRow();
  if (lastOld >= REKAP_AWAL) {
    sh.getRange(REKAP_AWAL, 1, lastOld - REKAP_AWAL + 1, REKAP_KOLOM).getValues().forEach(function (r) {
      const k = nisKey_(r[0]);
      if (k) lama[k] = { status: String(r[8]).trim(), saat: r[9], catatan: r[11], waktu: r[12] };
    });
  }

  const dihitung = tanggal.filter(function (t) { return hariFinal_(t, today, jamNow); });
  const stat = {};
  siswa.forEach(function (s) { stat[s.key] = { d: 0, a: 0, last: '' }; });
  dihitung.forEach(function (t) {
    const h = hari[t];
    const heldD = adaSesi_(h, 'Dhuhur'), heldA = adaSesi_(h, 'Ashar');
    siswa.forEach(function (s) {
      const st = stat[s.key];
      if (heldD && !h.Dhuhur[s.key]) { st.d++; if (t > st.last) st.last = t; }
      if (heldA && !h.Ashar[s.key]) { st.a++; if (t > st.last) st.last = t; }
    });
  });

  const rows = [];
  siswa.forEach(function (s) {
    const st = stat[s.key];
    const total = st.d + st.a;
    if (!total) return;
    const lm = lama[s.key] || {};
    let status = lm.status === 'Sudah ditangani' ? 'Sudah ditangani' : 'Belum ditangani';
    let saat = '', sejak = '', waktu = '';
    if (status === 'Sudah ditangani') {
      saat = (lm.saat === '' || lm.saat === undefined || isNaN(Number(lm.saat))) ? total : Number(lm.saat);
      sejak = Math.max(0, total - saat);
      waktu = lm.waktu || '';
      if (sejak > p.batas) { status = 'Belum ditangani'; saat = ''; sejak = ''; waktu = ''; }
    }
    rows.push([s.nis, s.nama, s.kelas, s.wali, st.d, st.a, total, st.last, status, saat, sejak, lm.catatan || '', waktu]);
  });
  rows.sort(function (a, b) { return (b[6] - a[6]) || urutAlfa_(a[2], b[2]) || urutAlfa_(a[1], b[1]); });

  const hdr = ['NIS', 'Nama Lengkap', 'Kelas', 'Wali Kelas', 'Tidak Dhuhur', 'Tidak Ashar', 'Total Tidak Sholat',
    'Terakhir Tidak Sholat', 'Status Penanganan', 'Total saat Ditangani', 'Bertambah sejak Ditangani', 'Catatan', 'Waktu Ditangani'];
  const ringkas = dihitung.slice(0, 6).map(function (t) {
    return t.slice(5) + ' (D' + Object.keys(hari[t].Dhuhur).length + '/A' + Object.keys(hari[t].Ashar).length + ')';
  }).join(', ');

  sh.getRange(1, 1).setValue('REKAP PER SISWA - berapa kali tidak sholat berjamaah').setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setValue('Batas peringatan (lebih dari):');
  sh.getRange(3, 1).setValue('Tanggal dikecualikan (pisah koma):');
  sh.getRange(4, 1).setValue('Dihitung dari:');
  if (sh.getRange(2, 3).getValue() === '') sh.getRange(2, 3).setValue(BATAS_DEFAULT);
  sh.getRange(3, 3).setNumberFormat('@');
  sh.getRange(2, 3, 2, 1).setBackground('#FFF9C4').setFontWeight('bold');
  sh.getRange(4, 3).setValue(dihitung.length
    ? dihitung.length + ' hari sholat. Terbaru: ' + ringkas + (dihitung.length > 6 ? ', ...' : '') +
      '   (D = jumlah hadir Dhuhur, A = Ashar)'
    : 'belum ada hari yang selesai dihitung');
  sh.getRange(5, 1).setValue('Oranye = tidak sholat LEBIH DARI batas dan belum ditangani. Ubah kolom Status menjadi "Sudah ditangani" agar warnanya kembali.');

  sh.getRange(REKAP_HDR, 1, Math.max(sh.getMaxRows() - REKAP_HDR + 1, 1), REKAP_KOLOM).clear();
  sh.getRange(REKAP_HDR, 1, 1, REKAP_KOLOM).setValues([hdr])
    .setFontWeight('bold').setBackground(HIJAU).setFontColor('#ffffff').setWrap(true);
  ensureRows_(sh, REKAP_AWAL + Math.max(rows.length, 1));
  const tinggi = sh.getMaxRows() - REKAP_AWAL + 1;
  sh.getRange(REKAP_AWAL, 1, tinggi, 1).setNumberFormat('@');
  sh.getRange(REKAP_AWAL, 8, tinggi, 1).setNumberFormat('@');
  sh.getRange(REKAP_AWAL, 13, tinggi, 1).setNumberFormat('@');
  if (rows.length) {
    sh.getRange(REKAP_AWAL, 1, rows.length, REKAP_KOLOM).setValues(rows);
    sh.getRange(REKAP_AWAL, 5, rows.length, 3).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(REKAP_AWAL, 8, rows.length, 1).setHorizontalAlignment('center');
  } else {
    sh.getRange(REKAP_AWAL, 2).setValue('Belum ada siswa yang tercatat tidak sholat.');
  }

  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Belum ditangani', 'Sudah ditangani'], true).setAllowInvalid(false).build();
  sh.getRange(REKAP_AWAL, 9, tinggi, 1).setDataValidation(dv);

  const aturan = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($I' + REKAP_AWAL + '<>"Sudah ditangani",$G' + REKAP_AWAL + '>$C$2)')
    .setBackground(ORANYE).setBold(true)
    .setRanges([sh.getRange(REKAP_AWAL, 1, tinggi, REKAP_KOLOM)])
    .build();
  sh.setConditionalFormatRules([aturan]);

  sh.setFrozenRows(REKAP_HDR);
  [110, 260, 80, 180, 90, 90, 100, 120, 150, 110, 130, 240, 150].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return rows.length;
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() === SH.SISWA) { cacheDel_('siswa'); return; } // data siswa berubah: segarkan cache
    if (sh.getName() !== SH.REKAP) return;
    const c0 = e.range.getColumn(), c1 = c0 + e.range.getNumColumns() - 1;
    if (9 < c0 || 9 > c1) return;
    const r0 = Math.max(e.range.getRow(), REKAP_AWAL), r1 = e.range.getRow() + e.range.getNumRows() - 1;
    for (let r = r0; r <= r1; r++) {
      const status = String(sh.getRange(r, 9).getValue()).trim();
      if (status === 'Sudah ditangani') {
        sh.getRange(r, 10).setValue(Number(sh.getRange(r, 7).getValue()) || 0);
        sh.getRange(r, 11).setValue(0);
        sh.getRange(r, 13).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
      } else {
        sh.getRange(r, 10).clearContent();
        sh.getRange(r, 11).clearContent();
        sh.getRange(r, 13).clearContent();
      }
    }
  } catch (err) {}
}

/* ---------- PEMBARUAN OTOMATIS (TRIGGER) ---------- */
function aktifkanPembaruanOtomatis() {
  matikanPembaruanOtomatis(true);
  ScriptApp.newTrigger('perbaruiLaporanOtomatis').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('Aktif. Tab "5. Tidak Sholat Hari Ini - MERAH" dan "6. Rekap Per Siswa" akan diperbarui tiap 10 menit, Senin-Kamis pukul 11.00-17.00 WITA, dan sekali lagi di luar jam itu bila ada absen terlambat yang baru masuk.');
}
function matikanPembaruanOtomatis(senyap) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'perbaruiLaporanOtomatis') ScriptApp.deleteTrigger(t);
  });
  if (senyap !== true) SpreadsheetApp.getUi().alert('Pembaruan otomatis dimatikan.');
}

/* ---------- CADANGAN OTOMATIS ---------- */
function aktifkanCadangan() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buatCadangan') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buatCadangan').timeBased().everyDays(1).atHour(18).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('Aktif. Salinan Spreadsheet dibuat tiap hari sekitar pukul 18.00 di folder Drive "Cadangan Absen Sholat" (10 salinan terakhir disimpan).');
}
function buatCadangan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const it = DriveApp.getFoldersByName('Cadangan Absen Sholat');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Cadangan Absen Sholat');
  DriveApp.getFileById(ss.getId()).makeCopy('Cadangan ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HHmm'), folder);
  const files = [], fi = folder.getFiles();
  while (fi.hasNext()) files.push(fi.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(10).forEach(function (f) { f.setTrashed(true); });
}

/* ---------- UTIL ---------- */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nama);
  if (!sh) { ensureSheets(); sh = ss.getSheetByName(nama); }
  return sh;
}
function nisKey_(v) {
  return String(v).trim().replace(/^0+/, '');
}
function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}
