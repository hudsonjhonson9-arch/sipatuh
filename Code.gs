/**
 * SIPATUH - Backend Google Apps Script (API murni untuk frontend di Netlify)
 * Inspektorat Kabupaten Sumba Barat
 *
 * Setup:
 * 1. Buka Extensions > Apps Script dari Spreadsheet SIPATUH_DB
 * 2. Paste file ini sebagai Code.gs
 * 3. Jalankan initDatabase() sekali (manual run) untuk membuat header semua sheet
 * 4. Isi ROOT_FOLDER_ID dengan ID folder Drive untuk bukti dukung
 * 5. Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone   <-- WAJIB persis ini, bukan "Anyone with Google account",
 *        kalau tidak, request dari luar (Netlify) diarahkan ke halaman login Google dan
 *        akan selalu gagal dengan error CORS di browser.
 * 6. Setelah deploy, pakai URL yang diakhiri /exec (bukan /dev) di frontend.
 */
const ROOT_FOLDER_ID = '15mn4pggtH6oKW9Wyj6nsQ29Eh1qbInXW'; // folder root utk bukti dukung
const SESSION_TTL_SECONDS = 8 * 60 * 60; // 8 jam
// Set false untuk mematikan login sepenuhnya — semua request dianggap
// sebagai GUEST_SESSION di bawah. Set true lagi kapan saja untuk aktifkan ulang.
const AUTH_ENABLED = true;
const GUEST_SESSION = { id: 'guest', nama: 'Tanpa Login', email: '', role: 'opd', opd_id: '' };

// ============================================================
// SETUP / INIT
// ============================================================

const SHEET_HEADERS = {
  OPD: ['id', 'kode_opd', 'nama_opd', 'created_at'],
  Users: ['id', 'nama', 'email', 'nip', 'role', 'opd_id', 'password_hash', 'created_at'],
  LHP: ['id', 'no_lhp', 'nama_lhp', 'tahun_pemeriksaan', 'entitas', 'created_at'],
  Temuan: ['id', 'lhp_id', 'uraian_temuan', 'nilai_temuan', 'asal_opd', 'jenis_temuan', 'kategori_keuangan', 'uraian_rekomendasi', 'nilai_rekomendasi', 'target_penyelesaian', 'created_at'],
  DetilTemuan: ['id', 'temuan_id', 'nama', 'jumlah', 'keterangan', 'nomor_referensi'],
  Rekomendasi: ['id', 'temuan_id', 'uraian_rekomendasi', 'nilai_rekomendasi', 'target_penyelesaian', 'created_at'],
  TindakLanjut: ['id', 'rekomendasi_id', 'uraian_tindak_lanjut', 'status', 'alasan_tidak_dapat_dtl', 'penjelasan_perubahan', 'nilai_setoran', 'dibuat_oleh', 'created_at'],
  BuktiDukung: ['id', 'tindak_lanjut_id', 'jenis', 'nama_file', 'drive_file_id', 'drive_url', 'tanggal_upload', 'diupload_oleh'],
  Verifikasi: ['id', 'bukti_dukung_id', 'auditor_id', 'status', 'catatan_perbaikan', 'tanggal_verifikasi'],
  RiwayatPerubahan: ['id', 'tabel', 'record_id', 'field_berubah', 'nilai_lama', 'nilai_baru', 'diubah_oleh', 'diubah_pada'],
  Pengaduan: ['id', 'nama_pelapor', 'kontak', 'anonim', 'opd_terkait', 'uraian_pengaduan', 'bukti_file_id', 'bukti_url', 'status', 'catatan_tindak_lanjut', 'ditangani_oleh', 'created_at', 'updated_at'],
  Setoran: ['id', 'opd_id', 'temuan_id', 'detil_temuan_id', 'uraian_temuan', 'nomor_referensi', 'nama_instansi', 'pelaku_kerugian', 'nama_penyetor', 'no_hp', 'jumlah_temuan', 'jumlah_setoran', 'tanggal_setor', 'nama_file', 'drive_file_id', 'drive_url', 'status', 'catatan_verifikasi', 'dibuat_oleh', 'created_at', 'updated_at'],
};

function initDatabase() {
  const ss = SpreadsheetApp.getActive();
  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      var existing = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
      var existingSet = new Set(existing);
      var reordered = headers.filter(function(h) { return existingSet.has(h); });
      var missing = headers.filter(function(h) { return !existingSet.has(h); });
      if (reordered.length !== existing.length || missing.length) {
        sheet.getRange(1, 1, 1, reordered.length).setValues([reordered]);
        if (missing.length) {
          sheet.getRange(1, reordered.length + 1, 1, missing.length).setValues([missing]);
        }
      }
    }
    sheet.setFrozenRows(1);
  });
}

function backfillDetilRef() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('DetilTemuan');
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const idIdx = h.indexOf('id');
  const temuanIdIdx = h.indexOf('temuan_id');
  const refIdx = h.indexOf('nomor_referensi');
  if (refIdx === -1) return 'Kolom nomor_referensi belum ada. Jalankan initDatabase() dulu.';
  var updated = 0;
  var temuanCache = {};
  for (var r = 1; r < data.length; r++) {
    if (!data[r][refIdx]) {
      var temuanId = data[r][temuanIdIdx];
      if (!temuanCache[temuanId]) {
        var t = sheetToObjects('Temuan').find(function(t) { return t.id === temuanId; });
        temuanCache[temuanId] = t ? t.asal_opd : '';
      }
      var ref = generateNomorReferensi(temuanCache[temuanId] || 'XXXX');
      sheet.getRange(r + 1, refIdx + 1).setValue(ref);
      updated++;
    }
  }
  return 'Berhasil mengisi ' + updated + ' nomor referensi.';
}

/**
 * Jalankan SEKALI setelah mengisi/mengedit sheet OPD secara manual.
 * Mengisi kolom id yang masih kosong dengan UUID, dan kolom created_at
 * yang masih kosong dengan waktu sekarang. Aman dijalankan berkali-kali.
 */
function fillMissingOpdIds() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('OPD');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf('id') + 1;
  const createdCol = header.indexOf('created_at') + 1;

  for (let r = 2; r <= data.length; r++) {
    const row = data[r - 1];
    if (!row[idCol - 1]) sheet.getRange(r, idCol).setValue(Utilities.getUuid());
    if (!row[createdCol - 1]) sheet.getRange(r, createdCol).setValue(new Date());
  }
  Logger.log('Selesai. Cek sheet OPD, catat id tiap OPD untuk dipakai di seedInitialUsers().');
}

/**
 * Memperbaiki data yang kolomnya bergeser karena SHEET_HEADERS
 * tidak sesuai urutan kolom di sheet. Jalan SEKALI setelah deploy
 * perubahan appendRowObject dan initDatabase.
 * Bekerja untuk sheet: Setoran, BuktiDukung, dan sheet lain.
 */
function fixAllSheetColumns() {
  var hasil = [];
  Object.keys(SHEET_HEADERS).forEach(function(name) {
    var sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getDataRange().getValues();
    var oldHeaders = data[0];
    var targetHeaders = SHEET_HEADERS[name];
    var colMap = targetHeaders.map(function(h) { return oldHeaders.indexOf(h); });
    if (colMap.indexOf(-1) !== -1) {
      hasil.push(name + ': kolom header kurang, jalankan initDatabase() dulu');
      return;
    }
    if (colMap.every(function(ci, i) { return ci === i; })) {
      hasil.push(name + ': sudah sesuai');
      return;
    }
    var corrected = [targetHeaders];
    for (var r = 1; r < data.length; r++) {
      var newRow = targetHeaders.map(function(h, i) { return data[r][colMap[i]]; });
      corrected.push(newRow);
    }
    sheet.clear();
    sheet.getRange(1, 1, corrected.length, corrected[0].length).setValues(corrected);
    hasil.push(name + ': diperbaiki ' + (data.length - 1) + ' baris');
  });
  SpreadsheetApp.flush();
  return hasil.join('\n');
}

/**
 * Contoh seed user awal. GANTI email/password/opd_id sesuai kebutuhan,
 * lalu jalankan sekali dari editor. Setelah sukses, hapus/komentari
 * fungsi ini supaya password tidak tertinggal di kode.
 */
function seedInitialUsers() {
  Logger.log(registerUser({
    nama: 'Admin Inspektorat',
    email: 'inspektorat@sumbabaratkab.go.id',
    role: 'ppupd',
    password: 'GantiPasswordIni123',
  }));

  Logger.log(registerUser({
    nama: 'Admin DPMPTSP',
    email: 'dpmptsp@sumbabaratkab.go.id',
    role: 'opd',
    opd_id: 'PASTE_ID_OPD_DARI_SHEET',
    password: 'GantiPasswordIni123',
  }));
}

// ============================================================
// HELPERS UMUM (baca/tulis sheet sebagai object)
// ============================================================

function sheetToObjects(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  return data.map(row => Object.fromEntries(header.map((h, i) => [h, row[i]])));
}

function appendRowObject(sheetName, obj) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var header;
  try {
    var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    header = range.getValues()[0];
  } catch(e) {
    header = SHEET_HEADERS[sheetName];
  }
  SHEET_HEADERS[sheetName].forEach(function(h) {
    if (header.indexOf(h) === -1) header.push(h);
  });
  const row = header.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
  return obj;
}

function logRiwayat(tabel, recordId, field, nilaiLama, nilaiBaru, diubahOleh) {
  appendRowObject('RiwayatPerubahan', {
    id: Utilities.getUuid(),
    tabel, record_id: recordId, field_berubah: field,
    nilai_lama: String(nilaiLama), nilai_baru: String(nilaiBaru),
    diubah_oleh: diubahOleh, diubah_pada: new Date(),
  });
}

// ============================================================
// AUTH - password manual + session token via CacheService
// ============================================================

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function registerUser(payload) {
  const users = sheetToObjects('Users');
  if (users.some(u => u.email === payload.email)) {
    return { success: false, error: 'Email sudah terdaftar' };
  }
  const user = {
    id: Utilities.getUuid(),
    nama: payload.nama,
    email: payload.email,
    nip: payload.nip || '',
    role: payload.role,
    opd_id: payload.opd_id || '',
    password_hash: hashPassword(payload.password),
    created_at: new Date(),
  };
  appendRowObject('Users', user);
  delete user.password_hash;
  return { success: true, user };
}

function login(email, password) {
  const users = sheetToObjects('Users');
  const user = users.find(u => u.email === email);
  if (!user || user.password_hash !== hashPassword(password)) {
    return { success: false, error: 'Email atau password salah' };
  }
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put('session_' + token, JSON.stringify({
    id: user.id, nama: user.nama, email: user.email, role: user.role, opd_id: user.opd_id,
  }), SESSION_TTL_SECONDS);
  return {
    success: true, token,
    user: { id: user.id, nama: user.nama, email: user.email, role: user.role, opd_id: user.opd_id },
  };
}

function validateSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session_' + token);
  return raw ? JSON.parse(raw) : null;
}

function requireRole(session, allowedRoles) {
  if (!AUTH_ENABLED) return true;
  if (session && session.role === 'superadmin') return true;
  return session && allowedRoles.includes(session.role);
}

function generateNomorReferensi(asalOpd) {
  var opdList = sheetToObjects('OPD');
  var opd = opdList.find(function(o) { return o.id === asalOpd; });
  var kode = opd ? (opd.kode_opd || opd.nama_opd.slice(0,4).toUpperCase()) : 'XXXX';
  var temuan = sheetToObjects('Temuan').filter(function(t) { return t.asal_opd === asalOpd; });
  var count = temuan.length + 1;
  var padded = ('000' + count).slice(-3);
  var unik = Math.random().toString(36).substring(2,5).toUpperCase();
  return kode + '-' + padded + '-' + unik;
}

function cariTemuanByRef(ref) {
  var detil = sheetToObjects('DetilTemuan').filter(function(d) { return d.nomor_referensi === ref; });
  if (!detil.length) return { found: false, error: 'Referensi tidak ditemukan' };
  var d = detil[0];
  var temuan = sheetToObjects('Temuan').filter(function(t) { return t.id === d.temuan_id; });
  if (!temuan.length) return { found: false, error: 'Temuan tidak ditemukan' };
  var t = temuan[0];
  var opdList = sheetToObjects('OPD');
  var opd = opdList.find(function(o) { return o.id === t.asal_opd || o.kode_opd === t.asal_opd; });
  return {
    found: true,
    temuan: {
      id: t.id,
      uraian_temuan: t.uraian_temuan,
      asal_opd: t.asal_opd,
      nama_opd: opd ? opd.nama_opd : '',
    },
    detil: {
      id: d.id,
      nama: d.nama || '',
      jumlah: d.jumlah || 0,
      nomor_referensi: d.nomor_referensi,
    }
  };
}

// ============================================================
// STATUS TERKINI & PROGRESS (dihitung, tidak disimpan ulang)
// ============================================================

function getStatusTerkini() {
  const rows = sheetToObjects('TindakLanjut');
  const latest = {};
  rows.forEach(r => {
    const created = new Date(r.created_at);
    if (!latest[r.rekomendasi_id] || created > latest[r.rekomendasi_id]._created) {
      latest[r.rekomendasi_id] = { status: r.status, nilai_setoran: r.nilai_setoran, _created: created };
    }
  });
  return latest;
}

function getProgressOpd(session) {
  const opdRows = sheetToObjects('OPD');
  const temuanRows = sheetToObjects('Temuan');
  const rekomRows = sheetToObjects('Rekomendasi');
  const statusTerkini = getStatusTerkini();

  const rekomByTemuan = {};
  rekomRows.forEach(r => {
    if (!rekomByTemuan[r.temuan_id]) rekomByTemuan[r.temuan_id] = [];
    rekomByTemuan[r.temuan_id].push(r);
  });

  const result = {};
  opdRows.forEach(o => result[o.id] = { nama_opd: o.nama_opd, total: 0, selesai: 0 });

  temuanRows.forEach(t => {
    if (!result[t.asal_opd]) return;
    result[t.asal_opd].total++;
    const rekomList = rekomByTemuan[t.id] || [];
    const ownRekomSesuai = t.uraian_rekomendasi && statusTerkini[t.id]?.status === 'sesuai';
    const allRekomSesuai = rekomList.length > 0 && rekomList.every(r => statusTerkini[r.id]?.status === 'sesuai');
    if ((rekomList.length && allRekomSesuai) || (!rekomList.length && ownRekomSesuai)) result[t.asal_opd].selesai++;
  });

  Object.values(result).forEach(r => {
    r.persentase = r.total ? Math.round((r.selesai / r.total) * 1000) / 10 : 0;
  });

  if (session && session.role === 'opd') {
    const own = {};
    if (result[session.opd_id]) own[session.opd_id] = result[session.opd_id];
    return own;
  }
  return result;
}

function getProgressKabupaten() {
  const rekomRows = sheetToObjects('Rekomendasi');
  const statusTerkini = getStatusTerkini();
  const counts = { sesuai: 0, belum_sesuai: 0, belum_ditindaklanjuti: 0, tidak_dapat_ditindaklanjuti: 0 };

  rekomRows.forEach(r => {
    const status = statusTerkini[r.id]?.status || 'belum_ditindaklanjuti';
    counts[status] = (counts[status] || 0) + 1;
  });

  const total = rekomRows.length;
  return {
    total_rekomendasi: total,
    ...counts,
    persentase_kabupaten: total ? Math.round((counts.sesuai / total) * 1000) / 10 : 0,
  };
}

function getListTemuan(filter, session) {
  const lhpRows = sheetToObjects('LHP');
  const temuanRows = sheetToObjects('Temuan');
  const rekomRows = sheetToObjects('Rekomendasi');
  const detilRows = sheetToObjects('DetilTemuan');
  const statusTerkini = getStatusTerkini();
  const detilByTemuan = {};
  detilRows.forEach(d => {
    if (!detilByTemuan[d.temuan_id]) detilByTemuan[d.temuan_id] = [];
    detilByTemuan[d.temuan_id].push(d);
  });

  const lhpById = Object.fromEntries(lhpRows.map(l => [l.id, l]));

  let result = temuanRows.map(t => {
    const lhp = lhpById[t.lhp_id] || {};
    const rekomFromSheet = rekomRows.filter(r => r.temuan_id === t.id).map(r => ({ ...r, status: statusTerkini[r.id]?.status || 'belum_ditindaklanjuti' }));
    const rekomendasi = t.uraian_rekomendasi
      ? [{ id: t.id, temuan_id: t.id, uraian_rekomendasi: t.uraian_rekomendasi, nilai_rekomendasi: t.nilai_rekomendasi || 0, target_penyelesaian: t.target_penyelesaian || '', status: statusTerkini[t.id]?.status || 'belum_ditindaklanjuti' }].concat(rekomFromSheet)
      : rekomFromSheet;
    return {
      ...t,
      no_lhp: lhp.no_lhp, nama_lhp: lhp.nama_lhp, tahun_pemeriksaan: lhp.tahun_pemeriksaan,
      rekomendasi,
      detil: detilByTemuan[t.id] || [],
    };
  });

  const opdIdFilter = (session && session.role === 'opd') ? session.opd_id : filter.opd_id;
  if (opdIdFilter) result = result.filter(t => t.asal_opd === opdIdFilter);
  if (filter.tahun) result = result.filter(t => String(t.tahun_pemeriksaan) === String(filter.tahun));
  if (filter.status) result = result.filter(t => t.rekomendasi.some(r => r.status === filter.status));
  if (filter.jenis) result = result.filter(t => t.jenis_temuan === filter.jenis);
  if (filter.kategori) result = result.filter(t => t.kategori_keuangan === filter.kategori);

  return result;
}

function getRiwayatBukti(session) {
  const opdRows = sheetToObjects('OPD');
  const opdMap = {};
  opdRows.forEach(o => opdMap[o.id] = o.nama_opd);
  opdRows.forEach(o => { if (o.kode_opd) opdMap[o.kode_opd] = o.nama_opd; });

  const users = sheetToObjects('Users');
  const userMap = {};
  users.forEach(u => userMap[u.id] = u.nama);

  const result = [];

  const setoranRows = sheetToObjects('Setoran');
  setoranRows.forEach(s => {
    result.push({
      id: s.id,
      tanggal: s.tanggal_setor || s.created_at,
      opd_id: s.opd_id,
      opd_nama: opdMap[s.opd_id] || '',
      jenis: 'Setoran',
      uraian: s.uraian_temuan,
      nama_file: s.nama_file,
      drive_url: s.drive_url,
      status: s.status,
      catatan: s.catatan_verifikasi,
      dibuat_oleh_nama: userMap[s.dibuat_oleh] || s.dibuat_oleh || 'Publik',
      created_at: s.created_at,
      pelaku_kerugian: s.pelaku_kerugian,
      jumlah_setoran: s.jumlah_setoran,
    });
  });

  const tlRows = sheetToObjects('TindakLanjut');
  const rekomRows = sheetToObjects('Rekomendasi');
  const temuanRows = sheetToObjects('Temuan');
  const buktiRows = sheetToObjects('BuktiDukung');
  const verifRows = sheetToObjects('Verifikasi');

  const rekomById = {};
  rekomRows.forEach(r => rekomById[r.id] = r);
  const temuanById = {};
  temuanRows.forEach(t => temuanById[t.id] = t);

  const latestVerif = {};
  verifRows.forEach(v => {
    if (!latestVerif[v.bukti_dukung_id] || new Date(v.tanggal_verifikasi) > new Date(latestVerif[v.bukti_dukung_id].tanggal_verifikasi)) {
      latestVerif[v.bukti_dukung_id] = v;
    }
  });

  buktiRows.forEach(b => {
    const tl = tlRows.find(t => t.id === b.tindak_lanjut_id);
    if (!tl) return;
    const rekom = rekomById[tl.rekomendasi_id] || {};
    const temuan = temuanById[rekom.temuan_id] || {};
    const verif = latestVerif[b.id];
    result.push({
      tanggal: b.tanggal_upload,
      opd_id: temuan.asal_opd || '',
      opd_nama: opdMap[temuan.asal_opd] || '',
      jenis: 'Bukti Tindak Lanjut',
      uraian: tl.uraian_tindak_lanjut,
      nama_file: b.nama_file,
      drive_url: b.drive_url,
      status: verif ? verif.status : 'belum_diverifikasi',
      catatan: verif ? verif.catatan_perbaikan : '',
      dibuat_oleh_nama: userMap[b.diupload_oleh] || b.diupload_oleh || '-',
    });
  });

  if (session && session.role === 'opd') {
    return result.filter(r => r.opd_id === session.opd_id).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  }
  return result.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

// ============================================================
// CREATE + BATCH IMPORT TEMUAN
// ============================================================

function resolveLhpId(payload) {
  if (payload.lhp_id) return payload.lhp_id;
  if (!payload.nama_lhp) return '';
  const rows = sheetToObjects('LHP');
  const found = rows.find(r => String(r.nama_lhp || '').trim().toLowerCase() === String(payload.nama_lhp).trim().toLowerCase());
  if (found) return found.id;
  const id = Utilities.getUuid();
  appendRowObject('LHP', {
    id: id,
    no_lhp: payload.no_lhp || '',
    nama_lhp: payload.nama_lhp,
    tahun_pemeriksaan: payload.tahun_pemeriksaan || '',
    entitas: payload.entitas || '',
    created_at: new Date(),
  });
  return id;
}

function createTemuan(payload, session) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const temuanId = Utilities.getUuid();
    const lhp_id = resolveLhpId(payload);
    appendRowObject('Temuan', {
      id: temuanId,
      lhp_id,
      uraian_temuan: payload.uraian_temuan,
      nilai_temuan: payload.nilai_temuan || 0,
      asal_opd: payload.asal_opd,
      jenis_temuan: payload.jenis_temuan || 'administratif',
      kategori_keuangan: payload.kategori_keuangan || '',
      uraian_rekomendasi: payload.uraian_rekomendasi || '',
      nilai_rekomendasi: payload.nilai_rekomendasi || 0,
      target_penyelesaian: payload.target_penyelesaian || '',
      created_at: new Date(),
    });

    const refs = [];
    if (payload.detil && payload.detil.length) {
      payload.detil.forEach(d => {
        const ref = generateNomorReferensi(payload.asal_opd);
        refs.push(ref);
        appendRowObject('DetilTemuan', {
          id: Utilities.getUuid(),
          temuan_id: temuanId,
          nama: d.nama || '',
          jumlah: d.jumlah || 0,
          keterangan: d.keterangan || '',
          nomor_referensi: ref,
        });
      });
    }

    return { success: true, id: temuanId, nomor_referensi: refs };
  } finally {
    lock.releaseLock();
  }
}

function batchImportTemuan(payload, session) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lhpId = payload.lhp_id || Utilities.getUuid();
    
    if (!payload.lhp_id) {
      appendRowObject('LHP', {
        id: lhpId,
        no_lhp: payload.no_lhp || '',
        nama_lhp: payload.nama_lhp || 'LKPD 2024',
        tahun_pemeriksaan: payload.tahun || '2024',
        entitas: payload.entitas || '',
        created_at: new Date(),
      });
    }

    (payload.temuan || []).forEach(t => {
      const temuanId = Utilities.getUuid();
      appendRowObject('Temuan', {
        id: temuanId,
        lhp_id: lhpId,
        uraian_temuan: t.uraian_temuan,
        nilai_temuan: t.nilai_temuan || 0,
        asal_opd: t.opd_id || t.asal_opd,
        jenis_temuan: t.jenis_temuan || 'administratif',
        kategori_keuangan: t.kategori_keuangan || '',
        uraian_rekomendasi: t.uraian_rekomendasi || '',
        nilai_rekomendasi: t.nilai_rekomendasi || 0,
        target_penyelesaian: t.target_penyelesaian || '',
        created_at: new Date(),
      });

      (t.detil || []).forEach(d => {
        appendRowObject('DetilTemuan', {
          id: Utilities.getUuid(),
          temuan_id: temuanId,
          nama: d.nama || '',
          jumlah: d.jumlah || 0,
          keterangan: d.keterangan || '',
          nomor_referensi: generateNomorReferensi(t.opd_id || t.asal_opd),
        });
      });
    });

    return { success: true, lhp_id: lhpId, count: payload.temuan.length };
  } finally {
    lock.releaseLock();
  }
}

function clearTemuanByLhp(payload) {
  const noLhp = (payload.no_lhp || '').trim();
  if (!noLhp) return { success: true, deleted: 0 };
  const lhpSheet = SpreadsheetApp.getActive().getSheetByName('LHP');
  const lhpData = lhpSheet.getDataRange().getValues();
  const lhpHead = lhpData[0];
  const noLhpIdx = lhpHead.indexOf('no_lhp');
  const lhpIdIdx = lhpHead.indexOf('id');
  let lhpId = '';
  for (let r = 1; r < lhpData.length; r++) {
    if (String(lhpData[r][noLhpIdx] || '').trim() === noLhp) { lhpId = lhpData[r][lhpIdIdx]; break; }
  }
  if (!lhpId) return { success: true, deleted: 0 };

  const ss = SpreadsheetApp.getActive();
  const temuanSheet = ss.getSheetByName('Temuan');
  const tData = temuanSheet.getDataRange().getValues();
  const tHead = tData[0];
  const lhpIdx = tHead.indexOf('lhp_id');
  const tidIdx = tHead.indexOf('id');
  const temuanIds = [];
  for (let r = tData.length - 1; r >= 1; r--) {
    if (tData[r][lhpIdx] === lhpId) {
      temuanIds.push(tData[r][tidIdx]);
      temuanSheet.deleteRow(r + 1);
    }
  }
  if (!temuanIds.length) return { success: true, deleted: 0 };
  ['DetilTemuan', 'Rekomendasi'].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    const d = sheet.getDataRange().getValues();
    const h = d[0];
    const refIdx = h.indexOf('temuan_id');
    for (let r = d.length - 1; r >= 1; r--) {
      if (temuanIds.includes(d[r][refIdx])) sheet.deleteRow(r + 1);
    }
  });
  return { success: true, deleted: temuanIds.length };
}

// ============================================================
// TINDAK LANJUT
// ============================================================

function verifyRekomendasiBelongsToOpd(rekomendasiId, opdId) {
  const rekomRows = sheetToObjects('Rekomendasi');
  const temuanRows = sheetToObjects('Temuan');
  const rekom = rekomRows.find(r => r.id === rekomendasiId);
  if (!rekom) return false;
  const temuan = temuanRows.find(t => t.id === rekom.temuan_id);
  return !!temuan && temuan.asal_opd === opdId;
}

function submitTindakLanjut(payload, session) {
  if (session.role === 'opd' && !verifyRekomendasiBelongsToOpd(payload.rekomendasi_id, session.opd_id)) {
    return { success: false, error: 'Rekomendasi ini bukan milik OPD Anda' };
  }
  const entry = {
    id: Utilities.getUuid(),
    rekomendasi_id: payload.rekomendasi_id,
    uraian_tindak_lanjut: payload.uraian_tindak_lanjut,
    status: payload.status,
    alasan_tidak_dapat_dtl: payload.alasan_tidak_dapat_dtl || '',
    penjelasan_perubahan: payload.penjelasan_perubahan || '',
    nilai_setoran: payload.nilai_setoran || 0,
    dibuat_oleh: session.id,
    created_at: new Date(),
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    appendRowObject('TindakLanjut', entry);
  } finally {
    lock.releaseLock();
  }
  return { success: true, id: entry.id };
}

// ============================================================
// UPLOAD BUKTI - Drive, struktur folder: ROOT / Tahun / OPD
// ============================================================

function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function uploadBuktiDukung(payload, session) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const tahunFolder = getOrCreateFolder(root, String(payload.tahun));
  const opdFolder = getOrCreateFolder(tahunFolder, payload.opd_kode);

  const blob = Utilities.newBlob(
    Utilities.base64Decode(payload.base64_data),
    payload.mime_type,
    payload.nama_file
  );
  const file = opdFolder.createFile(blob);

  const entry = {
    id: Utilities.getUuid(),
    tindak_lanjut_id: payload.tindak_lanjut_id,
    jenis: payload.mime_type.includes('pdf') ? 'pdf' : 'foto',
    nama_file: payload.nama_file,
    drive_file_id: file.getId(),
    drive_url: file.getUrl(),
    tanggal_upload: new Date(),
    diupload_oleh: session.id,
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    appendRowObject('BuktiDukung', entry);
  } finally {
    lock.releaseLock();
  }
  return { success: true, id: entry.id, drive_url: entry.drive_url };
}

// ============================================================
// SETORAN
// ============================================================

function submitSetoran(payload, session) {
  let fileId = '', fileUrl = '';
  if (payload.base64_data) {
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const folder = getOrCreateFolder(root, 'Setoran');
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64_data), payload.mime_type, payload.nama_file);
    const file = folder.createFile(blob);
    fileId = file.getId();
    fileUrl = file.getUrl();
  }

  const id = Utilities.getUuid();
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    appendRowObject('Setoran', {
      id: id,
      opd_id: payload.opd_id,
      temuan_id: payload.temuan_id || '',
      detil_temuan_id: payload.detil_temuan_id || '',
      uraian_temuan: payload.uraian_temuan,
      nomor_referensi: payload.nomor_referensi || '',
      nama_instansi: payload.nama_instansi,
      pelaku_kerugian: payload.pelaku_kerugian,
      nama_penyetor: payload.nama_penyetor,
      no_hp: payload.no_hp,
      jumlah_temuan: payload.jumlah_temuan || 0,
      jumlah_setoran: payload.jumlah_setoran || 0,
      tanggal_setor: payload.tanggal_setor || '',
      nama_file: payload.nama_file || '',
      drive_file_id: fileId,
      drive_url: fileUrl,
      status: 'menunggu',
      catatan_verifikasi: '',
      dibuat_oleh: session?.id || 'public',
      created_at: now,
      updated_at: now,
    });
  } finally {
    lock.releaseLock();
  }
  return { success: true, id: id, created_at: now.toISOString(), uraian_temuan: payload.uraian_temuan, pelaku_kerugian: payload.pelaku_kerugian, jumlah_setoran: payload.jumlah_setoran || 0, tanggal_setor: payload.tanggal_setor || '', nama_file: payload.nama_file || '', drive_url: fileUrl };
}

function listSetoran(session) {
  const rows = sheetToObjects('Setoran');
  if (session && session.role === 'opd') return rows.filter(r => r.opd_id === session.opd_id);
  return rows;
}

function verifySetoran(payload, session) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Setoran');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  const statusIdx = header.indexOf('status');
  const catatanIdx = header.indexOf('catatan_verifikasi');
  const updatedIdx = header.indexOf('updated_at');

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === payload.id) {
      sheet.getRange(r + 1, statusIdx + 1).setValue(payload.status);
      sheet.getRange(r + 1, catatanIdx + 1).setValue(payload.catatan || '');
      sheet.getRange(r + 1, updatedIdx + 1).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: 'Setoran tidak ditemukan' };
}

function updateFileSetoran(payload, session) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Setoran');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  const fileIdIdx = header.indexOf('drive_file_id');
  const urlIdx = header.indexOf('drive_url');
  const namaIdx = header.indexOf('nama_file');
  const statusIdx = header.indexOf('status');
  const updatedIdx = header.indexOf('updated_at');

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === payload.id) {
      const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
      const folder = getOrCreateFolder(root, 'Setoran');
      const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64_data), payload.mime_type, payload.nama_file);
      const file = folder.createFile(blob);

      sheet.getRange(r + 1, fileIdIdx + 1).setValue(file.getId());
      sheet.getRange(r + 1, urlIdx + 1).setValue(file.getUrl());
      sheet.getRange(r + 1, namaIdx + 1).setValue(payload.nama_file);
      sheet.getRange(r + 1, statusIdx + 1).setValue('menunggu');
      sheet.getRange(r + 1, updatedIdx + 1).setValue(new Date());
      return { success: true, drive_url: file.getUrl() };
    }
  }
  return { success: false, error: 'Setoran tidak ditemukan' };
}

// ============================================================
// VERIFIKASI
// ============================================================

function getListTindakLanjut(session) {
  const tlRows = sheetToObjects('TindakLanjut');
  const buktiRows = sheetToObjects('BuktiDukung');
  const verifRows = sheetToObjects('Verifikasi');
  const rekomRows = sheetToObjects('Rekomendasi');
  const temuanRows = sheetToObjects('Temuan');
  const opdRows = sheetToObjects('OPD');
  var opdMap = Object.fromEntries(opdRows.map(function(o) { return [o.id, o.nama_opd]; }));
  opdRows.forEach(function(o) { if (o.kode_opd) opdMap[o.kode_opd] = o.nama_opd; });

  const buktiByTl = {};
  buktiRows.forEach(b => {
    if (!buktiByTl[b.tindak_lanjut_id]) buktiByTl[b.tindak_lanjut_id] = [];
    buktiByTl[b.tindak_lanjut_id].push(b);
  });

  const verifByBukti = {};
  verifRows.forEach(v => {
    if (!verifByBukti[v.bukti_dukung_id]) verifByBukti[v.bukti_dukung_id] = [];
    verifByBukti[v.bukti_dukung_id].push(v);
  });

  const latestVerif = {};
  Object.entries(verifByBukti).forEach(([buktiId, list]) => {
    latestVerif[buktiId] = list.sort((a, b) => new Date(b.tanggal_verifikasi) - new Date(a.tanggal_verifikasi))[0];
  });

  const rekomById = {};
  rekomRows.forEach(r => rekomById[r.id] = r);

  const temuanById = {};
  temuanRows.forEach(t => temuanById[t.id] = t);

  return tlRows.map(tl => {
    const rekom = rekomById[tl.rekomendasi_id] || {};
    const temuan = temuanById[rekom.temuan_id] || {};
    const buktiList = buktiByTl[tl.id] || [];
    return {
      id: tl.id,
      rekomendasi_id: tl.rekomendasi_id,
      uraian_rekomendasi: rekom.uraian_rekomendasi || '',
      uraian_tindak_lanjut: tl.uraian_tindak_lanjut,
      status_tl: tl.status,
      nilai_setoran: tl.nilai_setoran || 0,
      created_at: tl.created_at,
      opd_id: temuan.asal_opd || '',
      opd_nama: opdMap[temuan.asal_opd] || '',
      temuan_id: rekom.temuan_id || '',
      uraian_temuan: temuan.uraian_temuan || '',
      bukti: buktiList.map(b => ({
        id: b.id,
        nama_file: b.nama_file || '',
        drive_url: String(b.drive_url || '').startsWith('http') ? b.drive_url : '',
        tanggal_upload: b.tanggal_upload,
        status_verifikasi: latestVerif[b.id]?.status || 'belum_diverifikasi',
        catatan_perbaikan: latestVerif[b.id]?.catatan_perbaikan || '',
        tanggal_verifikasi: latestVerif[b.id]?.tanggal_verifikasi || '',
      })),
    };
  });
}

function submitVerifikasi(payload, session) {
  const entry = {
    id: Utilities.getUuid(),
    bukti_dukung_id: payload.bukti_dukung_id,
    auditor_id: session.id,
    status: payload.status,
    catatan_perbaikan: payload.catatan_perbaikan || '',
    tanggal_verifikasi: new Date(),
  };
  appendRowObject('Verifikasi', entry);
  return { success: true, id: entry.id };
}

// ============================================================
// PENGADUAN MASYARAKAT
// ============================================================

function submitPengaduan(payload, session) {
  let buktiFileId = '', buktiUrl = '';

  if (payload.base64_data) {
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const pengaduanFolder = getOrCreateFolder(root, 'Pengaduan');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.base64_data), payload.mime_type, payload.nama_file
    );
    const file = pengaduanFolder.createFile(blob);
    buktiFileId = file.getId();
    buktiUrl = file.getUrl();
  }

  const entry = {
    id: Utilities.getUuid(),
    nama_pelapor: payload.anonim ? '' : (payload.nama_pelapor || ''),
    kontak: payload.anonim ? '' : (payload.kontak || ''),
    anonim: !!payload.anonim,
    opd_terkait: payload.opd_terkait || '',
    uraian_pengaduan: payload.uraian_pengaduan,
    bukti_file_id: buktiFileId,
    bukti_url: buktiUrl,
    status: 'baru',
    catatan_tindak_lanjut: '',
    ditangani_oleh: '',
    created_at: new Date(),
    updated_at: new Date(),
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    appendRowObject('Pengaduan', entry);
  } finally {
    lock.releaseLock();
  }
  return { success: true, id: entry.id };
}

function getListPengaduan(filter) {
  let rows = sheetToObjects('Pengaduan');
  if (filter.status) rows = rows.filter(r => r.status === filter.status);
  return rows.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
}

function updatePengaduanStatus(payload, session) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Pengaduan');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  const statusIdx = header.indexOf('status');
  const catatanIdx = header.indexOf('catatan_tindak_lanjut');
  const ditanganiIdx = header.indexOf('ditangani_oleh');
  const updatedIdx = header.indexOf('updated_at');

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === payload.id) {
      sheet.getRange(r + 1, statusIdx + 1).setValue(payload.status);
      sheet.getRange(r + 1, catatanIdx + 1).setValue(payload.catatan_tindak_lanjut || '');
      sheet.getRange(r + 1, ditanganiIdx + 1).setValue(session.id);
      sheet.getRange(r + 1, updatedIdx + 1).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: 'Pengaduan tidak ditemukan' };
}

// ============================================================
// UPDATE TEMUAN
// ============================================================

function updateTemuan(payload, session) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Temuan');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  const kategoriIdx = header.indexOf('kategori_keuangan');

  if (idIdx < 0) return { success: false, error: 'Kolom id tidak ditemukan' };

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === payload.id) {
      if (payload.kategori_keuangan !== undefined && kategoriIdx >= 0)
        sheet.getRange(r + 1, kategoriIdx + 1).setValue(payload.kategori_keuangan);
      return { success: true };
    }
  }
  return { success: false, error: 'Temuan tidak ditemukan' };
}

// ============================================================
// PENGATURAN — Kelola OPD, Kelola Pengguna, Ganti Password
// ============================================================

function addOpd(payload) {
  const opdRows = sheetToObjects('OPD');
  if (opdRows.some(o => o.kode_opd === payload.kode_opd)) {
    return { success: false, error: 'Kode OPD sudah terdaftar' };
  }
  const entry = {
    id: Utilities.getUuid(),
    kode_opd: payload.kode_opd,
    nama_opd: payload.nama_opd,
    created_at: new Date(),
  };
  appendRowObject('OPD', entry);
  return { success: true, opd: entry };
}

function listUsers() {
  return sheetToObjects('Users').map(({ password_hash, ...rest }) => rest);
}

function deleteUser(userId, session) {
  if (userId === session.id) return { success: false, error: 'Tidak bisa menghapus akun sendiri' };
  const sheet = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][idIdx] === userId) { sheet.deleteRow(r + 1); return { success: true }; }
  }
  return { success: false, error: 'Pengguna tidak ditemukan' };
}

function changePassword(payload, session) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('id');
  const hashIdx = header.indexOf('password_hash');

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === session.id) {
      if (data[r][hashIdx] !== hashPassword(payload.old_password)) {
        return { success: false, error: 'Password lama salah' };
      }
      sheet.getRange(r + 1, hashIdx + 1).setValue(hashPassword(payload.new_password));
      return { success: true };
    }
  }
  return { success: false, error: 'Akun tidak ditemukan' };
}

// ============================================================
// ROUTER - doGet / doPost
// (Backend murni JSON API — tampilan sekarang di-host di Netlify,
//  bukan lagi disajikan lewat HtmlService.)
// ============================================================

const PUBLIC_ACTIONS = ['login', 'submitSetoran', 'listOpd', 'cariTemuanByRef'];

function doGet(e) {
  let result;
  try {
    const action = e.parameter.action;
    const session = validateSession(e.parameter.token) || (!AUTH_ENABLED ? GUEST_SESSION : null);

    if (AUTH_ENABLED && !PUBLIC_ACTIONS.includes(action) && !session) {
      return jsonOutput({ error: 'Unauthorized' });
    }

    switch (action) {
      case 'progressOpd':
        result = getProgressOpd(session); break;
      case 'progressKabupaten':
        result = getProgressKabupaten(); break;
      case 'listTemuan':
        result = getListTemuan({...e.parameter, jenis: e.parameter.jenis, kategori: e.parameter.kategori}, session); break;
      case 'listOpd':
        result = sheetToObjects('OPD'); break;
      case 'listPengaduan':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = getListPengaduan(e.parameter); break;
      case 'listUsers':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = listUsers(); break;
      case 'listSetoran':
        result = listSetoran(session); break;
      case 'listTindakLanjut':
        if (!requireRole(session, ['auditor', 'ppupd', 'superadmin'])) { result = { error: 'Forbidden' }; break; }
        result = getListTindakLanjut(session); break;
      case 'listRiwayatBukti':
        result = getRiwayatBukti(session); break;
      case 'cariTemuanByRef':
        result = cariTemuanByRef(e.parameter.ref); break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonOutput(result);
}

function doPost(e) {
  // Seluruh isi fungsi dibungkus try/catch — termasuk JSON.parse dan
  // validateSession — supaya APAPUN yang terjadi, respons yang balik
  // selalu lewat jsonOutput(). Kalau ada exception yang lolos ke luar
  // fungsi ini, Apps Script menyajikan halaman error HTML generik yang
  // TIDAK punya header CORS, dan browser akan melaporkannya sebagai
  // error CORS (padahal aslinya error 500 biasa) — ini penyebab paling
  // umum kasus "blocked by CORS policy" saat frontend di-host terpisah.
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const session = validateSession(body.token) || (!AUTH_ENABLED ? GUEST_SESSION : null);

    if (AUTH_ENABLED && !PUBLIC_ACTIONS.includes(action) && !session) {
      return jsonOutput({ error: 'Unauthorized' });
    }

    switch (action) {
      case 'login':
        result = login(body.email, body.password); break;
      case 'register':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = registerUser(body.payload); break;
      case 'deleteUser':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = deleteUser(body.payload.userId, session); break;
      case 'submitTindakLanjut':
        if (!requireRole(session, ['opd'])) { result = { error: 'Forbidden' }; break; }
        result = submitTindakLanjut(body.payload, session); break;
      case 'uploadBukti':
        if (!requireRole(session, ['opd'])) { result = { error: 'Forbidden' }; break; }
        result = uploadBuktiDukung(body.payload, session); break;
      case 'verifikasi':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = submitVerifikasi(body.payload, session); break;
      case 'submitPengaduan':
        result = submitPengaduan(body.payload, session); break;
      case 'updatePengaduanStatus':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = updatePengaduanStatus(body.payload, session); break;
      case 'addOpd':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = addOpd(body.payload); break;
      case 'changePassword':
        result = changePassword(body.payload, session); break;
      case 'submitSetoran':
        result = submitSetoran(body.payload, session); break;
      case 'verifySetoran':
        result = verifySetoran(body.payload, session); break;
      case 'updateFileSetoran':
        result = updateFileSetoran(body.payload, session); break;
      case 'createTemuan':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = createTemuan(body.payload, session); break;
      case 'importTemuan':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = batchImportTemuan(body.payload, session); break;
      case 'updateTemuan':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = updateTemuan(body.payload, session); break;
      case 'clearTemuanByLhp':
        if (!requireRole(session, ['auditor', 'ppupd'])) { result = { error: 'Forbidden' }; break; }
        result = clearTemuanByLhp(body.payload); break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: 'Server error: ' + err.message };
  }
  return jsonOutput(result);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}