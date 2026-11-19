/**
 * ============================================================================
 * BACKEND SISTEM DASHBOARD DATA SEKOLAH (MURID, GURU & AKP)
 * Sokongan: Muat Naik Batch Excel, Carian Kod Sekolah, Keselamatan PIN 101010
 * ============================================================================
 */

const CONFIG = {
  PASSWORD: '101010',
  SHEETS: {
    MURID: 'DATA_MURID',
    GURU: 'DATA_GURU',
    AKP: 'DATA_AKP'
  },
  // PILIHAN TAMBAHAN: Jika anda mahu baca terus daripada 3 fail Google Sheets
  // berasingan yang sedia ada dalam Google Drive, masukkan ID fail di sini.
  // Biarkan kosong '' jika mahu guna 3 tab dalam Google Sheet ini sahaja.
  EXTERNAL_IDS: {
    MURID: '',
    GURU: '',
    AKP: ''
  }
};

/**
 * Jalankan fungsi ini SEKALI sahaja di editor Apps Script untuk membina 3 tab automatik.
 */
function setupSistem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = [CONFIG.SHEETS.MURID, CONFIG.SHEETS.GURU, CONFIG.SHEETS.AKP];

  names.forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    sheet.setFrozenRows(1);
  });

  return 'Persediaan 3 tab pangkalan data (DATA_MURID, DATA_GURU, DATA_AKP) berjaya.';
}

/**
 * Pintu masuk GET (untuk semakan status API sahaja)
 */
function doGet(e) {
  return jsonResponse_({
    status: 'ok',
    app: 'dashboard-data-sekolah',
    message: 'API Dashboard Data Sekolah aktif.',
    version: '3.0'
  });
}

/**
 * Pintu masuk POST utama (mengelakkan ralat CORS dengan text/plain JSON)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ status: 'error', message: 'Tiada payload diterima.' });
    }

    const req = JSON.parse(e.postData.contents);

    // Semakan Kata Laluan Keselamatan
    if (String(req.pin || '') !== CONFIG.PASSWORD) {
      return jsonResponse_({ status: 'error', message: 'Kata laluan tidak sah. Akses ditolak.' });
    }

    const action = req.action;

    if (action === 'verifyPin') {
      return jsonResponse_({ status: 'success', message: 'Pengesahan berjaya.' });
    }

    if (action === 'getMetadata') {
      return jsonResponse_(getMetadataCached_(req.refresh === true));
    }

    if (action === 'rebuildMetadata') {
      return jsonResponse_(getMetadataCached_(true));
    }

    if (action === 'getSchoolData') {
      return jsonResponse_(handleGetSchoolData_(req.kodSekolah, req.refresh === true));
    }

    if (action === 'queryStats') {
      return jsonResponse_(handleQueryStats_(req));
    }

    if (action === 'queryRows') {
      return jsonResponse_(handleQueryRows_(req));
    }

    if (action === 'distinctValues') {
      return jsonResponse_(handleDistinctValues_(req));
    }

    if (action === 'uploadBatch') {
      return jsonResponse_(handleUploadBatch_(req));
    }

    return jsonResponse_({ status: 'error', message: 'Tindakan tidak dikenali: ' + action });
  } catch (err) {
    return jsonResponse_({
      status: 'error',
      message: 'Ralat pelayan: ' + err.message
    });
  }
}

/**
 * Dapatkan objek Sheet sasaran mengikut kategori
 */
function getTargetSheet_(category) {
  const extId = CONFIG.EXTERNAL_IDS[category];
  if (extId && extId.trim() !== '') {
    const extSs = SpreadsheetApp.openById(extId.trim());
    return extSs.getSheets()[0];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONFIG.SHEETS[category];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Normalisasi Kod Sekolah (huruf besar, buang ruang kosong & simbol)
 */
function normalizeCode_(val) {
  return String(val || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Cari indeks kolum secara fleksibel berdasarkan variasi nama header
 */
function findColIndex_(headers, candidates) {
  const normHeaders = headers.map(function (h) {
    return String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  });

  for (let i = 0; i < candidates.length; i++) {
    const target = candidates[i].toUpperCase().replace(/[^A-Z0-9]/g, '');
    const idx = normHeaders.indexOf(target);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Ambil senarai header, jumlah rekod, dan senarai sekolah unik dari 3 jadual
 * Dioptimumkan: Hanya baca lajur Kod & Nama Sekolah supaya pantas untuk data besar
 */
/**
 * ============================================================================
 * CACHE INDEKS (supaya app dibuka dengan pantas)
 * Indeks (header, jumlah rekod, senarai sekolah & lokasi baris setiap sekolah)
 * disimpan dalam Script Properties. Ia dibina semula secara automatik selepas
 * muat naik. Jika data diubah terus dalam Google Sheet, jalankan fungsi
 * binaSemulaIndeks() sekali di editor Apps Script.
 * ============================================================================
 */
const META_PREFIX = 'META_CHUNK_';
const META_CHUNK_SIZE = 8000;

function saveJsonToProps_(prefix, obj) {
  const props = PropertiesService.getScriptProperties();
  clearJsonFromProps_(prefix);
  const text = JSON.stringify(obj);
  const chunks = {};
  let n = 0;
  for (let i = 0; i < text.length; i += META_CHUNK_SIZE) {
    chunks[prefix + n] = text.substring(i, i + META_CHUNK_SIZE);
    n++;
  }
  chunks[prefix + 'COUNT'] = String(n);
  props.setProperties(chunks, false);
}

function loadJsonFromProps_(prefix) {
  const all = PropertiesService.getScriptProperties().getProperties();
  const n = Number(all[prefix + 'COUNT'] || 0);
  if (!n) return null;
  let text = '';
  for (let i = 0; i < n; i++) {
    if (all[prefix + i] === undefined) return null;
    text += all[prefix + i];
  }
  try { return JSON.parse(text); } catch (e) { return null; }
}

function clearJsonFromProps_(prefix) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf(prefix) === 0) props.deleteProperty(k);
  });
}

function invalidateMetadata_() {
  clearJsonFromProps_(META_PREFIX);
  bumpDataVersion_();
}

/** Versi data: berubah setiap kali data dimuat naik / disusun. Cache lama automatik tidak dipakai. */
function getDataVersion_() {
  const p = PropertiesService.getScriptProperties();
  let v = p.getProperty('DATA_VER');
  if (!v) { v = String(Date.now()); p.setProperty('DATA_VER', v); }
  return v;
}

function bumpDataVersion_() {
  PropertiesService.getScriptProperties().setProperty('DATA_VER', String(Date.now()));
}

/**
 * Susun setiap tab mengikut Kod Sekolah jika baris sekolah bertaburan.
 * Selepas disusun, data satu sekolah berada dalam satu blok baris, jadi
 * bacaan data sekolah menjadi SANGAT pantas (1 bacaan sahaja setiap tab).
 */
function sortSheetsIfScattered_(meta) {
  if (!meta || !meta.index) return false;
  const lock = LockService.getScriptLock();
  const hadLock = lock.hasLock();
  if (!hadLock && !lock.tryLock(3000)) return false;
  let sorted = false;
  try {
    ['MURID', 'GURU', 'AKP'].forEach(function (cat) {
      const idx = meta.index[cat];
      if (!idx) return;
      const kods = Object.keys(idx);
      if (!kods.length) return;
      let runs = 0;
      kods.forEach(function (k) { runs += idx[k].length; });
      if (runs <= Math.ceil(kods.length * 1.2)) return; // sudah tersusun
      const headers = meta.headers[cat] || [];
      const kodIdx = findColIndex_(headers, ['KODSEKOLAH', 'KOD_SEKOLAH', 'KOD SEKOLAH', 'KOD']);
      if (kodIdx === -1) return;
      const sheet = getTargetSheet_(cat);
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 3 || lastCol < 1) return;
      sheet.getRange(2, 1, lastRow - 1, lastCol).sort({ column: kodIdx + 1, ascending: true });
      sorted = true;
    });
    if (sorted) {
      SpreadsheetApp.flush();
      bumpDataVersion_();
    }
  } finally {
    if (!hadLock) lock.releaseLock();
  }
  return sorted;
}

function getMetadataCached_(forceRebuild) {
  if (!forceRebuild) {
    const cached = loadJsonFromProps_(META_PREFIX);
    if (cached && cached.status === 'success') {
      cached.cached = true;
      cached.dataVer = getDataVersion_();
      delete cached.index; // indeks baris tidak perlu dihantar ke pelayar
      return cached;
    }
  }
  let fresh = handleGetMetadata_();
  try {
    if (sortSheetsIfScattered_(fresh)) fresh = handleGetMetadata_();
  } catch (e) {}
  fresh.dataVer = getDataVersion_();
  try {
    saveJsonToProps_(META_PREFIX, fresh);
  } catch (e) {
    // Jika indeks terlalu besar untuk disimpan, simpan tanpa indeks baris
    try {
      const lite = JSON.parse(JSON.stringify(fresh));
      delete lite.index;
      saveJsonToProps_(META_PREFIX, lite);
    } catch (e2) {}
  }
  const out = JSON.parse(JSON.stringify(fresh));
  delete out.index;
  out.cached = false;
  return out;
}

/**
 * ============================================================================
 * AUTO-SEGAR (SAMBUNGAN SENTIASA SEDIA)
 * Jalankan pasangAutoSegar() SEKALI di editor Apps Script (klik Run).
 * Ia memasang pemicu yang membina semula indeks setiap jam secara automatik,
 * supaya aplikasi sentiasa dibuka dengan pantas dan data sentiasa terkini
 * walaupun data diubah terus dalam Google Sheet. Tiada butang perlu ditekan.
 * ============================================================================
 */
function pasangAutoSegar() {
  buangAutoSegar();
  ScriptApp.newTrigger('autoSegarIndeks_').timeBased().everyHours(1).create();
  autoSegarIndeks_();
  return 'Auto-segar dipasang: indeks dibina semula setiap 1 jam.';
}

function buangAutoSegar() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoSegarIndeks_') ScriptApp.deleteTrigger(t);
  });
  return 'Auto-segar dibuang.';
}

function autoSegarIndeks_() {
  const started = Date.now();
  const lock = LockService.getScriptLock();
  // Jangan bina semula semasa muat naik sedang berjalan
  if (!lock.tryLock(5000)) return;
  try {
    getMetadataCached_(true);
  } finally {
    lock.releaseLock();
  }
  // Sediakan data SETIAP sekolah lebih awal supaya pilihan sekolah keluar serta-merta
  const meta = loadJsonFromProps_(META_PREFIX);
  if (meta) panaskanCacheSekolah_(meta, started + 5 * 60 * 1000);
}

/** Jalankan secara manual untuk sediakan data semua sekolah serta-merta. */
function sediakanDataSemuaSekolah() {
  const started = Date.now();
  let meta = loadJsonFromProps_(META_PREFIX);
  if (!meta || !meta.index) { getMetadataCached_(true); meta = loadJsonFromProps_(META_PREFIX); }
  const n = panaskanCacheSekolah_(meta, started + 5 * 60 * 1000);
  return 'Data ' + n + ' set sekolah telah disediakan.';
}

/** Jalankan secara manual jika data diubah terus dalam Google Sheet. */
function binaSemulaIndeks() {
  const m = getMetadataCached_(true);
  return 'Indeks dibina semula: ' + m.schools.length + ' sekolah, Murid ' + m.counts.MURID +
    ', Guru ' + m.counts.GURU + ', AKP ' + m.counts.AKP + '.';
}

function handleGetMetadata_() {
  const categories = ['MURID', 'GURU', 'AKP'];
  const headersMap = {};
  const countsMap = {};
  const schoolMap = {};
  const indexMap = {}; // { CAT: { KOD: [[barisMula, bilangan], ...] } }

  categories.forEach(function (cat) {
    indexMap[cat] = {};
    const sheet = getTargetSheet_(cat);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      headersMap[cat] = [];
      countsMap[cat] = 0;
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h || '').trim();
    });
    headersMap[cat] = headers;
    countsMap[cat] = Math.max(0, lastRow - 1);

    if (lastRow < 2) return;

    const kodIdx = findColIndex_(headers, ['KODSEKOLAH', 'KOD_SEKOLAH', 'KOD SEKOLAH', 'KOD']);
    const namaIdx = findColIndex_(headers, ['NAMA_SEKOLAH', 'NAMASEKOLAH', 'NAMA SEKOLAH', 'SEKOLAH']);

    if (kodIdx === -1) return;

    // Baca lajur Kod Sekolah sahaja secara langsung (sangat pantas)
    const kodVals = sheet.getRange(2, kodIdx + 1, lastRow - 1, 1).getValues();
    const namaVals = namaIdx !== -1
      ? sheet.getRange(2, namaIdx + 1, lastRow - 1, 1).getValues()
      : null;

    const catIdx = indexMap[cat];
    for (let r = 0; r < kodVals.length; r++) {
      const kod = normalizeCode_(kodVals[r][0]);
      if (!kod) continue;
      const sheetRow = r + 2;
      const runs = catIdx[kod] || (catIdx[kod] = []);
      const last = runs[runs.length - 1];
      if (last && last[0] + last[1] === sheetRow) last[1]++;
      else runs.push([sheetRow, 1]);
      const nama = namaVals ? String(namaVals[r][0] || '').trim() : '';
      if (!schoolMap[kod] || (nama && schoolMap[kod].nama === '')) {
        schoolMap[kod] = { kod: kod, nama: nama };
      }
    }
  });

  const schools = Object.keys(schoolMap)
    .sort()
    .map(function (k) { return schoolMap[k]; });

  return {
    status: 'success',
    headers: headersMap,
    counts: countsMap,
    schools: schools,
    index: indexMap,
    builtAt: new Date().toISOString()
  };
}

/**
 * ============================================================================
 * CACHE DATA SEKOLAH (supaya pilih sekolah keluar dalam 1-2 saat)
 * Data setiap sekolah dimampatkan dan disimpan dalam CacheService.
 * ============================================================================
 */
const SC_CHUNK = 90000;
const SC_TTL = 21600; // 6 jam (maksimum CacheService)

function scKey_(ver, cat, kod) { return 'S' + ver + '|' + cat + '|' + kod; }

function packJson_(obj) {
  const blob = Utilities.newBlob(JSON.stringify(obj), 'application/json');
  return Utilities.base64Encode(Utilities.gzip(blob).getBytes());
}

function unpackJson_(b64) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString());
}

function scAddEntries_(base, obj, out) {
  const text = packJson_(obj);
  let n = 0;
  for (let i = 0; i < text.length; i += SC_CHUNK) {
    out[base + '#' + n] = text.substring(i, i + SC_CHUNK);
    n++;
  }
  out[base + '#n'] = String(n);
}

function scPutMany_(entries, ttl) {
  const cache = CacheService.getScriptCache();
  const TTL = ttl || SC_TTL;
  let batch = {}, size = 0, count = 0;
  const send = function () {
    if (!count) return;
    try { cache.putAll(batch, TTL); }
    catch (e) {
      Object.keys(batch).forEach(function (k) { try { cache.put(k, batch[k], TTL); } catch (e2) {} });
    }
    batch = {}; size = 0; count = 0;
  };
  Object.keys(entries).forEach(function (k) {
    batch[k] = entries[k];
    size += entries[k].length;
    count++;
    if (count >= 60 || size > 700000) send();
  });
  send();
}

function scGetMany_(bases) {
  const out = {};
  try {
    const cache = CacheService.getScriptCache();
    const ns = cache.getAll(bases.map(function (b) { return b + '#n'; }));
    const keys = [];
    bases.forEach(function (b) {
      const n = Number(ns[b + '#n'] || 0);
      for (let i = 0; i < n; i++) keys.push(b + '#' + i);
    });
    if (!keys.length) return out;
    const vals = cache.getAll(keys);
    bases.forEach(function (b) {
      const n = Number(ns[b + '#n'] || 0);
      if (!n) return;
      let t = '';
      for (let i = 0; i < n; i++) {
        const v = vals[b + '#' + i];
        if (v === undefined || v === null) { t = null; break; }
        t += v;
      }
      if (t) { try { out[b] = unpackJson_(t); } catch (e) {} }
    });
  } catch (e) {}
  return out;
}

/**
 * Ambil semua rekod Murid, Guru, dan AKP bagi satu Kod Sekolah tertentu
 */
function handleGetSchoolData_(kodInput, forceRefresh) {
  const targetKod = normalizeCode_(kodInput);
  if (!targetKod) {
    return { status: 'error', message: 'Sila masukkan Kod Sekolah yang sah.' };
  }

  const categories = ['MURID', 'GURU', 'AKP'];
  const result = {
    status: 'success',
    kodSekolah: targetKod,
    namaSekolah: '',
    data: {},
    fromCache: true
  };

  const ver = getDataVersion_();
  result.dataVer = ver;
  const bases = categories.map(function (c) { return scKey_(ver, c, targetKod); });
  const cached = forceRefresh ? {} : scGetMany_(bases);
  let meta = null;
  const toPut = {};

  categories.forEach(function (cat, i) {
    let ds = cached[bases[i]];
    if (!ds) {
      result.fromCache = false;
      if (!meta) meta = loadJsonFromProps_(META_PREFIX);
      ds = readSchoolCategory_(cat, targetKod, meta);
      scAddEntries_(bases[i], ds, toPut);
    }
    result.data[cat] = { headers: ds.headers || [], rows: ds.rows || [] };
    if (!result.namaSekolah && ds.nama) result.namaSekolah = ds.nama;
  });

  if (Object.keys(toPut).length) {
    try { scPutMany_(toPut); } catch (e) {}
  }
  return result;
}

/** Baca rekod satu sekolah bagi satu kategori terus dari Google Sheet. */
function readSchoolCategory_(cat, targetKod, meta) {
  const sheet = getTargetSheet_(cat);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [], nama: '' };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });

  const kodIdx = findColIndex_(headers, ['KODSEKOLAH', 'KOD_SEKOLAH', 'KOD SEKOLAH', 'KOD']);
  const namaIdx = findColIndex_(headers, ['NAMA_SEKOLAH', 'NAMASEKOLAH', 'NAMA SEKOLAH', 'SEKOLAH']);
  if (kodIdx === -1) return { headers: headers, rows: [], nama: '' };

  const rows = [];
  let nama = '';
  const take = function (row) {
    if (normalizeCode_(row[kodIdx]) !== targetKod) return;
    rows.push(row);
    if (!nama && namaIdx !== -1 && row[namaIdx]) nama = String(row[namaIdx]).trim();
  };

  // LALUAN PANTAS: guna indeks lokasi baris (hanya baca baris sekolah ini)
  const index = meta && meta.index ? meta.index : null;
  if (index && index[cat]) {
    const runs = index[cat][targetKod] || [];
    runs.forEach(function (run) {
      if (run[0] > lastRow) return;
      const n = Math.min(run[1], lastRow - run[0] + 1);
      sheet.getRange(run[0], 1, n, lastCol).getDisplayValues().forEach(take);
    });
    return { headers: headers, rows: rows, nama: nama };
  }

  // LALUAN BIASA: baca lajur Kod Sekolah sahaja, kemudian baris yang sepadan
  const kodCol = sheet.getRange(2, kodIdx + 1, lastRow - 1, 1).getValues();
  const runsFound = [];
  for (let r = 0; r < kodCol.length; r++) {
    if (normalizeCode_(kodCol[r][0]) !== targetKod) continue;
    const sr = r + 2;
    const lr = runsFound[runsFound.length - 1];
    if (lr && lr[0] + lr[1] === sr) lr[1]++;
    else runsFound.push([sr, 1]);
  }
  runsFound.forEach(function (run) {
    sheet.getRange(run[0], 1, run[1], lastCol).getDisplayValues().forEach(take);
  });
  return { headers: headers, rows: rows, nama: nama };
}

/**
 * Sediakan cache data untuk SEMUA sekolah sekali gus (dipanggil oleh pemicu auto-segar).
 * Jika tab sudah tersusun ikut Kod Sekolah, seluruh tab dibaca berperingkat sekali sahaja.
 */
function panaskanCacheSekolah_(meta, deadlineMs) {
  if (!meta || !meta.index) return 0;
  const ver = getDataVersion_();
  let done = 0;

  ['MURID', 'GURU', 'AKP'].forEach(function (cat) {
    if (Date.now() > deadlineMs) return;
    const idx = meta.index[cat];
    const headers = meta.headers[cat] || [];
    if (!idx || !headers.length) return;
    const kodIdx = findColIndex_(headers, ['KODSEKOLAH', 'KOD_SEKOLAH', 'KOD SEKOLAH', 'KOD']);
    const namaIdx = findColIndex_(headers, ['NAMA_SEKOLAH', 'NAMASEKOLAH', 'NAMA SEKOLAH', 'SEKOLAH']);
    if (kodIdx === -1) return;
    const kods = Object.keys(idx);
    const contiguous = kods.every(function (k) { return idx[k].length === 1; });

    let entries = {}, entryCount = 0;
    const flushEntries = function (force) {
      if (entryCount && (force || entryCount >= 30)) {
        scPutMany_(entries);
        entries = {}; entryCount = 0;
      }
    };
    const addSchool = function (kod, rows) {
      let nama = '';
      if (namaIdx !== -1) {
        for (let i = 0; i < rows.length && !nama; i++) nama = String(rows[i][namaIdx] || '').trim();
      }
      scAddEntries_(scKey_(ver, cat, kod), { headers: headers, rows: rows, nama: nama }, entries);
      entryCount++;
      done++;
      flushEntries(false);
    };

    if (contiguous) {
      const sheet = getTargetSheet_(cat);
      const lastRow = sheet.getLastRow();
      const lastCol = headers.length;
      const STEP = 4000;
      let curKod = null, curRows = [], stopped = false;
      for (let r = 2; r <= lastRow; r += STEP) {
        if (Date.now() > deadlineMs) { stopped = true; break; }
        const n = Math.min(STEP, lastRow - r + 1);
        const vals = sheet.getRange(r, 1, n, lastCol).getDisplayValues();
        for (let i = 0; i < vals.length; i++) {
          const k = normalizeCode_(vals[i][kodIdx]);
          if (!k) continue;
          if (k !== curKod) {
            if (curKod) addSchool(curKod, curRows);
            curKod = k; curRows = [];
          }
          curRows.push(vals[i]);
        }
      }
      if (!stopped && curKod) addSchool(curKod, curRows);
    } else {
      kods.forEach(function (k) {
        if (Date.now() > deadlineMs) return;
        const ds = readSchoolCategory_(cat, k, meta);
        addSchool(k, ds.rows);
      });
    }
    flushEntries(true);
  });
  return done;
}

/**
 * Muat naik data Excel secara berperingkat (Batch Upload)
 */
function handleUploadBatch_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const category = String(req.category || '').toUpperCase();
    if (!CONFIG.SHEETS[category]) {
      return { status: 'error', message: 'Kategori tidak sah: ' + category };
    }

    const mode = req.mode === 'append' ? 'append' : 'replace';
    const batchIndex = Number(req.batchIndex || 0);
    const headers = Array.isArray(req.headers) ? req.headers : [];
    const rows = Array.isArray(req.rows) ? req.rows : [];

    const sheet = getTargetSheet_(category);

    // Data berubah: indeks lama tidak sah lagi (akan dibina semula selepas muat naik)
    invalidateMetadata_();

    // Batch pertama: tetapkan header atau kosongkan sheet jika mod Replace
    if (batchIndex === 0) {
      if (mode === 'replace') {
        sheet.clearContents();
        if (headers.length > 0) {
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.setFrozenRows(1);
        }
      } else if (sheet.getLastRow() === 0 && headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
      }
    }

    if (rows.length > 0) {
      const colCount = headers.length > 0 ? headers.length : sheet.getLastColumn();
      const normalizedRows = rows.map(function (r) {
        const out = new Array(colCount);
        for (let c = 0; c < colCount; c++) {
          out[c] = (r[c] === undefined || r[c] === null) ? '' : String(r[c]);
        }
        return out;
      });

      const startRow = sheet.getLastRow() + 1;
      const targetRange = sheet.getRange(startRow, 1, normalizedRows.length, colCount);
      // Format sebagai teks '@' supaya No KP & No Telefon kekalkan angka sifar di hadapan
      targetRange.setNumberFormat('@');
      targetRange.setValues(normalizedRows);
    }

    return {
      status: 'success',
      category: category,
      batchIndex: batchIndex,
      rowsWritten: rows.length,
      totalRowsNow: Math.max(0, sheet.getLastRow() - 1)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper pulangan JSON
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * ============================================================================
 * CARIAN MELUAS / STATISTIK SELURUH NEGARA
 * req.categories: ['MURID','GURU','AKP'] (pilih mana-mana)
 * req.conditions: [{ col, value, values, mode: 'in'|'contains'|'equals'|'notempty' }] (semua mesti padan)
 * req.groupBy: nama kolum (pilihan) untuk pecahan, cth. NEGERI
 * ============================================================================
 */
/**
 * ============================================================================
 * PEMBACA KOLUM PANTAS (dikongsi oleh Kira Jumlah, Pilih Nilai & Papar Senarai)
 * - Kolum yang bersebelahan dibaca dalam SATU panggilan (bukan satu demi satu).
 * - Setiap kolum disimpan (dimampatkan) dalam cache 15 minit. Kunci cache
 *   mengandungi versi data + bilangan baris/kolum, jadi bila data dimuat naik
 *   atau baris ditambah, data lama tidak akan dipakai (keputusan kekal tepat).
 * Pulangkan { indeksKolum: [nilai seperti dipaparkan dalam Sheet, ...] }
 * ============================================================================
 */
const COL_TTL = 900;          // 15 minit
const COL_MAX_CHUNKS = 12;    // kolum terlalu besar tidak dicache

function readColsDisp_(cat, sheet, idxList, lastRow, lastCol) {
  const n = lastRow - 1;
  const out = {};
  const uniq = (idxList || []).filter(function (i, k, a) { return i >= 0 && a.indexOf(i) === k; });
  if (!uniq.length || n < 1) return out;
  const ver = getDataVersion_();
  const baseOf = function (i) { return 'C' + ver + '|' + cat + '|' + lastRow + 'x' + lastCol + '|' + i; };

  const hits = scGetMany_(uniq.map(baseOf));
  const missing = [];
  uniq.forEach(function (i) {
    const h = hits[baseOf(i)];
    if (Array.isArray(h) && h.length === n) out[i] = h; else missing.push(i);
  });
  if (!missing.length) return out;

  // Kumpul kolum berdekatan (jarak <= 3) supaya dibaca sekali gus
  missing.sort(function (a, b) { return a - b; });
  const runs = [];
  let s0 = missing[0], e0 = missing[0];
  for (let k = 1; k < missing.length; k++) {
    if (missing[k] - e0 <= 3) e0 = missing[k];
    else { runs.push([s0, e0]); s0 = e0 = missing[k]; }
  }
  runs.push([s0, e0]);

  const want = {};
  missing.forEach(function (i) { want[i] = true; });
  const entries = {};
  runs.forEach(function (run) {
    const st = run[0], en = run[1];
    const block = sheet.getRange(2, st + 1, n, en - st + 1).getDisplayValues();
    for (let c = st; c <= en; c++) {
      if (!want[c]) continue;
      const j = c - st;
      const col = new Array(n);
      for (let r = 0; r < n; r++) col[r] = block[r][j];
      out[c] = col;
      try {
        const tmp = {};
        scAddEntries_(baseOf(c), col, tmp);
        if (Number(tmp[baseOf(c) + '#n']) <= COL_MAX_CHUNKS) Object.keys(tmp).forEach(function (k) { entries[k] = tmp[k]; });
      } catch (e) {}
    }
  });
  if (Object.keys(entries).length) { try { scPutMany_(entries, COL_TTL); } catch (e) {} }
  return out;
}

function handleQueryStats_(req) {
  const cats = (req.categories || []).filter(function (c) { return ['MURID', 'GURU', 'AKP'].indexOf(c) !== -1; });
  if (!cats.length) return { status: 'error', message: 'Sila pilih sekurang-kurangnya satu data.' };
  const conds = cleanConds_(req.conditions);
  const groupBy = String(req.groupBy || '').trim();

  const cache = CacheService.getScriptCache();
  const shape = cats.map(function (c) { const sh = getTargetSheet_(c); return sh.getLastRow() + 'x' + sh.getLastColumn(); }).join(',');
  const ck = 'Q3|' + getDataVersion_() + '|' + shape + '|' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([cats, conds, groupBy])));
  const hit = cache.get(ck);
  if (hit && !req.refresh) { try { return JSON.parse(hit); } catch (e) {} }

  const norm = function (v) { return String(v == null ? '' : v).trim().toUpperCase(); };
  const results = {};
  cats.forEach(function (cat) {
    const sheet = getTargetSheet_(cat);
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    const res = { total: Math.max(0, lastRow - 1), count: 0, groups: {}, missing: [] };
    results[cat] = res;
    if (lastRow < 2 || lastCol < 1) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
    const colOf = function (name) {
      let i = headers.indexOf(name);
      if (i === -1) { const n = norm(name); i = headers.findIndex(function (h) { return norm(h) === n; }); }
      return i;
    };
    const condIdx = [];
    conds.forEach(function (c) {
      const i = colOf(c.col);
      if (i === -1) res.missing.push(c.col); else condIdx.push(i);
    });
    if (res.missing.length) { res.count = null; return; }
    const gi = groupBy ? colOf(groupBy) : -1;
    if (groupBy && gi === -1) res.missing.push(groupBy);
    const colsData = readColsDisp_(cat, sheet, condIdx.concat(gi >= 0 ? [gi] : []), lastRow, lastCol);
    const condCols = conds.map(function (c, k) { return { vals: colsData[condIdx[k]], test: makeCondTest_(c) }; });
    const gVals = gi >= 0 ? colsData[gi] : null;
    const gLabel = {};
    for (let r = 0; r < lastRow - 1; r++) {
      let ok = true;
      for (let k = 0; k < condCols.length && ok; k++) {
        ok = condCols[k].test(normQ_(condCols[k].vals[r]));
      }
      if (!ok) continue;
      res.count++;
      if (gVals) {
        // Gabung nilai yang sama walaupun beza huruf besar/kecil atau ruang
        const nk = normQ_(gVals[r]);
        if (!gLabel[nk]) gLabel[nk] = labelQ_(gVals[r]) || '(Kosong)';
        const gk = gLabel[nk];
        res.groups[gk] = (res.groups[gk] || 0) + 1;
      }
    }
  });
  const out = { status: 'success', results: results, builtAt: new Date().toISOString() };
  try { const s = JSON.stringify(out); if (s.length < 95000) cache.put(ck, s, COL_TTL); } catch (e) {}
  return out;
}



/**
 * ============================================================================
 * ENJIN SYARAT CARIAN MELUAS (padanan TEPAT)
 * - Nilai dibanding seperti yang DIPAPARKAN dalam Sheet (getDisplayValues),
 *   jadi tarikh, nombor & sifar di depan kekal sama seperti yang dilihat.
 * - Ruang berlebihan / ruang tak nampak dibuang, huruf besar/kecil diabaikan.
 * mode: 'in' (pilih & tick nilai seperti Excel) | 'contains' | 'equals' | 'notempty'
 * ============================================================================
 */
function normQ_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u00A0\u2007\u202F\u200B\u200C\u200D\uFEFF]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function labelQ_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u00A0\u2007\u202F\u200B\u200C\u200D\uFEFF]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanConds_(list) {
  return (list || []).filter(function (c) {
    if (!c || !c.col) return false;
    const mode = c.mode || 'contains';
    if (mode === 'in') return Array.isArray(c.values) && c.values.length > 0;
    if (mode === 'notempty') return true;
    return normQ_(c.value) !== '';
  }).slice(0, 8).map(function (c) {
    const mode = c.mode || 'contains';
    return {
      col: String(c.col),
      mode: mode,
      value: mode === 'in' ? '' : String(c.value == null ? '' : c.value),
      values: mode === 'in' ? c.values.slice(0, 5000).map(function (x) { return String(x == null ? '' : x); }) : []
    };
  });
}

function makeCondTest_(c) {
  const mode = c.mode || 'contains';
  if (mode === 'notempty') return function (v) { return v !== ''; };
  if (mode === 'in') {
    const set = {};
    (c.values || []).forEach(function (x) { set[normQ_(x)] = true; });
    return function (v) { return Object.prototype.hasOwnProperty.call(set, v); };
  }
  const want = normQ_(c.value);
  if (mode === 'equals') return function (v) { return v === want; };
  return function (v) { return v.indexOf(want) !== -1; };
}

function sheetHeadersQ_(sheet, lastCol) {
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
}

function colOfQ_(headers, name) {
  let i = headers.indexOf(name);
  if (i === -1) { const n = normQ_(name); i = headers.findIndex(function (h) { return normQ_(h) === n; }); }
  return i;
}

/**
 * ============================================================================
 * CARIAN MELUAS — SENARAI NILAI UNIK SATU KOLUM (untuk menu tick gaya Excel)
 * req.categories, req.col, req.conditions (syarat lain), req.search, req.limit
 * ============================================================================
 */
function handleDistinctValues_(req) {
  const cats = (req.categories || []).filter(function (c) { return ['MURID', 'GURU', 'AKP'].indexOf(c) !== -1; });
  if (!cats.length) return { status: 'error', message: 'Sila pilih sekurang-kurangnya satu data.' };
  const col = String(req.col || '').trim();
  if (!col) return { status: 'error', message: 'Sila pilih kolum dahulu.' };
  const conds = cleanConds_(req.conditions);
  const search = normQ_(req.search);
  const limit = Math.max(50, Math.min(Number(req.limit) || 3000, 5000));

  const cache = CacheService.getScriptCache();
  const shape = cats.map(function (c) { const sh = getTargetSheet_(c); return sh.getLastRow() + 'x' + sh.getLastColumn(); }).join(',');
  const ck = 'D2|' + getDataVersion_() + '|' + shape + '|' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([cats, col, conds, search, limit])));
  const hit = cache.get(ck);
  if (hit && !req.refresh) { try { return JSON.parse(hit); } catch (e) {} }

  const map = {};
  const missingIn = [];
  let matchedRows = 0;
  cats.forEach(function (cat) {
    const sheet = getTargetSheet_(cat);
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    const headers = sheetHeadersQ_(sheet, lastCol);
    const vi = colOfQ_(headers, col);
    if (vi === -1) { missingIn.push(cat); return; }
    const n = lastRow - 1;
    const condIdx = conds.map(function (c) { return colOfQ_(headers, c.col); });
    if (condIdx.indexOf(-1) !== -1) return; // data ini tiada kolum syarat lain -> tiada rekod sepadan
    const colsData = readColsDisp_(cat, sheet, condIdx.concat([vi]), lastRow, lastCol);
    const condCols = conds.map(function (c, k) { return { vals: colsData[condIdx[k]], test: makeCondTest_(c) }; });
    const vals = colsData[vi];
    for (let r = 0; r < n; r++) {
      let ok = true;
      for (let k = 0; k < condCols.length && ok; k++) ok = condCols[k].test(normQ_(condCols[k].vals[r]));
      if (!ok) continue;
      const key = normQ_(vals[r]);
      if (search && key.indexOf(search) === -1) continue;
      matchedRows++;
      if (map[key]) map[key].n++;
      else map[key] = { v: labelQ_(vals[r]), n: 1 };
    }
  });

  let list = Object.keys(map).map(function (k) { return [map[k].v, map[k].n]; });
  const distinct = list.length;
  const truncated = distinct > limit;
  if (truncated) list.sort(function (a, b) { return b[1] - a[1]; }), list = list.slice(0, limit);

  const out = { status: 'success', col: col, values: list, distinct: distinct, rows: matchedRows, truncated: truncated, missingIn: missingIn };
  try { const s = JSON.stringify(out); if (s.length < 95000) cache.put(ck, s, COL_TTL); } catch (e) {}
  return out;
}

/**
 * ============================================================================
 * CARIAN MELUAS — SENARAI REKOD (untuk paparan jadual, cetak, Excel & PDF)
 * req.category: 'MURID' | 'GURU' | 'AKP'
 * req.conditions: sama seperti queryStats
 * req.columns: senarai kolum untuk dipaparkan (kosong = semua kolum)
 * req.limit: had rekod dipulangkan (lalai 5000, maksimum 60000)
 * req.sortCol / req.sortAsc: (pilihan) susun di pelayan sebelum dihadkan
 * ============================================================================
 */
function handleQueryRows_(req) {
  const cat = String(req.category || '');
  if (['MURID', 'GURU', 'AKP'].indexOf(cat) === -1) return { status: 'error', message: 'Kategori tidak sah.' };
  const conds = cleanConds_(req.conditions);
  const limit = Math.max(1, Math.min(Number(req.limit) || 5000, 60000));

  const sheet = getTargetSheet_(cat);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { status: 'success', category: cat, headers: [], rows: [], count: 0, total: 0, truncated: false };
  }
  const nRows = lastRow - 1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
  const norm = function (v) { return String(v == null ? '' : v).trim().toUpperCase(); };
  const colOf = function (name) {
    let i = headers.indexOf(name);
    if (i === -1) { const n = norm(name); i = headers.findIndex(function (h) { return norm(h) === n; }); }
    return i;
  };

  // Kolum paparan
  let cols = (req.columns || []).map(function (c) { return colOf(String(c)); })
    .filter(function (i, k, arr) { return i !== -1 && arr.indexOf(i) === k; }).slice(0, 60);
  if (!cols.length) cols = headers.map(function (h, i) { return h ? i : -1; }).filter(function (i) { return i !== -1; });

  const missing = [];
  const condIdx = conds.map(function (c) { const i = colOf(c.col); if (i === -1) missing.push(c.col); return i; });
  if (missing.length) return { status: 'error', message: 'Kolum tiada dalam data ' + cat + ': ' + missing.join(', ') };
  const sortIdx = req.sortCol ? colOf(String(req.sortCol)) : -1;

  // Keputusan penuh dicache 15 minit: tekan "Papar Senarai" kali kedua = serta-merta
  const rk = 'R2|' + getDataVersion_() + '|' + cat + '|' + lastRow + 'x' + lastCol + '|' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([conds, cols, limit, sortIdx, req.sortAsc !== false])));
  if (!req.refresh) {
    const hit = scGetMany_([rk])[rk];
    if (hit && hit.status === 'success') { hit.cached = true; return hit; }
  }

  // Baca kolum syarat dahulu, kemudian kolum paparan (dari cache jika ada)
  const colsData = readColsDisp_(cat, sheet, condIdx.concat(cols).concat(sortIdx >= 0 ? [sortIdx] : []), lastRow, lastCol);
  const dispAt = function (r, i) { return colsData[i][r]; };

  const condObjs = conds.map(function (c, k) { return { i: condIdx[k], test: makeCondTest_(c) }; });
  const matched = [];
  for (let r = 0; r < nRows; r++) {
    let ok = true;
    for (let k = 0; k < condObjs.length && ok; k++) {
      ok = condObjs[k].test(normQ_(dispAt(r, condObjs[k].i)));
    }
    if (ok) matched.push(r);
  }

  if (sortIdx >= 0) {
    const asc = req.sortAsc !== false;
    matched.sort(function (a, b) {
      const x = String(dispAt(a, sortIdx) || ''), y = String(dispAt(b, sortIdx) || '');
      if (!x && y) return 1;
      if (x && !y) return -1;
      const nx = Number(x), ny = Number(y);
      const c = (!isNaN(nx) && !isNaN(ny) && x !== '' && y !== '') ? nx - ny : x.localeCompare(y);
      return asc ? c : -c;
    });
  }

  const take = matched.slice(0, limit);
  const rows = take.map(function (r) { return cols.map(function (i) { return dispAt(r, i); }); });
  const out = {
    status: 'success',
    category: cat,
    headers: cols.map(function (i) { return headers[i]; }),
    rows: rows,
    count: matched.length,
    total: nRows,
    truncated: matched.length > take.length,
    builtAt: new Date().toISOString()
  };
  try {
    const tmp = {};
    scAddEntries_(rk, out, tmp);
    if (Number(tmp[rk + '#n']) <= 30) scPutMany_(tmp, COL_TTL);
  } catch (e) {}
  return out;
}
