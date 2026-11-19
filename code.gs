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
    version: '1.0'
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
      return jsonResponse_(handleGetMetadata_());
    }

    if (action === 'getSchoolData') {
      return jsonResponse_(handleGetSchoolData_(req.kodSekolah));
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
function handleGetMetadata_() {
  const categories = ['MURID', 'GURU', 'AKP'];
  const headersMap = {};
  const countsMap = {};
  const schoolMap = {};

  categories.forEach(function (cat) {
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

    for (let r = 0; r < kodVals.length; r++) {
      const kod = normalizeCode_(kodVals[r][0]);
      if (!kod) continue;
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
    schools: schools
  };
}

/**
 * Ambil semua rekod Murid, Guru, dan AKP bagi satu Kod Sekolah tertentu
 */
function handleGetSchoolData_(kodInput) {
  const targetKod = normalizeCode_(kodInput);
  if (!targetKod) {
    return { status: 'error', message: 'Sila masukkan Kod Sekolah yang sah.' };
  }

  const categories = ['MURID', 'GURU', 'AKP'];
  const result = {
    status: 'success',
    kodSekolah: targetKod,
    namaSekolah: '',
    data: {}
  };

  const CHUNK_SIZE = 10000;

  categories.forEach(function (cat) {
    const sheet = getTargetSheet_(cat);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      result.data[cat] = { headers: [], rows: [] };
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h || '').trim();
    });

    const kodIdx = findColIndex_(headers, ['KODSEKOLAH', 'KOD_SEKOLAH', 'KOD SEKOLAH', 'KOD']);
    const namaIdx = findColIndex_(headers, ['NAMA_SEKOLAH', 'NAMASEKOLAH', 'NAMA SEKOLAH', 'SEKOLAH']);

    if (kodIdx === -1) {
      result.data[cat] = { headers: headers, rows: [] };
      return;
    }

    const matchedRows = [];

    // Baca secara berperingkat (chunk) bagi mengelakkan had memori Apps Script
    for (let startRow = 2; startRow <= lastRow; startRow += CHUNK_SIZE) {
      const numRows = Math.min(CHUNK_SIZE, lastRow - startRow + 1);
      const values = sheet.getRange(startRow, 1, numRows, lastCol).getDisplayValues();

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (normalizeCode_(row[kodIdx]) === targetKod) {
          matchedRows.push(row);
          if (!result.namaSekolah && namaIdx !== -1 && row[namaIdx]) {
            result.namaSekolah = String(row[namaIdx]).trim();
          }
        }
      }
    }

    result.data[cat] = {
      headers: headers,
      rows: matchedRows
    };
  });

  return result;
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
