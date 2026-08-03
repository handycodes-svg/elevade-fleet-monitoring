/**
 * ELEVADE Fleet Monitoring — Screening MR2 Only (Fixed Reg & Matching) — Google Apps Script
 * =========================================================================================
 * Alur:
 *   1. Upload CSV export ELEVADE (format "INBOX Warning & Fault Messages")
 *      ke folder Google Drive yang ditentukan di CONFIG.DRIVE_FOLDER_ID.
 *   2. Buka spreadsheet ini -> menu "ELEVADE MR2" -> "Import CSV & Rebuild Screening MR2".
 *   3. Script akan:
 *        - Baca semua CSV di folder & gabungkan dengan _DATA_STORE
 *        - Buang duplikat & data lebih tua dari RETENTION_DAYS
 *        - Ambil daftar OPEN MR2 dari spreadsheet eksternal (mengabaikan MR3)
 *        - Cocokkan ACARS vs MR2 dengan isolasi A/C Reg ketat & filter ATA Chapter.
 *        - Kolom "Status MR2" HANYA menampilkan "MR2" bila cocok (kosong bila tidak ada).
 *
 * === FIX LOG ===
 * [31 Jul 2026] DRIVE_FOLDER_ID diperbaiki -> sebelumnya menunjuk ke folder yang SALAH
 *               (folder itu isinya "IAA A/C STATUS AND NTC LIST", bukan ELEVADE_CSV_INBOX).
 *               Sekarang menunjuk ke folder ELEVADE_CSV_INBOX yang benar.
 * [31 Jul 2026] Deteksi file CSV diperluas: cek ekstensi .csv DAN MIME type
 *               (MimeType.CSV / text/csv / text/plain), supaya file yang di-upload
 *               tanpa ekstensi jelas tetap terbaca.
 * [31 Jul 2026] Pesan error "Tidak ada CSV" sekarang menampilkan daftar SEMUA file
 *               yang ada di folder (nama + MIME type) supaya gampang didiagnosis
 *               kalau masalah muncul lagi di kemudian hari.
 */

// ======================= CONFIG =======================
var CONFIG = {
  DRIVE_FOLDER_ID: '1kFEMkoTInhlRE0FCe9EtxyT69QKOVX-N',  // <-- FIXED: folder ELEVADE_CSV_INBOX yang benar

  DATA_STORE_SHEET: '_DATA_STORE',
  OUTPUT_SHEET: 'Screening MR2',
  PROCESSED_SUBFOLDER: 'processed',

  WF_WINDOW_MIN: 10,          // Fault terkait warning bila <=10 menit (flight & tanggal sama)
  NEW_THRESHOLD_DAYS: 3,
  NO_RECENT_THRESHOLD_DAYS: 21,
  RETENTION_DAYS: 30,         // Rolling window 30 hari dari tanggal data TERBARU

  // Sumber status MR2 (spreadsheet TERPISAH)
  MR_SPREADSHEET_ID: '1nU7X3rwHJFeFBixdqwYXqgBoIe1etu7v9pxdtBuqjcQ',
  MR_SHEET_NAME: 'STG_MR',
  MR_MR3_MARKER: 'OPEN MR3',  // Batas akhir blok MR2 (berhenti membaca saat teks ini muncul)
  MR_MIN_KEYWORDS: 2,         // Minimal kata kunci cocok untuk dianggap match MR2

  // ATA chapter yang dianggap terkait walau tidak identik (77 <-> 71, 77 <-> 75)
  ATA_EQUIV: {
    '77': ['71', '75', '73', '70'],
    '71': ['77'],
    '75': ['77'],
    '73': ['77'],
    '70': ['77']
  },

  COLUMNS: ['Priority','Message','Message Type','Report','A/C Model','A/C Reg',
            'Flight No','Phase','ATA','Source','Occurence','Sent','Repetitive'],

  // === EMAIL SUMMARY CONFIG ===
  EMAIL_RECIPIENTS: ['handy@airasia.com,irwansyahendra@airasia.com,nugrohobudiprayitno@airasia.com,aliabdullah@airasia.com,wawangunawan@airasia.com,nursumirat@airasia.com,bibid@airasia.com,ruslisaidih@airasia.com,pujiadiprabowo@airasia.com,dimasmahardka@airasia.com,muhammadbagus@airasia.com,mochsellabramantio@airasia.com'],
  EMAIL_SUBJECT_PREFIX: 'ELEVADE Monitoring',
  EMAIL_TRIGGER_HOUR: 17,  // jam 17:00 / 5 sore

  // Message yang TIDAK perlu tampil di email (operational/nuisance, bukan defect).
  EMAIL_EXCLUDE_MESSAGES: [
    'BRAKES HOT'
  ]
};

// Daftar spurious message (exact match, whitespace-stripped)
var SPURIOUS_MESSAGES = [
  'ICAO ADDRESS/HF1(3RE1)', 'NO SDU DATA', 'COM SATCOM FAULT',
  'FWC2 :NO DATA FROM ECU1B', 'COM SATCOM DATA FAULT',
  'SDAC2:NO DATA FROM PHC2', 'SDAC2:NO DATA FROM PHC1',
  'SDAC1:NO DATA FROM PHC2', 'SDAC1:NO DATA FROM PHC1',
  'AFS:FADEC2', 'NO EIU 2 DATA', 'NO EIU 2 DATA (INTM)',
  'AFS:FADEC1', 'POWER SUPPLY INTERRUPT', 'FWC1 :NO DATA FROM FCDC1',
  'AFS:ELAC2', 'FWC1 :NO DATA FROM FQI1A/1B', 'FWC2 :NO DATA FROM FQI1A/1B',
  'ICAO ADDRESS/HF2(3RE2)','DIR1 (101RH) / OBRM','AFS:RA1','AFS:28V PWR 11XU1','AFS:FMGC2','AFS:FMGC1','AFS:ELAC1','OBRM/CIDS(101RH)-SDF1/',''
];

// ======================= MENU =======================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ELEVADE MR2')
    .addItem('Import CSV & Rebuild Screening MR2', 'importAndRebuild')
    .addSeparator()
    .addItem('Rebuild from stored data only', 'rebuildFromStore')
    .addItem('Show data store info', 'showStoreInfo')
    .addItem('Diagnose MR2 source', 'diagnoseMr2')
    .addSeparator()
    .addItem('Setup Email Trigger (Jam 17:00)', 'setupDailyEmailTrigger')
    .addItem('Kirim Summary Email Sekarang (Test)', 'sendScreeningSummaryEmail')
    .addItem('Kirim Email ke Saya Saja (Manual)', 'sendScreeningSummaryEmailToMe')
    .addItem('Hapus Trigger Email', 'removeEmailTriggers')
    .addToUi();
}

// ======================= MAIN =======================
function importAndRebuild() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (CONFIG.DRIVE_FOLDER_ID === 'PASTE_FOLDER_ID_DISINI') {
      ui.alert('Setup belum lengkap', 'Isi dulu CONFIG.DRIVE_FOLDER_ID dengan ID folder Google Drive tempat upload CSV.', ui.ButtonSet.OK);
      return;
    }

    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    var newRows = [];
    var processedNames = [];
    var fileList = [];
    var allFilesDebug = [];

    var it = folder.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      var fname = f.getName();
      var mime = f.getMimeType();
      allFilesDebug.push(fname + '  [' + mime + ']');

      var isCsvExt  = fname.toLowerCase().trim().slice(-4) === '.csv';
      var isCsvMime = (mime === MimeType.CSV || mime === 'text/csv' || mime === 'text/plain');

      if (isCsvExt || isCsvMime) {
        fileList.push(f);
      }
    }

    if (fileList.length === 0) {
      var debugMsg = 'Tidak ada file CSV baru yang dikenali di folder Drive.\n\n';
      if (allFilesDebug.length === 0) {
        debugMsg += 'Folder benar-benar kosong.\nFolder ID: ' + CONFIG.DRIVE_FOLDER_ID +
                    '\n\nCek apakah file di-upload ke folder yang BENAR (bukan folder lain dengan nama mirip).';
      } else {
        debugMsg += 'File yang ADA di folder (' + allFilesDebug.length + '):\n' +
                    allFilesDebug.join('\n') +
                    '\n\nTidak ada satupun yang cocok kriteria CSV (ekstensi .csv atau MIME CSV/text). ' +
                    'Cek nama & format file export di atas — mungkin ekstensinya bukan .csv, atau ada spasi/karakter aneh di akhir nama file.';
      }
      ui.alert('Tidak ada CSV', debugMsg, ui.ButtonSet.OK);
      return;
    }

    for (var i = 0; i < fileList.length; i++) {
      var parsed = parseElevadeCsv(fileList[i].getBlob().getDataAsString());
      for (var j = 0; j < parsed.length; j++) newRows.push(parsed[j]);
      processedNames.push(fileList[i].getName());
    }

    var store = loadDataStore();
    var combined = store.concat(newRows);
    var deduped = dedupRows(combined);

    var pruneResult = pruneOldData(deduped);
    var kept = pruneResult.kept;
    saveDataStore(kept);

    moveToProcessed(folder, fileList);
    buildMr2ScreeningSheet(kept);

    var cutoffMsg = pruneResult.cutoff
      ? '\nRetensi ' + CONFIG.RETENTION_DAYS + ' hari: data sebelum ' + fmtDate(pruneResult.cutoff) + ' dibuang (' + pruneResult.removed + ' baris).'
      : '';
    ui.alert('Selesai',
      processedNames.length + ' file CSV diproses.\n' +
      'Total data tersimpan: ' + kept.length + ' baris.' + cutoffMsg + '\n' +
      'Sheet "' + CONFIG.OUTPUT_SHEET + '" sudah diperbarui.',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', String(e) + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
  }
}

function rebuildFromStore() {
  var ui = SpreadsheetApp.getUi();
  var store = loadDataStore();
  if (store.length === 0) {
    ui.alert('Data store kosong', 'Belum ada data tersimpan. Jalankan "Import CSV & Rebuild" dulu.', ui.ButtonSet.OK);
    return;
  }
  buildMr2ScreeningSheet(store);
  ui.alert('Selesai', 'Sheet "' + CONFIG.OUTPUT_SHEET + '" di-rebuild dari ' + store.length + ' baris tersimpan.', ui.ButtonSet.OK);
}

function showStoreInfo() {
  var store = loadDataStore();
  var ui = SpreadsheetApp.getUi();
  if (store.length === 0) { ui.alert('Data store kosong.'); return; }
  var minD = null, maxD = null, acSet = {};
  for (var i = 0; i < store.length; i++) {
    var d = store[i].Occurence;
    if (d) { if (!minD || d < minD) minD = d; if (!maxD || d > maxD) maxD = d; }
    acSet[store[i]['A/C Reg']] = true;
  }
  ui.alert('Data store info',
    'Total baris: ' + store.length + '\n' +
    'Rentang: ' + fmtDate(minD) + '  s/d  ' + fmtDate(maxD) + '\n' +
    'Jumlah pesawat: ' + Object.keys(acSet).length,
    ui.ButtonSet.OK);
}

// ======================= CSV PARSING =======================
function parseElevadeCsv(text) {
  var allRows = Utilities.parseCsv(text);
  if (allRows.length === 0) return [];

  var headerIdx = -1;
  var scanLimit = Math.min(allRows.length, 30);
  for (var hr = 0; hr < scanLimit; hr++) {
    var rowCells = allRows[hr].map(function(x) { return String(x).trim(); });
    if (rowCells.indexOf('Message') !== -1 &&
        rowCells.indexOf('Occurence') !== -1 &&
        rowCells.indexOf('A/C Reg') !== -1) {
      headerIdx = hr;
      break;
    }
  }
  if (headerIdx === -1) {
    throw new Error('Format CSV tidak dikenali: baris header tidak ditemukan di 30 baris pertama.');
  }

  var header = allRows[headerIdx];
  var idx = {};
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim();
    if (CONFIG.COLUMNS.indexOf(h) !== -1) idx[h] = c;
  }

  var out = [];
  for (var r = headerIdx + 1; r < allRows.length; r++) {
    var row = allRows[r];
    if (!row || row.length === 0) continue;
    var msg = idx['Message'] !== undefined ? String(row[idx['Message']]).trim() : '';
    if (msg === '') continue;
    var occ = parseElevadeDate(row[idx['Occurence']]);
    if (!occ) continue;

    out.push({
      'Priority':     getCell(row, idx, 'Priority'),
      'Message':      msg,
      'Message Type': getCell(row, idx, 'Message Type'),
      'Report':       getCell(row, idx, 'Report'),
      'A/C Model':    getCell(row, idx, 'A/C Model'),
      'A/C Reg':      getCell(row, idx, 'A/C Reg'),
      'Flight No':    getCell(row, idx, 'Flight No'),
      'Phase':        getCell(row, idx, 'Phase'),
      'ATA':          getCell(row, idx, 'ATA'),
      'Source':       getCell(row, idx, 'Source'),
      'Occurence':    occ,
      'Sent':         getCell(row, idx, 'Sent'),
      'Repetitive':   getCell(row, idx, 'Repetitive')
    });
  }
  return out;
}

function getCell(row, idx, name) {
  if (idx[name] === undefined) return '';
  var v = row[idx[name]];
  return v === undefined || v === null ? '' : String(v).trim();
}

function parseElevadeDate(val) {
  if (val === undefined || val === null) return null;
  var s = String(val).trim();
  if (s === '') return null;
  s = s.replace(' at ', ' ');
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    return new Date(
      parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10),
      parseInt(m[4],10), parseInt(m[5],10), m[6] ? parseInt(m[6],10) : 0);
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ======================= DATA STORE =======================
function loadDataStore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.DATA_STORE_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(1, 1, last, CONFIG.COLUMNS.length).getValues();
  var header = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var obj = {};
    for (var c = 0; c < header.length; c++) {
      obj[header[c]] = row[c];
    }
    if (obj['Occurence'] && !(obj['Occurence'] instanceof Date)) {
      obj['Occurence'] = parseElevadeDate(obj['Occurence']);
    }
    if (obj['Message'] && String(obj['Message']).trim() !== '') out.push(obj);
  }
  return out;
}

function saveDataStore(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.DATA_STORE_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.DATA_STORE_SHEET);
  sh.clear();
  var out = [CONFIG.COLUMNS.slice()];
  for (var i = 0; i < rows.length; i++) {
    var row = [];
    for (var c = 0; c < CONFIG.COLUMNS.length; c++) {
      var key = CONFIG.COLUMNS[c];
      var v = rows[i][key];
      if (key === 'Occurence' && v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      row.push(v === undefined || v === null ? '' : v);
    }
    out.push(row);
  }
  sh.getRange(1, 1, out.length, CONFIG.COLUMNS.length).setValues(out);
  sh.hideSheet();
}

function pruneOldData(rows) {
  if (rows.length === 0) return { kept: rows, removed: 0, cutoff: null };
  var maxTime = null;
  for (var i = 0; i < rows.length; i++) {
    var occ = rows[i]['Occurence'];
    if (!(occ instanceof Date)) occ = parseElevadeDate(occ);
    if (occ && (maxTime === null || occ > maxTime)) maxTime = occ;
  }
  if (maxTime === null) return { kept: rows, removed: 0, cutoff: null };

  var maxDate = new Date(maxTime.getFullYear(), maxTime.getMonth(), maxTime.getDate());
  var cutoff = new Date(maxDate.getTime());
  cutoff.setDate(cutoff.getDate() - (CONFIG.RETENTION_DAYS - 1));

  var kept = [];
  var removed = 0;
  for (var i = 0; i < rows.length; i++) {
    var occ = rows[i]['Occurence'];
    if (!(occ instanceof Date)) occ = parseElevadeDate(occ);
    if (occ && occ >= cutoff) kept.push(rows[i]);
    else removed++;
  }
  return { kept: kept, removed: removed, cutoff: cutoff };
}

function dedupRows(rows) {
  var seen = {};
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var occ = r['Occurence'];
    var occKey = (occ instanceof Date)
      ? Utilities.formatDate(occ, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : String(occ);
    var key = String(r['A/C Reg']).trim() + '||' + String(r['Message']).trim() + '||' +
              String(r['ATA']).trim() + '||' + occKey;
    if (!seen[key]) {
      seen[key] = true;
      out.push(r);
    }
  }
  return out;
}

// ======================= HELPERS (Logic) =======================
function ataRelated(wAta, fAta) {
  if (!wAta || !fAta) return true;
  var wCh = String(wAta).slice(0, 2);
  var fCh = String(fAta).slice(0, 2);
  if (wCh === fCh) return true;
  var equiv = CONFIG.ATA_EQUIV[wCh];
  if (equiv && equiv.indexOf(fCh) !== -1) return true;
  return false;
}

var SUBSYS_KEYWORDS = ['ELAC', 'SEC', 'SFCC', 'SLAT', 'FLAP', 'THS', 'RUDDER',
                       'AILERON', 'SPOILER', 'FCDC', 'STABILIZER', 'PITCH', 'ROLL', 'YAW'];
var SUBSYS_SYNONYMS = { 'SLT': 'SLAT', 'RUD': 'RUDDER', 'AIL': 'AILERON', 'SPLR': 'SPOILER' };

function subsystemsOf(msg) {
  var m = String(msg).toUpperCase();
  for (var syn in SUBSYS_SYNONYMS) {
    m = m.replace(new RegExp('\\b' + syn + '\\b', 'g'), SUBSYS_SYNONYMS[syn]);
  }
  var found = {};
  for (var i = 0; i < SUBSYS_KEYWORDS.length; i++) {
    if (m.indexOf(SUBSYS_KEYWORDS[i]) !== -1) found[SUBSYS_KEYWORDS[i]] = true;
  }
  return found;
}

function correlationOk(wAta, fAta, wMsg, fMsg) {
  if (!ataRelated(wAta, fAta)) return false;
  var wCh = String(wAta).slice(0, 2);
  if (wCh === '27') {
    var ws = subsystemsOf(wMsg);
    var wHasSubsys = false;
    for (var k in ws) { wHasSubsys = true; break; }
    if (!wHasSubsys) return true;

    if (String(wAta).slice(0, 4) === String(fAta).slice(0, 4)) return true;
    var fs = subsystemsOf(fMsg);
    for (var k2 in ws) { if (fs[k2]) return true; }
    return false;
  }
  return true;
}

function descKeywordMatch(warnMsg, faultMsg) {
  var stop = {'1':1,'2':1,'3':1,'A':1,'B':1,'L':1,'R':1,'SYS':1,'OR':1,'OF':1,
              'NO':1,'FROM':1,'FAULT':1,'THE':1,'AND':1,'TO':1,'IN':1};
  function toks(s) {
    var parts = String(s).toUpperCase().split(/[^A-Z0-9]+/);
    var set = {};
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p.length >= 2 && !stop[p] && !/^\d+$/.test(p)) set[p] = true;
    }
    return set;
  }
  var wt = toks(warnMsg), ft = toks(faultMsg);
  for (var k in wt) { if (ft[k]) return true; }
  return false;
}

function aggPriority(list) {
  var rank = { 'High': 3, 'Medium': 2, 'Low': 1, 'None': 0, '': 0 };
  var best = 'None', bestR = -1;
  for (var i = 0; i < list.length; i++) {
    var p = list[i] && String(list[i]).trim() !== '' ? String(list[i]).trim() : 'None';
    var rr = rank[p] === undefined ? 0 : rank[p];
    if (rr > bestR) { bestR = rr; best = p; }
  }
  return best;
}

function trendOf(count, firstDate, lastDate, snapshot) {
  var daysFirst = daysBetween(firstDate, snapshot);
  var daysLast = daysBetween(lastDate, snapshot);
  if (daysFirst <= CONFIG.NEW_THRESHOLD_DAYS) return 'NEW ' + count + 'x';
  if (daysLast > CONFIG.NO_RECENT_THRESHOLD_DAYS) return 'no recent';
  return 'stable ' + count + 'x';
}

function daysBetween(d1, d2) {
  var a = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
  var b = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function recommendedAction(count, trend, priority, hasMr2) {
  var isNew = trend.indexOf('NEW') === 0;
  var isNoRecent = (trend === 'no recent');
  if (hasMr2) return ['MONITOR (MR2 OPEN)', 'watch'];
  if (priority === 'High' && isNew) return ['INVESTIGATE NOW', 'urgent'];
  if (count >= 10 && trend.indexOf('stable') === 0) return ['SCHEDULE INSPECTION (chronic)', 'high'];
  if (isNew) return ['MONITOR next flights', 'watch'];
  if (isNoRecent) return ['MONITOR only (may be cleared)', 'low'];
  return ['ROUTINE monitoring', 'low'];
}

function fmtDate(d) {
  if (!d) return '';
  if (!(d instanceof Date)) { d = parseElevadeDate(d); if (!d) return ''; }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

function sameMinuteWindow(t1, t2, minutes) {
  return Math.abs(t1.getTime() - t2.getTime()) <= minutes * 60 * 1000;
}

function sameDate(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

function modeKey(counts) {
  var best = '', bestN = -1;
  for (var k in counts) { if (counts[k] > bestN) { bestN = counts[k]; best = k; } }
  return best;
}

// ======================= MR2 ONLY LOOKUP (FIXED & ISOLATED) =======================
function diagnoseMr2() {
  var ui = SpreadsheetApp.getUi();
  var msg = '';
  try {
    var ss = SpreadsheetApp.openById(CONFIG.MR_SPREADSHEET_ID);
    msg += '✓ Spreadsheet terbuka: "' + ss.getName() + '"\n';
    var sh = ss.getSheetByName(CONFIG.MR_SHEET_NAME);
    if (!sh) {
      msg += '✗ Tab "' + CONFIG.MR_SHEET_NAME + '" TIDAK ketemu.\n';
      ui.alert('Diagnosa MR2', msg, ui.ButtonSet.OK);
      return;
    }
    msg += '✓ Tab "' + CONFIG.MR_SHEET_NAME + '" ketemu.\n';

    var list = loadMr2List();
    var regs = {};
    for (var i = 0; i < list.length; i++) regs[list[i].regCode] = true;
    msg += 'Baris MR2 aktif terbaca: ' + list.length + '\n';
    msg += 'Kode Reg MR2 terdeteksi (' + Object.keys(regs).length + ' pesawat): ' + Object.keys(regs).sort().join(', ') + '\n\n';

    if (list.length > 0) {
      msg += 'Contoh baris MR2 pertama:\n';
      for (var j = 0; j < Math.min(6, list.length); j++) {
        var ataTag = list[j].ata ? ' [ATA ' + list[j].ata + ']' : '';
        msg += '  [' + list[j].regCode + ataTag + '] ' + list[j].desc.slice(0, 45) + '\n';
      }
    }
  } catch (e) {
    msg += '✗ ERROR: ' + e + '\n\nPastikan akun memiliki hak akses view.';
  }
  ui.alert('Diagnosa MR2 Only', msg, ui.ButtonSet.OK);
}

function loadMr2List() {
  var out = [];
  var ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.MR_SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Gagal buka spreadsheet MR (ID: ' + CONFIG.MR_SPREADSHEET_ID + '). Detail: ' + e);
  }
  var sh = ss.getSheetByName(CONFIG.MR_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + CONFIG.MR_SHEET_NAME + '" tidak ditemukan.');
  
  var last = sh.getLastRow();
  if (last < 2) return out;
  var values = sh.getRange(1, 1, last, 4).getValues();

  var currentReg = '';
  for (var r = 1; r < values.length; r++) {
    var a = String(values[r][0] || '').trim();
    var b = String(values[r][1] || '').trim();
    var c = String(values[r][2] || '').trim();
    var d = String(values[r][3] || '').trim();

    var rowText = (a + ' ' + b + ' ' + c + ' ' + d).toUpperCase();

    if (rowText.indexOf(CONFIG.MR_MR3_MARKER.toUpperCase()) !== -1) break;
    if (rowText.indexOf('NON MEL CATEGORY') !== -1) break;
    if (rowText.indexOf('OPEN MR3') !== -1) break;

    var aUpper = a.toUpperCase();
    if (aUpper.indexOf('TOTAL') === 0 || aUpper.indexOf('DATE') === 0 || 
        aUpper.indexOf('A/C REG') === 0 || aUpper === 'REG' || aUpper === '') {
      if (b === '') currentReg = '';
    }

    var regHere = extractRegCode(a);
    if (regHere) {
      currentReg = regHere;
    }

    if (b === '' || b.toUpperCase() === 'NIL' || b.toUpperCase() === 'DESC' || b.toUpperCase() === 'DESCRIPTION') continue;

    var regCode = regHere || currentReg;
    if (!regCode) continue;

    var ataCh = extractAtaChapter(c, b);
    out.push({ regCode: regCode, desc: b, ata: ataCh, status: 'MR2' });
  }
  return out;
}

var IGNORE_REGS = {
  'DAT':1, 'DTE':1, 'TOT':1, 'NIL':1, 'MEL':1, 'OPN':1, 'MR2':1, 'MR3':1,
  'CAT':1, 'ITM':1, 'DES':1, 'REF':1, 'REG':1, 'MNT':1, 'STG':1, 'FLT':1,
  'SUM':1, 'DAY':1, 'MON':1, 'TUE':1, 'WED':1, 'THU':1, 'FRI':1, 'SAT':1,
  'SUN':1, 'JAN':1, 'FEB':1, 'MAR':1, 'APR':1, 'MAY':1, 'JUN':1, 'JUL':1,
  'AUG':1, 'SEP':1, 'OCT':1, 'NOV':1, 'DEC':1, 'NON':1, 'AIR':1, 'ENG':1
};

function extractRegCode(s) {
  if (!s) return '';
  var t = String(s).trim().toUpperCase();
  t = t.replace(/^A\/C\s*REG\s*[:#-]?\s*/i, '').replace(/^REG\s*[:#-]?\s*/i, '');
  
  var mPk = t.match(/PK-?([A-Z0-9]{3})\b/);
  if (mPk && !IGNORE_REGS[mPk[1]]) return mPk[1];
  
  var mStart = t.match(/^([A-Z]{3})\b/);
  if (mStart && !IGNORE_REGS[mStart[1]]) return mStart[1];
  
  return '';
}

function extractAtaChapter(refStr, descStr) {
  var combined = (String(refStr || '') + ' ' + String(descStr || '')).toUpperCase();
  var m = combined.match(/\bATA\s*(\d{2})\b|\b(\d{2})-\d{2}\b|\b(\d{2})\.\d{2}\b/);
  if (m) return m[1] || m[2] || m[3];
  return '';
}

function extractUnits(s) {
  var m = String(s).toUpperCase();
  var units = {};
  if (/\bAPU\b/.test(m)) units['APU'] = true;
  if (/\bENG\s*1\b|\bENG1\b|\bENGINE\s*1\b/.test(m)) units['ENG1'] = true;
  if (/\bENG\s*2\b|\bENG2\b|\bENGINE\s*2\b/.test(m)) units['ENG2'] = true;
  if (/\bENG\s*3\b|\bENG3\b|\bENGINE\s*3\b/.test(m)) units['ENG3'] = true;
  if (/\bL\s*WING\b|\bLEFT\s*WING\b/.test(m)) units['LWING'] = true;
  if (/\bR\s*WING\b|\bRIGHT\s*WING\b/.test(m)) units['RWING'] = true;
  if (/\bL\s*TK\b|\bLEFT\s*TANK\b/.test(m)) units['LTK'] = true;
  if (/\bR\s*TK\b|\bRIGHT\s*TANK\b/.test(m)) units['RTK'] = true;
  if (/\bCTR\b|\bCENTER\b|\bCENTRE\b/.test(m)) units['CTR'] = true;
  return units;
}

function unitsConflict(msgUnits, mrUnits) {
  var mHas = false, dHas = false;
  for (var k in msgUnits) { mHas = true; break; }
  for (var k2 in mrUnits) { dHas = true; break; }
  if (!mHas || !dHas) return false;
  for (var u in msgUnits) { if (mrUnits[u]) return false; }
  return true;
}

function mrTokens(s) {
  var stop = {'AND':1,'THE':1,'FOR':1,'WAS':1,'HAS':1,'BEEN':1,'ON':1,'IN':1,'AT':1,'OF':1,
              'TO':1,'OR':1,'IS':1,'AS':1,'PER':1,'NOT':1,'BUT':1,'A':1,'AN':1,'NIL':1,
              'FROM':1,'WITH':1,'FOUND':1,'FAULT':1,'CHK':1,'CHECK':1,'SYS':1,'SYSTEM':1,
              'MSG':1,'MESSAGE':1,'DUE':1,'REQ':1,'REQUIRED':1,'MAINT':1,'MAINTENANCE':1,
              'OPEN':1,'CLOSE':1,'CLOSED':1,'TEST':1,'OPS':1,'FLIGHT':1,'AFTER':1,'BEFORE':1,
              'DURING':1,'REPORTED':1,'REPORT':1,'AIR':1};
  var parts = String(s).toUpperCase().split(/[^A-Z0-9]+/);
  var set = {};
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (p.length >= 3 && !stop[p] && !/^\d+$/.test(p)) set[p] = true;
  }
  return set;
}

function lookupMr2Status(msg, ata, acReg, mrIndex) {
  var reg = extractRegCode(acReg);
  if (!reg || !mrIndex[reg]) return '';
  var list = mrIndex[reg];
  var mtoks = mrTokens(msg);
  var mcount = 0; for (var k in mtoks) mcount++;
  if (mcount === 0) return '';
  var need = Math.min(CONFIG.MR_MIN_KEYWORDS, mcount);
  var msgUnits = extractUnits(msg);

  for (var i = 0; i < list.length; i++) {
    if (list[i].ata && ata && !ataRelated(ata, list[i].ata)) continue;

    if (unitsConflict(msgUnits, list[i].units)) continue;

    var dtoks = list[i].tokens;
    var common = 0;
    for (var t in mtoks) { if (dtoks[t]) common++; }
    if (common >= need) return 'MR2';
  }
  return '';
}

function buildMr2Index(mrList) {
  var idx = {};
  for (var i = 0; i < mrList.length; i++) {
    var reg = mrList[i].regCode;
    if (!idx[reg]) idx[reg] = [];
    idx[reg].push({
      desc: mrList[i].desc,
      ata: mrList[i].ata,
      status: 'MR2',
      tokens: mrTokens(mrList[i].desc),
      units: extractUnits(mrList[i].desc)
    });
  }
  return idx;
}

// ======================= BUILD OUTPUT SHEET =======================
function buildMr2ScreeningSheet(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var spuriousSet = {};
  for (var i = 0; i < SPURIOUS_MESSAGES.length; i++) spuriousSet[SPURIOUS_MESSAGES[i].trim()] = true;
  
  var data = [];
  var snapshot = null;
  for (var i = 0; i < rows.length; i++) {
    var msg = String(rows[i]['Message']).trim();
    if (spuriousSet[msg]) continue;
    var occ = rows[i]['Occurence'];
    if (!(occ instanceof Date)) occ = parseElevadeDate(occ);
    if (!occ) continue;
    data.push({
      ac: String(rows[i]['A/C Reg']).trim(),
      msg: msg,
      type: String(rows[i]['Message Type']).trim(),
      report: String(rows[i]['Report']).trim(),
      flight: String(rows[i]['Flight No']).trim(),
      ata: String(rows[i]['ATA']).trim(),
      source: String(rows[i]['Source']).trim(),
      priority: String(rows[i]['Priority']).trim(),
      occ: occ
    });
    if (!snapshot || occ > snapshot) snapshot = occ;
  }
  var snapDate = new Date(snapshot.getFullYear(), snapshot.getMonth(), snapshot.getDate());

  var mrIndex = {};
  try {
    mrIndex = buildMr2Index(loadMr2List());
  } catch (e) {
    Logger.log('MR2 load error: ' + e);
  }

  var faultsByFlight = {};
  var warnings = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i].type === 'WARNING') warnings.push(data[i]);
    else if (data[i].type === 'FAULT') {
      var k = data[i].ac + '|' + data[i].flight;
      if (!faultsByFlight[k]) faultsByFlight[k] = [];
      faultsByFlight[k].push(data[i]);
    }
  }

  var warnGroups = {};
  for (var i = 0; i < warnings.length; i++) {
    var w = warnings[i];
    var key = w.ac + '|' + w.msg;
    if (!warnGroups[key]) warnGroups[key] = { ac: w.ac, msg: w.msg, occs: [], prios: [], atas: {}, rows: [] };
    warnGroups[key].occs.push(w.occ);
    warnGroups[key].prios.push(w.priority);
    warnGroups[key].atas[w.ata] = (warnGroups[key].atas[w.ata] || 0) + 1;
    warnGroups[key].rows.push(w);
  }

  var explained = {};
  var warnFaultMap = {};
  var warnLastTime = {};

  for (var key in warnGroups) {
    var g = warnGroups[key];
    var lastOcc = g.occs[0], lastRow = g.rows[0];
    for (var j = 1; j < g.occs.length; j++) {
      if (g.occs[j] > lastOcc) { lastOcc = g.occs[j]; lastRow = g.rows[j]; }
    }
    warnLastTime[key] = lastOcc;
    var wAta = lastRow.ata;
    var fkey = g.ac + '|' + lastRow.flight;
    var candidateFaults = faultsByFlight[fkey] || [];

    var matched = {};
    for (var f = 0; f < candidateFaults.length; f++) {
      var ft = candidateFaults[f];
      if (!sameDate(ft.occ, lastOcc)) continue;
      if (!sameMinuteWindow(ft.occ, lastOcc, CONFIG.WF_WINDOW_MIN)) continue;
      if (!correlationOk(wAta, ft.ata, g.msg, ft.msg)) continue;
      var mk = ft.msg + '||' + ft.ata;
      var dm = descKeywordMatch(g.msg, ft.msg);
      if (!matched[mk] || ft.occ > matched[mk].cooc) {
        matched[mk] = { fmsg: ft.msg, fata: ft.ata, source: ft.source, cooc: ft.occ, descMatch: dm };
      }
      if (!explained[g.ac]) explained[g.ac] = {};
      explained[g.ac][mk] = true;
    }
    var arr = [];
    for (var mk in matched) arr.push(matched[mk]);
    arr.sort(function(a, b) { return a.fmsg < b.fmsg ? -1 : (a.fmsg > b.fmsg ? 1 : 0); });
    warnFaultMap[key] = arr;
  }

  var faultGroups = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i].type !== 'FAULT') continue;
    var f = data[i];
    var key = f.ac + '|' + f.msg;
    if (!faultGroups[key]) faultGroups[key] = { ac: f.ac, msg: f.msg, occs: [], prios: [], atas: {}, sources: {} };
    faultGroups[key].occs.push(f.occ);
    faultGroups[key].prios.push(f.priority);
    faultGroups[key].atas[f.ata] = (faultGroups[key].atas[f.ata] || 0) + 1;
    faultGroups[key].sources[f.source] = (faultGroups[key].sources[f.source] || 0) + 1;
  }

  var acSet = {};
  for (var i = 0; i < data.length; i++) acSet[data[i].ac] = true;
  var acList = Object.keys(acSet).sort();

  var HEADERS = ['Type / Item', 'ATA', 'Source', 'Count', 'Priority', 'Trend', 'Last Occurred', 'Recommended Action', 'Status MR2'];
  var outRows = [];
  outRows.push({ cells: HEADERS, style: 'header' });

  for (var a = 0; a < acList.length; a++) {
    var ac = acList[a];
    outRows.push({ cells: ['✈  ' + ac, '', '', '', '', '', '', '', ''], style: 'band' });

    var wKeys = [];
    for (var key in warnGroups) if (warnGroups[key].ac === ac) wKeys.push(key);
    wKeys.sort(function(k1, k2) { return warnGroups[k2].occs.length - warnGroups[k1].occs.length; });

    for (var wi = 0; wi < wKeys.length; wi++) {
      var key = wKeys[wi];
      var g = warnGroups[key];
      var cnt = g.occs.length;
      var prio = aggPriority(g.prios);
      var firstOcc = g.occs[0], lastOcc = g.occs[0];
      for (var j = 1; j < g.occs.length; j++) {
        if (g.occs[j] < firstOcc) firstOcc = g.occs[j];
        if (g.occs[j] > lastOcc) lastOcc = g.occs[j];
      }
      var trend = trendOf(cnt, firstOcc, lastOcc, snapDate);
      var ataMode = modeKey(g.atas);
      
      var mr2Status = lookupMr2Status(g.msg, ataMode, ac, mrIndex);
      var hasMr2 = (mr2Status === 'MR2');
      var act = recommendedAction(cnt, trend, prio, hasMr2);

      outRows.push({
        cells: ['⚠ WARNING: ' + g.msg, ataMode, '', cnt, prio, trend, fmtDate(warnLastTime[key]), act[0], mr2Status],
        style: 'warn', level: act[1]
      });

      var faults = warnFaultMap[key] || [];
      if (faults.length > 0) {
        for (var f = 0; f < faults.length; f++) {
          var descTag = faults[f].descMatch ? '  ✓desc' : '';
          outRows.push({
            cells: ['        └── ' + faults[f].fmsg + descTag, faults[f].fata, faults[f].source, '', '', '', fmtDate(faults[f].cooc), '', ''],
            style: 'cause', descMatch: faults[f].descMatch
          });
        }
      } else {
        outRows.push({
          cells: ['        └── (no ATA-matched fault at the last warning event)', '', '', '', '', '', '', '', ''],
          style: 'nocause'
        });
      }
    }

    var expl = explained[ac] || {};
    var sKeys = [];
    for (var key in faultGroups) {
      if (faultGroups[key].ac !== ac) continue;
      var fg = faultGroups[key];
      var isExplained = false;
      for (var at in fg.atas) {
        if (expl[fg.msg + '||' + at]) { isExplained = true; break; }
      }
      if (!isExplained) sKeys.push(key);
    }
    if (sKeys.length > 0) {
      sKeys.sort(function(k1, k2) { return faultGroups[k2].occs.length - faultGroups[k1].occs.length; });
      outRows.push({ cells: ['   ── FAULT without warning (early signal) ──', '', '', '', '', '', '', '', ''], style: 'subdiv' });
      for (var si = 0; si < sKeys.length; si++) {
        var fg = faultGroups[sKeys[si]];
        var cnt = fg.occs.length;
        var prio = aggPriority(fg.prios);
        var firstOcc = fg.occs[0], lastOcc = fg.occs[0];
        for (var j = 1; j < fg.occs.length; j++) {
          if (fg.occs[j] < firstOcc) firstOcc = fg.occs[j];
          if (fg.occs[j] > lastOcc) lastOcc = fg.occs[j];
        }
        var trend = trendOf(cnt, firstOcc, lastOcc, snapDate);
        var ataMode = modeKey(fg.atas);

        var mr2Status = lookupMr2Status(fg.msg, ataMode, ac, mrIndex);
        var hasMr2 = (mr2Status === 'MR2');
        var act = recommendedAction(cnt, trend, prio, hasMr2);

        outRows.push({
          cells: ['● FAULT: ' + fg.msg, ataMode, modeKey(fg.sources), cnt, prio, trend, fmtDate(lastOcc), act[0], mr2Status],
          style: 'fault', level: act[1]
        });
      }
    }
  }

  writeSheet(ss, outRows);
}

// ======================= WRITE & STYLE =======================
function writeSheet(ss, outRows) {
  var sh = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.OUTPUT_SHEET);
  sh.clear();
  sh.clearFormats();

  var nCols = 9;
  var titleRows = 3;

  sh.getRange(1, 1).setValue('MR2 SCREENING & PREVENTIVE VIEW — ELEVADE ACARS')
    .setFontSize(14).setFontWeight('bold').setFontColor('#1F4E78');
  sh.getRange(2, 1).setValue('Fokus pemantauan indikasi ACARS terhadap daftar OPEN MR2 pesawat di seluruh armada (mengabaikan status MR3).')
    .setFontStyle('italic').setFontSize(9).setFontColor('#666666');
  sh.getRange(3, 1).setValue('Kolom Status MR2 HANYA menampilkan "MR2" bila item ACARS cocok dengan OPEN MR2 aktif. Bila tidak ada/belum dibuka, sel dibiarkan kosong.')
    .setFontStyle('italic').setFontSize(9).setFontColor('#666666');

  var startRow = titleRows + 1;
  var values = [];
  for (var i = 0; i < outRows.length; i++) values.push(outRows[i].cells);
  sh.getRange(startRow, 1, values.length, nCols).setValues(values);

  var aircraftBandRows = [];
  for (var i = 0; i < outRows.length; i++) {
    var rowNum = startRow + i;
    var st = outRows[i].style;
    var rng = sh.getRange(rowNum, 1, 1, nCols);

    if (st === 'header') {
      rng.setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
      sh.setRowHeight(rowNum, 22);
    } else if (st === 'band') {
      sh.getRange(rowNum, 1, 1, nCols).merge();
      sh.getRange(rowNum, 1).setBackground('#1F4E78').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
      aircraftBandRows.push(rowNum);
    } else if (st === 'warn') {
      rng.setBackground('#FCE4D6');
      sh.getRange(rowNum, 1).setFontWeight('bold').setFontColor(priorityColor(outRows[i].cells[4]));
      sh.getRange(rowNum, 2).setFontWeight('bold').setFontColor('#C00000');
      styleAction(sh, rowNum, outRows[i].level);
      styleMr2Cell(sh, rowNum, outRows[i].cells[8]);
    } else if (st === 'cause') {
      var causeColor = outRows[i].descMatch ? '#006100' : '#555555';
      sh.getRange(rowNum, 1).setFontColor(causeColor).setFontWeight(outRows[i].descMatch ? 'bold' : 'normal').setFontSize(9);
      sh.getRange(rowNum, 2, 1, 7).setFontColor('#555555').setFontSize(9);
    } else if (st === 'nocause') {
      sh.getRange(rowNum, 1).setFontColor('#999999').setFontStyle('italic').setFontSize(9);
    } else if (st === 'subdiv') {
      sh.getRange(rowNum, 1).setFontColor('#1F4E78').setFontWeight('bold').setFontStyle('italic').setFontSize(9);
    } else if (st === 'fault') {
      rng.setBackground('#DDEBF7');
      sh.getRange(rowNum, 1).setFontWeight('bold').setFontColor(priorityColor(outRows[i].cells[4]));
      sh.getRange(rowNum, 2).setFontWeight('bold').setFontColor('#1F4E78');
      styleAction(sh, rowNum, outRows[i].level);
      styleMr2Cell(sh, rowNum, outRows[i].cells[8]);
    }
  }

  sh.getRange(startRow, 3, outRows.length, 5).setHorizontalAlignment('center');
  sh.getRange(startRow, 9, outRows.length, 1).setHorizontalAlignment('center');

  for (var b = 0; b < aircraftBandRows.length; b++) {
    var endRow = (b + 1 < aircraftBandRows.length) ? aircraftBandRows[b + 1] - 1 : startRow + outRows.length - 1;
    sh.getRange(endRow, 1, 1, nCols).setBorder(null, null, true, null, null, null, '#1F4E78', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  var widths = [330, 65, 150, 50, 65, 80, 130, 220, 100];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);

  sh.setFrozenRows(startRow);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}

function priorityColor(prio) {
  var p = String(prio || '').trim().toLowerCase();
  if (p === 'high') return '#C00000';
  if (p === 'medium') return '#B8860B';
  return '#1F4E78';
}

function styleMr2Cell(sh, rowNum, val) {
  var cell = sh.getRange(rowNum, 9);
  var v = String(val || '').trim().toUpperCase();
  if (v === 'MR2') {
    cell.setBackground('#FCE4D6').setFontColor('#9C5700').setFontWeight('bold');
  } else {
    cell.setBackground(null);
  }
}

function styleAction(sh, rowNum, level) {
  var fills = { urgent: '#F8696B', high: '#FFC000', watch: '#FFEB9C', low: '#E2EFDA' };
  var fonts = { urgent: '#FFFFFF', high: '#7F6000', watch: '#9C6500', low: '#375623' };
  var cell = sh.getRange(rowNum, 8);
  cell.setBackground(fills[level] || '#E2EFDA')
      .setFontColor(fonts[level] || '#375623')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
}

// ======================= FILE MGMT =======================
function moveToProcessed(folder, files) {
  var processed = getOrCreateSubfolder(folder, CONFIG.PROCESSED_SUBFOLDER);
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    processed.addFile(f);
    folder.removeFile(f);
  }
}

function getOrCreateSubfolder(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// ======================= EMAIL SUMMARY (JAM 17:00) =======================
function setupDailyEmailTrigger() {
  var ui = SpreadsheetApp.getUi();
  removeEmailTriggers();
  ScriptApp.newTrigger('sendScreeningSummaryEmail')
    .timeBased()
    .atHour(CONFIG.EMAIL_TRIGGER_HOUR)
    .nearMinute(0)
    .everyDays(1)
    .create();
  ui.alert('Trigger Aktif',
    'Email summary akan otomatis terkirim setiap hari sekitar jam ' + CONFIG.EMAIL_TRIGGER_HOUR + ':00.\n\n' +
    'PENTING: pastikan CONFIG.EMAIL_RECIPIENTS sudah diisi dengan email yang benar sebelum jam kirim tiba, ' +
    'kalau tidak email tidak akan terkirim (safety check aktif).',
    ui.ButtonSet.OK);
}

function removeEmailTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendScreeningSummaryEmail') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  if (removed > 0) {
    SpreadsheetApp.getUi().alert('Trigger email dihapus (' + removed + ' trigger).');
  }
}

function sendScreeningSummaryEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
  if (!sh) {
    Logger.log('Sheet "' + CONFIG.OUTPUT_SHEET + '" tidak ditemukan. Email dibatalkan.');
    return;
  }

  var recipients = (CONFIG.EMAIL_RECIPIENTS || []).filter(function(e) {
    return e && e.indexOf('isi_email') === -1;
  });
  if (recipients.length === 0) {
    Logger.log('CONFIG.EMAIL_RECIPIENTS belum diisi dengan email valid. Email TIDAK dikirim.');
    return;
  }

  var lastRow = sh.getLastRow();
  var startRow = 5;
  if (lastRow < startRow) {
    Logger.log('Tidak ada data di sheet Screening MR2.');
    return;
  }

  var nCols = 9;
  var values = sh.getRange(startRow, 1, lastRow - startRow + 1, nCols).getValues();

  var now = new Date();
  var cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  cutoff.setDate(cutoff.getDate() - 3);

  var currentAc = '';
  var perAc = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var col1 = String(row[0]);

    if (col1.indexOf('✈') === 0) {
      currentAc = col1.replace('✈', '').trim();
      continue;
    }
    if (!currentAc) continue;

    var isWarn = col1.indexOf('⚠ WARNING:') === 0;
    var isFault = col1.indexOf('● FAULT:') === 0;
    if (!isWarn && !isFault) continue;

    var itemName = col1.replace(/^⚠ WARNING:\s*|^● FAULT:\s*/, '').trim();

    if (isExcludedMessage(itemName)) continue;

    var count = Number(row[3]) || 0;
    var lastOccRaw = row[6];
    var lastOccDate = (lastOccRaw instanceof Date) ? lastOccRaw : parseFormattedDate(lastOccRaw);
    var lastOccDisplay = (lastOccRaw instanceof Date)
      ? Utilities.formatDate(lastOccRaw, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
      : String(lastOccRaw);

    var isRepetitive = count >= 2;
    var isCurrent = lastOccDate && lastOccDate >= cutoff;

    if (!isRepetitive || !isCurrent) continue;

    var isToday = lastOccDate && sameDate(lastOccDate, now);

    if (!perAc[currentAc]) perAc[currentAc] = [];
    perAc[currentAc].push({
      type: isWarn ? 'WARNING' : 'FAULT',
      item: itemName,
      ata: row[1],
      count: count,
      lastOccurred: lastOccDisplay,
      lastOccDate: lastOccDate,
      isToday: isToday,
      action: String(row[7]).trim(),
      mr2: String(row[8]).trim()
    });
  }

  var acList = Object.keys(perAc);
  if (acList.length === 0) {
    Logger.log('Tidak ada item repetitive & current dalam 3 hari terakhir. Email tidak dikirim.');
    return;
  }

  var html = buildRepetitiveAlertHtml(perAc);
  var today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy');
  var subject = 'ELEVADE Monitoring per ' + today;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    htmlBody: html
  });

  Logger.log('Email repetitive alert terkirim ke: ' + recipients.join(', '));
}

function parseFormattedDate(s) {
  if (!s) return null;
  var months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  var m = String(s).trim().match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\s+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  var mon = months[m[2]];
  if (mon === undefined) return null;
  return new Date(parseInt(m[3],10), mon, parseInt(m[1],10), parseInt(m[4],10), parseInt(m[5],10));
}

function isExcludedMessage(itemName) {
  var list = CONFIG.EMAIL_EXCLUDE_MESSAGES || [];
  var name = String(itemName).toUpperCase();
  for (var i = 0; i < list.length; i++) {
    if (name.indexOf(String(list[i]).toUpperCase()) !== -1) return true;
  }
  return false;
}

function sortItemsForEmail(items) {
  return items.sort(function(a, b) {
    var rankA = (a.type === 'WARNING') ? 0 : 1;
    var rankB = (b.type === 'WARNING') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;

    var ta = a.lastOccDate ? a.lastOccDate.getTime() : 0;
    var tb = b.lastOccDate ? b.lastOccDate.getTime() : 0;
    return tb - ta;
  });
}

function formatAcReg(ac) {
  var t = String(ac || '').trim().toUpperCase();
  if (t.indexOf('PK-') === 0) return t;
  if (t.indexOf('PK') === 0) return 'PK-' + t.slice(2);
  return 'PK-' + t;
}

function buildRepetitiveAlertHtml(perAc) {
  var acList = Object.keys(perAc).sort();
  var totalItems = 0;
  var totalToday = 0;
  var bodyHtml = '';

  for (var i = 0; i < acList.length; i++) {
    var ac = acList[i];
    var items = sortItemsForEmail(perAc[ac]);
    totalItems += items.length;

    bodyHtml +=
      '<div style="margin:0 0 22px 0;">' +
      '<div style="background:#1F4E78;color:#ffffff;padding:8px 12px;font-weight:bold;font-size:14px;letter-spacing:0.3px;">' +
        escapeHtml(formatAcReg(ac)) +
      '</div>' +
      '<table style="border-collapse:collapse;width:100%;font-size:12px;">' +
      '<tr style="background:#F2F5F9;color:#1F4E78;">' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:left;width:80px;">Type</th>' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:left;">Defect / Message</th>' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;width:70px;">ATA</th>' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;width:55px;">Freq</th>' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;width:120px;">Last Occurred</th>' +
        '<th style="padding:7px 8px;border:1px solid #D9E1EC;text-align:left;width:190px;">Recommended Action</th>' +
      '</tr>';

    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      if (it.isToday) totalToday++;

      var typeColor = (it.type === 'WARNING') ? '#C00000' : '#1F4E78';
      var rowBg = it.isToday ? '#FFF6E5' : (j % 2 === 0 ? '#FFFFFF' : '#FAFBFD');
      var boldStyle = it.isToday ? 'font-weight:bold;' : '';
      var todayTag = it.isToday
        ? ' <span style="background:#C00000;color:#ffffff;font-size:10px;padding:1px 5px;border-radius:2px;">TODAY</span>'
        : '';
      var mr2Tag = (it.mr2 === 'MR2')
        ? ' <span style="background:#FCE4D6;color:#9C5700;font-size:10px;padding:1px 5px;border-radius:2px;font-weight:bold;">MR2 OPEN</span>'
        : '';

      bodyHtml +=
        '<tr style="background:' + rowBg + ';">' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;color:' + typeColor + ';font-weight:bold;">' + it.type + '</td>' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;' + boldStyle + '">' + escapeHtml(it.item) + todayTag + mr2Tag + '</td>' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;">' + escapeHtml(String(it.ata)) + '</td>' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;' + boldStyle + '">' + it.count + 'x</td>' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;text-align:center;' + boldStyle + '">' + escapeHtml(it.lastOccurred) + '</td>' +
          '<td style="padding:7px 8px;border:1px solid #D9E1EC;">' + escapeHtml(it.action) + '</td>' +
        '</tr>';
    }
    bodyHtml += '</table></div>';
  }

  var now = new Date();
  var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy');
  var snapshotStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');

  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;max-width:900px;">' +

    '<div style="border-left:5px solid #C00000;padding:2px 0 2px 12px;margin-bottom:14px;">' +
      '<div style="font-size:18px;font-weight:bold;color:#1F4E78;">ELEVADE Monitoring per ' + todayStr + '</div>' +
      '<div style="color:#666666;font-size:12px;margin-top:3px;">Snapshot ' + snapshotStr + ' &nbsp;|&nbsp; Maintenance Operations Center</div>' +
    '</div>' +

    '<div style="color:#000000;font-weight:bold;margin-bottom:14px;">Dear IAA MOC,</div>' +

    '<div style="margin-bottom:16px;font-size:13px;">' +
      '<span style="display:inline-block;background:#1F4E78;color:#ffffff;padding:6px 12px;margin-right:6px;">Total Items: <b>' + totalItems + '</b></span>' +
      '<span style="display:inline-block;background:#C00000;color:#ffffff;padding:6px 12px;margin-right:6px;">Occurred Today: <b>' + totalToday + '</b></span>' +
      '<span style="display:inline-block;background:#5A6B80;color:#ffffff;padding:6px 12px;">Aircraft Affected: <b>' + acList.length + '</b></span>' +
    '</div>' +

    bodyHtml +

    '<div style="border-top:1px solid #D9E1EC;padding-top:10px;margin-top:6px;color:#999999;font-size:11px;line-height:1.6;">' +
      'Automated email from the ELEVADE Fleet Monitoring system (sheet "Screening MR2"), sent daily at ' + CONFIG.EMAIL_TRIGGER_HOUR + ':00 WIB.<br>' +
      'Data sourced from ACARS warning &amp; fault messages. For full details per aircraft, see the "Screening MR2" sheet.' +
    '</div>' +

  '</div>';
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function sendScreeningSummaryEmailToMe() {
  var myEmail = Session.getActiveUser().getEmail();
  var originalRecipients = CONFIG.EMAIL_RECIPIENTS;
  CONFIG.EMAIL_RECIPIENTS = [myEmail];
  sendScreeningSummaryEmail();
  CONFIG.EMAIL_RECIPIENTS = originalRecipients;
  SpreadsheetApp.getUi().alert('Email dikirim ke: ' + myEmail, ui.ButtonSet.OK);
}
