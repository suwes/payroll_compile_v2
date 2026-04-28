// ============================================================
//  PAYROLL COMPILE AUTOMATION — Google Apps Script
//  Versi: 3.1 — Surgical Fix
//
//  PERUBAHAN DARI v3.0 (hanya 2 area, logic bisnis tidak diubah):
//
//  [FIX 1] SHEET FILTERING — processXlsxFile()
//    • Ditambahkan SHEET_NAME_MAP: payrollGroup → [sheet names yg boleh diproses]
//    • processXlsxFile() sekarang HANYA iterasi sheet yang ada di mapping,
//      bukan semua sheet dalam file xlsx.
//    • Fungsi baru: getSheetNamesForGroup(), loadSheetNamesFromTSV()
//    • CONFIG baru: MAPPING_TSV_FILE_ID (opsional — baca dari file TSV di Drive)
//
//  [FIX 2] TIMEOUT CONTINUATION — runPayrollCompile()
//    • Ditambahkan pengecekan elapsed time di setiap iterasi file
//    • Jika mendekati batas 30 menit (threshold 25 menit), state disimpan
//      ke PropertiesService lalu trigger dijadwalkan 1 menit kemudian
//    • continuePayrollCompile() melanjutkan dari file terakhir yang diproses
//    • Fungsi baru: continuePayrollCompile(), saveCompileState(),
//      loadCompileState(), clearCompileState(), scheduleContinuation(),
//      finalizeCompile()
//
//  TIDAK DIUBAH (dijamin tidak mengubah hasil compile):
//    • Seluruh COLUMN_MAPPING
//    • extractFromSheet(), buildRecord(), buildColumnMap()
//    • detectHeaderRow(), detectSubHeaderRow()
//    • isSummaryOrEmptyRow(), applyDefaults()
//    • applyCampaignIncLogic()
//    • batchWriteToSheet() — sudah optimal
//    • NUMERIC_COLS, OUTPUT_COLUMNS
//    • classifyPayrollGroup()
//    • writeToMasterLog(), autoFormatOutputSheet()
// ============================================================

// ─────────────────────────────────────────────────────────────
//  KONFIGURASI
// ─────────────────────────────────────────────────────────────
var CONFIG = {
  PAYROLL_FOLDER_ID:     "1cjnE3L5Y7niFP531NAkgz1NrJx52Woi1",
  MASTER_SHEET_NAME:     "Master",
  OUTPUT_FOLDER_PREFIX:  "PAYROLL COMPILE FIX",
  OUTPUT_FILE_PREFIX:    "Payroll Compile",
  MAX_HEADER_SCAN_ROWS:  25,
  MAX_HEADER_SCAN_COLS:  80,
  BATCH_SIZE:            2000,
  SLEEP_BETWEEN_FILES_MS: 1000,

  // [FIX 1] ID file mapping TSV di Google Drive (opsional).
  // Jika diisi, script membaca sheet name dari TSV ini.
  // Jika kosong (""), script memakai SHEET_NAME_MAP hardcoded di bawah.
  // Format TSV: kolom 1 = Nama File, kolom 2 = Group Payroll, kolom 3 = Sheet Name
  MAPPING_TSV_FILE_ID:   "",

  // [FIX 2] Batas aman eksekusi sebelum continuation dijadwalkan (ms)
  // Workspace limit = 30 menit. Kita berhenti di 25 menit untuk safety.
  SAFE_EXEC_TIME_MS:     25 * 60 * 1000,
};

// ─────────────────────────────────────────────────────────────
//  [FIX 1] SHEET_NAME_MAP — Hardcoded dari file mapping TSV
//  payrollGroup (output classifyPayrollGroup) → [sheet names yg valid]
//
//  Catatan: Jika ada lebih dari satu sheet name, script akan mencoba
//  setiap nama secara berurutan dan memproses SEMUA yang ditemukan.
//  Contoh: Daily Operator bisa pakai "Calculate" ATAU "Sheet1"
//  tergantung jenis file (regular vs rapel).
//
//  ⚠ JANGAN UBAH KEY — harus sama persis dengan output classifyPayrollGroup()
// ─────────────────────────────────────────────────────────────
var SHEET_NAME_MAP = {
  "Regular Operator":       ["Payroll"],
  "Regular Courier":        ["Summary"],
  "Regular Courier Plus":   ["Summary"],
  "Campaign Inc":           ["Campaign Incentive"],
  "Monthly Incentive Plus": ["Calculate","Calculate Bersih"],
  "Monthly Incentive":      ["Calculate","Calculate Bersih"],
  "Daily Courier Biweekly": ["Calculate Bersih","Calculate"],
  "Daily Courier Weekly":   ["Calculate","Calculate Bersih"],
  "Daily Courier":          ["Calculate", "Calculate Bersih"], // fallback
  "Daily Operator":         ["Calculate", ,"Calculate Bersih","Sheet1"],           // rapel pakai Sheet1
  "Unknown":                [],                                // skip
};

// Cache: diisi saat runtime (oleh loadSheetNamesFromTSV atau pakai hardcoded)
var _SHEET_NAME_MAP_RUNTIME = null;

// ─────────────────────────────────────────────────────────────
//  URUTAN KOLOM OUTPUT — TIDAK DIUBAH
// ─────────────────────────────────────────────────────────────
var OUTPUT_COLUMNS = [
  "periode","payroll_group","source_file",
  "id","name","work_location","role","no_telepon","nik","email",
  "no_bpjs_tk","no_bpjs_kes","total_day",
  "commission","adjustment_plus","incentive","additional","allowance",
  "alpa_deduction","subtotal_1",
  "adjustment_minus",
  "bpjs_tk","bpjs_kes","bpjs_jkk","bpjs_jkm","bpjs_jht","subtotal_2",
  "bpjs_tk_emp","bpjs_kes_emp","bpjs_jkk_emp","bpjs_jkm_emp","bpjs_jht_emp",
  "total_pot_bpjs_karyawan","total_pot_bpjs_tk","total_pot_bpjs_kes",
  "cod_live","lnd_live","fraud_live",
  "cod","lnd","fraud","total_fraud",
  "fraud_loss_sebelumnya","fraud_hold_potong/collect",
  "kasbon","biaya_aplikasi","biaya_perlengkapan","mess","biaya_biztrip",
  "potongan_kekurangan_bpjs","admin","pph21","total_potongan",
  "thp",
  "nama_rekening","no_rekening","bank_name","status_rek",
  "status_pembayaran","tanggal_pembayaran","id_transaksi","bukti_tf",
  "overpaid","double_transfer","alasan_pending"
];

// TIDAK DIUBAH
var NUMERIC_COLS = {
  "total_day":1,"commission":1,"adjustment_plus":1,"incentive":1,"additional":1,
  "allowance":1,"alpa_deduction":1,"subtotal_1":1,"adjustment_minus":1,
  "bpjs_tk":1,"bpjs_kes":1,"bpjs_jkk":1,"bpjs_jkm":1,"bpjs_jht":1,"subtotal_2":1,
  "bpjs_tk_emp":1,"bpjs_kes_emp":1,"bpjs_jkk_emp":1,"bpjs_jkm_emp":1,"bpjs_jht_emp":1,
  "total_pot_bpjs_karyawan":1,"total_pot_bpjs_tk":1,"total_pot_bpjs_kes":1,
  "cod_live":1,"lnd_live":1,"fraud_live":1,
  "cod":1,"lnd":1,"fraud":1,"total_fraud":1,
  "fraud_loss_sebelumnya":1,"fraud_hold_potong/collect":1,
  "kasbon":1,"biaya_aplikasi":1,"biaya_perlengkapan":1,"mess":1,"biaya_biztrip":1,
  "potongan_kekurangan_bpjs":1,"admin":1,"pph21":1,"total_potongan":1,"thp":1,
  "overpaid":1,"double_transfer":1
};

// ─────────────────────────────────────────────────────────────
//  COLUMN MAPPING — TIDAK DIUBAH
// ─────────────────────────────────────────────────────────────
var COLUMN_MAPPING = {
  "id":                       ["id","id driver","os id","id rider","opsid","driver id"],
  "name":                     ["name","nama","courier name","staff name","driver name"],
  "work_location":            ["location","hub","station","position"],
  "role":                     ["title","tittle","role","service"],
  "no_telepon":               ["no hp","phone no","no handphone"],
  "nik":                      ["nik ktp","nik"],
  "email":                    ["email"],
  "no_bpjs_tk":               ["no. bpjs tk"],
  "no_bpjs_kes":              ["no. bpjs kes"],
  "total_day":                ["total day","total days","attandance","day","absensi","calendar day"],
  "commission":               ["commission","basic commission","basic salary","rate/day",
                               "total via bpo","total","total basic salary"],
  "adjustment_plus":          ["adjustment (+)","adjustment +","adjustment","adjustment hk",
                               "adjustment rate","performance bonus","attendance adjustment"],
  "incentive":                ["attendance incentive","campaign incentive","campaign incentive (daily)",
                               "incentive mitra daily","incentive mitra - meal & fuel",
                               "incentive mitra - full rooster","incentive mitra - legacy closure",
                               "incentive mitra - hybrid pickup","incentive mitra - force majeur",
                               "incentive mitra - ramadan","incentive mitra - post ramadhan",
                               "incentive mitra - post ramadan",
                               "additional incentives pre-ramadhan-post",
                               "incentive trip lh","incentive lh","incentive fm",
                               "incentive lm prorate","incentive district",
                               "incentive performance cache","hypercare capacity incentives",
                               "relocation incentive","uang kehadiran","bonus hub","bonus quarterly",
                               "ramadan incentive","hypercare incentive","biztrip incentive",
                               "incentive mitra - campaign incentive",
                               "incentive mitra - campaign incenitive"],
  "additional":               ["additional commission","additional commision",
                               "additional incentive","additional commission point"],
  "allowance":                ["allowance","relocation allowance","performance allowance",
                               "atd allowance linehaul","allowance 2","allowance 3","allowance 4"],
  "alpa_deduction":           ["alpa deduction","deduction alpa + penalty"],
  "subtotal_1":               ["subtotal i","subtotal 1","subtotal","total basic commission"],
  "adjustment_minus":         ["adjustment (-)","adjustment-","adjustment -",
                               "adjustment fraud","adjustment bpjs","potongan h+2"],
  "bpjs_tk":                  ["bpjs tk"],
  "bpjs_kes":                 ["bpjs kes"],
  "bpjs_jkk":                 ["jkk (bpjs tk bpu)"],
  "bpjs_jkm":                 ["jkm (bpjs tk bpu)"],
  "bpjs_jht":                 ["bpjs pension","jht 2 %"],
  "subtotal_2":               ["subtotal ii","subtotal 2"],
  "total_pot_bpjs_karyawan":  ["total bpjs yang dipotong","total potongan bpjs karyawan"],
  "total_pot_bpjs_tk":        ["potongan bpjs tk","bpjs tk [jkk & jkm 0.54%]"],
  "total_pot_bpjs_kes":       ["potongan bpjs kesehatan"],
  "potongan_kekurangan_bpjs": ["kekurangan iuran bpjs tk","kekurangan iuran bpjs kesehatan",
                               "sisa kekurangan bpjs"],
  "cod_live":                 ["potongan cod live","cod live","cod (live)"],
  "lnd_live":                 ["potongan l&d live","l&d live","lost&damage (live)"],
  "cod":                      ["potongan cod","cod (pot invoice)","cod (invoice)","potongan cod invoice"],
  "lnd":                      ["potongan l&d","lost/damage","lost&damage","lost&damage (pot invoice)",
                               "lost&damage (invoice)","potongan l&d invoice"],
  "fraud_loss_sebelumnya":    ["jumlah fraud loss sebelumnya","jumlah final loss fraud sebelumnya",
                               "fraud periode sebelumnya","final loss by pic",
                               "potongan final loss","fraud loss"],
  "fraud_hold_potong/collect":["total fraud yang di hold/potong","fraud yang di hold/potong",
                               "total fraud yg dihold/dipotong","fraud yang tercollect"],
  "kasbon":                   ["kasbon","kasbon yang tercollect","potongan cashbond"],
  "biaya_aplikasi":           ["biaya aplikasi","aplikasi"],
  "biaya_perlengkapan":       ["perlengkapan","biaya perlengkapan yang tercollect",
                               "biaya perlengkapan","potongan biaya perlengkapan"],
  "mess":                     ["biaya mess","biaya mess yang tercollect"],
  "biaya_biztrip":            ["potongan biaya biztrip"],
  "admin":                    ["admin","adm"],
  "pph21":                    ["pph 21","pph21"],
  "overpaid":                 ["overpaid by spx","overpaid by spx (tercollect)","sisa overpaid"],
  "double_transfer":          ["double tf","potongan double tf","double tf loss"],
  "total_potongan":           ["total potongan","total deduction"],
  "thp":                      ["thp","net salary"],
  "no_rekening":              ["nomor rekening","no rekening"],
  "nama_rekening":            ["nama pemilik rekening","nama rekening"],
  "bank_name":                ["bank","nama bank","name bank"],
  "status_rek":               ["status rek"],
  "status_pembayaran":        ["status pembayaran","status tf"],
  "tanggal_pembayaran":       ["tanggal pembayaran","tanggal transfer","tanggal tf"],
  "id_transaksi":             ["id transaksi"],
  "bukti_tf":                 ["bukti tf","bukti transfer","bukti tranfer"],
  "alasan_pending":           ["alasan belum dibayarkan","alasan pembayaran",
                               "alasan belum terbayarkan"],
};

// Alias lookup cache — TIDAK DIUBAH
var _ALIAS_LOOKUP = null;
function getAliasLookup() {
  if (_ALIAS_LOOKUP) return _ALIAS_LOOKUP;
  _ALIAS_LOOKUP = {};
  for (var field in COLUMN_MAPPING) {
    var aliases = COLUMN_MAPPING[field];
    for (var i = 0; i < aliases.length; i++) {
      _ALIAS_LOOKUP[aliases[i].toLowerCase().trim()] = field;
    }
  }
  return _ALIAS_LOOKUP;
}

// ─────────────────────────────────────────────────────────────
//  KLASIFIKASI payroll_group — TIDAK DIUBAH
// ─────────────────────────────────────────────────────────────
function classifyPayrollGroup(fileName) {
  var f = fileName.toLowerCase();
  if (f.includes("campaign inc") || f.includes("campaign incentive")) return "Campaign Inc";
  if (f.includes("monthly incentive") && f.includes("plus"))           return "Monthly Incentive Plus";
  if (f.includes("monthly incentive"))                                  return "Monthly Incentive";
  if (f.includes("dedicated rider plus"))                               return "Regular Courier Plus";
  if (f.includes("dedicated rider"))                                    return "Regular Courier";
  if (f.includes("daily courier biweekly"))                             return "Daily Courier Biweekly";
  if (f.includes("daily courier weekly"))                               return "Daily Courier Weekly";
  if (f.includes("daily courier"))                                      return "Daily Courier";
  if ((f.includes("rapel") || f.includes("rekap")) && f.includes("daily operator")) return "Daily Operator";
  if (f.includes("daily operator"))                                     return "Daily Operator";
  if (f.includes("reguler operator") || f.includes("regular operator")) return "Regular Operator";
  if (f.includes("operator"))                                           return "Regular Operator";
  return "Unknown";
}

// ─────────────────────────────────────────────────────────────
//  [FIX 1] FUNGSI BARU: getSheetNamesForGroup()
//  Mengembalikan array nama sheet yang boleh diproses untuk
//  payrollGroup tertentu. Cek runtime map dulu, fallback ke hardcoded.
// ─────────────────────────────────────────────────────────────
function getSheetNamesForGroup(payrollGroup) {
  var map = _SHEET_NAME_MAP_RUNTIME || SHEET_NAME_MAP;
  return map[payrollGroup] || [];
}

// ─────────────────────────────────────────────────────────────
//  [FIX 1] FUNGSI BARU: loadSheetNamesFromTSV()
//  Membaca file TSV dari Drive dan membangun runtime SHEET_NAME_MAP.
//  Format TSV: kolom[0]=Nama File, kolom[1]=Group Payroll, kolom[2]=Sheet Name
//  Baris pertama dianggap header dan di-skip.
//
//  Hanya baris dengan kolom[1] (Group) tidak kosong yang diproses,
//  karena dalam format mapping TSV ini, banyak baris detail yang
//  kolom Group-nya kosong (inherit dari baris file di atasnya).
//
//  Return: object { "Group": ["Sheet1", "Sheet2", ...], ... }
//  atau null jika gagal / tidak dikonfigurasi.
// ─────────────────────────────────────────────────────────────
function loadSheetNamesFromTSV(fileId) {
  if (!fileId) return null;
  try {
    var content    = DriveApp.getFileById(fileId).getBlob().getDataAsString("UTF-8");
    // Support TSV (\t) dan CSV (,) — deteksi otomatis berdasarkan baris pertama
    var firstLine  = content.split("\n")[0];
    var delimiter  = firstLine.indexOf("\t") !== -1 ? "\t" : ",";

    var lines  = content.split("\n");
    var result = {};

    for (var i = 1; i < lines.length; i++) { // skip header row
      var line = lines[i].trim();
      if (!line) continue;

      // Parsing sederhana — handle quoted fields untuk CSV
      var cols = parseCsvRow(line, delimiter);
      if (cols.length < 3) continue;

      var group     = (cols[1] || "").trim();
      var sheetName = (cols[2] || "").trim();

      // Hanya proses baris dengan Group dan SheetName tidak kosong
      if (!group || !sheetName) continue;

      if (!result[group]) result[group] = [];
      if (result[group].indexOf(sheetName) === -1) {
        result[group].push(sheetName);
      }
    }

    Logger.log("TSV loaded. Groups: " + Object.keys(result).join(", "));
    return result;
  } catch (e) {
    Logger.log("loadSheetNamesFromTSV error: " + e.message + " — pakai hardcoded map.");
    return null;
  }
}

// Helper: parse satu baris CSV dengan support quoted fields
function parseCsvRow(line, delimiter) {
  if (delimiter === "\t") return line.split("\t");
  // Untuk CSV biasa dengan kemungkinan quoted field
  var result = [];
  var cur    = "";
  var inQuote = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') { inQuote = !inQuote; continue; }
    if (ch === delimiter && !inQuote) { result.push(cur); cur = ""; continue; }
    cur += ch;
  }
  result.push(cur);
  return result;
}

// ─────────────────────────────────────────────────────────────
//  UI MENU — Ditambah item "Lanjutkan Compile"
// ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Payroll Compile")
    .addItem("▶ Jalankan Kompilasi Periode...", "showPeriodeDialog")
    .addItem("⏩ Lanjutkan Compile (jika timeout)", "continuePayrollCompile") // [FIX 2]
    .addSeparator()
    .addItem("📋 Lihat Log Master", "showMasterSheet")
    .addItem("⚙ Info Konfigurasi", "showConfigInfo")
    .addToUi();
}

function showPeriodeDialog() {
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "📦 Payroll Compile",
    "Masukkan nama periode (contoh: Maret 2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var periode = result.getResponseText().trim();
  if (!periode) { ui.alert("⚠ Nama periode tidak boleh kosong."); return; }

  var konfirmasi = ui.alert(
    "Konfirmasi",
    'Compile payroll untuk periode: "' + periode + '"?\n' +
    'Proses mungkin memerlukan beberapa sesi jika data sangat besar.',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi === ui.Button.YES) runPayrollCompile(periode);
}

function showMasterSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (sheet) ss.setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('Sheet "' + CONFIG.MASTER_SHEET_NAME + '" belum ada.');
}

function showConfigInfo() {
  SpreadsheetApp.getUi().alert(
    "⚙ Konfigurasi",
    "PAYROLL_FOLDER_ID   : " + CONFIG.PAYROLL_FOLDER_ID +
    "\nMAPPING_TSV_FILE_ID : " + (CONFIG.MAPPING_TSV_FILE_ID || "(tidak diset — pakai hardcoded)"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] FUNGSI UTAMA — Dengan time-aware continuation
// ─────────────────────────────────────────────────────────────
function runPayrollCompile(periode) {
  var ui        = SpreadsheetApp.getUi();
  var startTime = new Date();

  // Pastikan tidak ada sesi compile lain yang sedang berjalan
  var existingState = loadCompileState();
  if (existingState) {
    var resume = ui.alert(
      "⚠ Sesi Sebelumnya Ditemukan",
      'Ada sesi compile yang belum selesai untuk periode "' + existingState.periode + '".\n' +
      'Pilih YES untuk melanjutkan, NO untuk mulai ulang dari awal.',
      ui.ButtonSet.YES_NO
    );
    if (resume === ui.Button.YES) {
      continuePayrollCompile();
      return;
    } else {
      clearCompileState();
    }
  }

  var logEntry = {
    periode: periode,
    tanggalRun: Utilities.formatDate(startTime, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    jumlahFile: 0, jumlahPayroll: 0,
    linkFile: "", linkFolder: "",
    status: "Failed", errorMessage: "",
    tsStart: startTime.toISOString(), tsSuccess: "",
    durasi: "", lastUpdate: ""
  };

  try {
    Logger.log("=== START Kompilasi: " + periode + " ===");

    // [FIX 1] Load sheet name map dari TSV (jika dikonfigurasi)
    _SHEET_NAME_MAP_RUNTIME = loadSheetNamesFromTSV(CONFIG.MAPPING_TSV_FILE_ID);
    if (_SHEET_NAME_MAP_RUNTIME) {
      Logger.log("Menggunakan sheet name map dari TSV.");
    } else {
      Logger.log("Menggunakan sheet name map hardcoded.");
    }

    // Build alias lookup sekali di awal
    getAliasLookup();

    // 1. Folder PAYROLL utama
    var payrollFolder = DriveApp.getFolderById(CONFIG.PAYROLL_FOLDER_ID);

    // 2. Sub-folder periode
    var pIter = payrollFolder.getFoldersByName(periode);
    if (!pIter.hasNext()) throw new Error('Folder periode "' + periode + '" tidak ditemukan.');
    var periodeFolder = pIter.next();
    logEntry.linkFolder = periodeFolder.getUrl();

    // 3. Daftar file .xlsx — simpan sebagai array ID (untuk PropertiesService)
    var xlsxFiles = getXlsxFiles(periodeFolder);
    if (xlsxFiles.length === 0) throw new Error("Tidak ada file .xlsx di folder " + periode);
    logEntry.jumlahFile = xlsxFiles.length;
    var fileIds = xlsxFiles.map(function(f) { return f.getId(); });
    Logger.log("Ditemukan " + xlsxFiles.length + " file xlsx.");

    // 4. Folder dan file output — dibuat di Drive root (My Drive), bukan di dalam folder PAYROLL
    var tahun        = extractTahun(periode);
    var outputFolder = getOrCreateOutputFolder(tahun);
    var outputSS     = createOutputSpreadsheet(CONFIG.OUTPUT_FILE_PREFIX + " " + periode, outputFolder);
    var outputSheet  = outputSS.getActiveSheet();
    outputSheet.setName("Payroll Compile");

    // Tulis header
    outputSheet.getRange(1, 1, 1, OUTPUT_COLUMNS.length).setValues([OUTPUT_COLUMNS]);
    outputSheet.getRange(1, 1, 1, OUTPUT_COLUMNS.length)
      .setFontWeight("bold").setBackground("#1E3A5F").setFontColor("#FFFFFF");
    outputSheet.setFrozenRows(1);

    // 5. Simpan state awal ke PropertiesService untuk antisipasi timeout
    var state = {
      periode:       periode,
      fileIds:       fileIds,
      currentIndex:  0,
      outputSsId:    outputSS.getId(),
      outputFolderId:outputFolder.getId(),
      totalRows:     0,
      errors:        [],
      logEntry:      logEntry,
      // Catat originalStartTime untuk durasi total yang akurat
      originalStartTime: startTime.toISOString()
    };
    saveCompileState(state);

    // 6. Proses file
    processFileLoop(state, startTime);

  } catch (err) {
    Logger.log("FATAL: " + err.message + "\n" + err.stack);
    logEntry.errorMessage = err.message;
    var et = new Date();
    logEntry.durasi     = ((et - startTime) / 60000).toFixed(2);
    logEntry.lastUpdate = Utilities.formatDate(et, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    logEntry.status     = "Failed";
    writeToMasterLog(logEntry);
    clearCompileState();
    SpreadsheetApp.getUi().alert("❌ Gagal!\n\n" + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] LOOP UTAMA PEMROSESAN FILE
//  Dipisah dari runPayrollCompile() agar bisa dipanggil
//  baik oleh runPayrollCompile() maupun continuePayrollCompile()
// ─────────────────────────────────────────────────────────────
function processFileLoop(state, sessionStartTime) {
  var fileIds     = state.fileIds;
  var startIdx    = state.currentIndex;
  var totalFiles  = fileIds.length;
  var outputSS    = SpreadsheetApp.openById(state.outputSsId);
  var outputSheet = outputSS.getSheetByName("Payroll Compile");

  Logger.log("Melanjutkan dari file index " + startIdx + " / " + totalFiles);

  for (var fi = startIdx; fi < totalFiles; fi++) {
    // ── [FIX 2] Cek elapsed time di awal setiap iterasi file ──────────
    var elapsed = new Date() - sessionStartTime;
    if (elapsed >= CONFIG.SAFE_EXEC_TIME_MS) {
      // Simpan progress dan jadwalkan continuation
      Logger.log("⏱ Mendekati batas waktu di file index " + fi + ". Menjadwalkan continuation...");
      state.currentIndex = fi;
      saveCompileState(state);
      scheduleContinuation();
      SpreadsheetApp.getUi().alert(
        "⏱ Proses dihentikan sementara.\n\n" +
        "Sudah diproses: " + fi + " dari " + totalFiles + " file.\n" +
        "Script akan otomatis lanjut dalam ~1 menit.\n" +
        "Atau klik menu \"⏩ Lanjutkan Compile\" untuk lanjut sekarang."
      );
      return; // Keluar dari fungsi — akan dilanjutkan via trigger
    }
    // ─────────────────────────────────────────────────────────────────

    var fileId   = fileIds[fi];
    var xlsxFile = DriveApp.getFileById(fileId);
    Logger.log("[" + (fi+1) + "/" + totalFiles + "] " + xlsxFile.getName());

    var tempSS = null;
    try {
      var outputFolder = DriveApp.getFolderById(state.outputFolderId);
      tempSS = convertXlsxToTempSheet(xlsxFile, outputFolder);
      if (!tempSS) {
        state.errors.push(xlsxFile.getName() + ": Gagal konversi.");
        continue;
      }

      var rows = processXlsxFile(tempSS, xlsxFile.getName(), state.periode);
      Logger.log("  → " + rows.length + " baris.");

      if (rows.length > 0) {
        batchWriteToSheet(outputSheet, rows);
        state.totalRows += rows.length;
      }

      deleteTempFile(tempSS);
      Utilities.sleep(CONFIG.SLEEP_BETWEEN_FILES_MS);

    } catch (fe) {
      Logger.log("ERROR " + xlsxFile.getName() + ": " + fe.message);
      state.errors.push(xlsxFile.getName() + ": " + fe.message);
      if (tempSS) { try { deleteTempFile(tempSS); } catch(e2) {} }
    }

    // Update state progress setelah tiap file selesai
    state.currentIndex = fi + 1;
    saveCompileState(state);
  }

  // Semua file selesai
  finalizeCompile(state, outputSheet, outputSS);
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] FUNGSI BARU: continuePayrollCompile()
//  Dipanggil oleh time-trigger atau manual dari menu
// ─────────────────────────────────────────────────────────────
function continuePayrollCompile() {
  var state = loadCompileState();
  if (!state) {
    SpreadsheetApp.getUi().alert("ℹ Tidak ada sesi compile yang perlu dilanjutkan.");
    return;
  }

  // Hapus trigger lama (jika ada) sebelum memulai sesi baru
  cancelContinuationTriggers();

  // Re-load alias lookup (hilang karena sesi baru)
  getAliasLookup();
  // Re-load sheet name map dari TSV jika dikonfigurasi
  _SHEET_NAME_MAP_RUNTIME = loadSheetNamesFromTSV(CONFIG.MAPPING_TSV_FILE_ID);

  Logger.log("=== CONTINUE Kompilasi: " + state.periode +
             " — Mulai dari file " + state.currentIndex + "/" + state.fileIds.length + " ===");

  processFileLoop(state, new Date()); // sessionStartTime direset untuk sesi ini
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] FUNGSI BARU: finalizeCompile()
//  Dipanggil setelah semua file selesai diproses
// ─────────────────────────────────────────────────────────────
function finalizeCompile(state, outputSheet, outputSS) {
  Logger.log("=== FINALISASI compile ===");

  autoFormatOutputSheet(outputSheet, state.totalRows);

  var logEntry       = state.logEntry;
  var originalStart  = new Date(state.originalStartTime);
  var endTime        = new Date();

  logEntry.jumlahPayroll = state.totalRows;
  logEntry.linkFile      = outputSS.getUrl();
  logEntry.status        = state.errors.length > 0 ? "Partial Success" : "Success";
  logEntry.errorMessage  = state.errors.join(" | ");
  logEntry.tsSuccess     = endTime.toISOString();
  logEntry.durasi        = ((endTime - originalStart) / 60000).toFixed(2);
  logEntry.lastUpdate    = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  writeToMasterLog(logEntry);
  clearCompileState();
  cancelContinuationTriggers();

  SpreadsheetApp.getUi().alert(
    "✅ Kompilasi Selesai!\n\n" +
    "Periode    : " + logEntry.periode + "\n" +
    "Jumlah File: " + state.fileIds.length + "\n" +
    "Total Data : " + state.totalRows + " baris\n" +
    "Durasi     : " + logEntry.durasi + " menit\n" +
    "Status     : " + logEntry.status +
    (state.errors.length > 0 ? "\n\n⚠ Partial Error:\n" + state.errors.slice(0,3).join("\n") : "")
  );
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] PROPERTYSERVICE — Simpan & load state compile
//  PropertiesService hanya support string, jadi state di-JSON.stringify
// ─────────────────────────────────────────────────────────────
var COMPILE_STATE_KEY = "PAYROLL_COMPILE_STATE";

function saveCompileState(state) {
  PropertiesService.getScriptProperties().setProperty(
    COMPILE_STATE_KEY, JSON.stringify(state)
  );
}

function loadCompileState() {
  var raw = PropertiesService.getScriptProperties().getProperty(COMPILE_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function clearCompileState() {
  PropertiesService.getScriptProperties().deleteProperty(COMPILE_STATE_KEY);
}

// ─────────────────────────────────────────────────────────────
//  [FIX 2] TRIGGER — Jadwalkan continuation 1 menit kemudian
// ─────────────────────────────────────────────────────────────
var TRIGGER_HANDLER = "continuePayrollCompile";

function scheduleContinuation() {
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .after(60 * 1000) // 1 menit
    .create();
  Logger.log("Continuation trigger dijadwalkan.");
}

function cancelContinuationTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("Trigger lama dihapus.");
    }
  }
}

// ─────────────────────────────────────────────────────────────
//  KONVERSI .xlsx → Google Sheets — TIDAK DIUBAH
// ─────────────────────────────────────────────────────────────
function convertXlsxToTempSheet(xlsxFile, outputFolder) {
  var tempName = "__TEMP__" + xlsxFile.getName();
  try {
    var copiedFile = Drive.Files.copy(
      {
        title:    tempName,
        mimeType: "application/vnd.google-apps.spreadsheet",
        parents:  [{ id: outputFolder.getId() }]
      },
      xlsxFile.getId()
    );
    Logger.log("  Konversi OK → ID: " + copiedFile.id);
    Utilities.sleep(2000);
    return SpreadsheetApp.openById(copiedFile.id);
  } catch (e) {
    Logger.log("  Konversi gagal: " + e.message);
    try {
      var leftovers = outputFolder.getFilesByName(tempName);
      while (leftovers.hasNext()) leftovers.next().setTrashed(true);
    } catch(e2) {}
    return null;
  }
}

// ─────────────────────────────────────────────────────────────
//  [FIX 1] processXlsxFile() — DIUBAH:
//  Sebelumnya: iterate semua sheets di tempSS
//  Sekarang  : hanya iterasi sheet yang namanya ada di SHEET_NAME_MAP
// ─────────────────────────────────────────────────────────────
function processXlsxFile(tempSS, fileName, periode) {
  var payrollGroup     = classifyPayrollGroup(fileName);
  var isCampaignInc    = (payrollGroup === "Campaign Inc");
  var allRows          = [];

  // [FIX 1] Ambil daftar sheet name yang diizinkan untuk group ini
  var allowedSheetNames = getSheetNamesForGroup(payrollGroup);

  if (allowedSheetNames.length === 0) {
    // Group "Unknown" atau tidak ada di mapping → skip semua sheet
    Logger.log("  Group '" + payrollGroup + "' tidak ada di mapping → file dilewati.");
    return [];
  }

  Logger.log("  Group: " + payrollGroup + " → Sheet target: [" + allowedSheetNames.join(", ") + "]");

  // [FIX 1] Hanya proses sheet yang namanya ada di allowedSheetNames
  //         (bukan semua sheet di file)
  for (var ai = 0; ai < allowedSheetNames.length; ai++) {
    var targetName = allowedSheetNames[ai];
    var sheet      = tempSS.getSheetByName(targetName);

    if (!sheet) {
      // Sheet tidak ada di file ini — bukan error, beberapa file memang tidak punya sheet tsb
      Logger.log('  Sheet "' + targetName + '" tidak ditemukan di file → skip.');
      continue;
    }

    if (sheet.getLastRow() < 2) {
      Logger.log('  Sheet "' + targetName + '" kosong → skip.');
      continue;
    }

    Logger.log('  Memproses sheet: "' + targetName + '"');
    try {
      var rows = extractFromSheet(sheet, fileName, periode, payrollGroup, isCampaignInc);
      Logger.log('  → ' + rows.length + ' baris dari "' + targetName + '".');
      allRows = allRows.concat(rows);
    } catch(e) {
      Logger.log('  ERROR sheet "' + targetName + '": ' + e.message);
    }
  }

  return allRows;
}

// ─────────────────────────────────────────────────────────────
//  extractFromSheet() — TIDAK DIUBAH (logic bisnis intact)
// ─────────────────────────────────────────────────────────────
function extractFromSheet(sheet, fileName, periode, payrollGroup, isCampaignInc) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var rawData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var headerInfo = detectHeaderRow(rawData, lastRow, lastCol);
  if (!headerInfo) { Logger.log("  Tidak ada header — skip."); return []; }

  var headerRowIdx = headerInfo.headerRowIdx;
  var startCol     = headerInfo.startCol;
  Logger.log("  Header di baris " + (headerRowIdx+1) + ", startCol " + (startCol+1));

  var subHeaderInfo = detectSubHeaderRow(rawData, headerRowIdx, startCol, lastCol);
  var dataOffset    = subHeaderInfo.hasSubHeader ? 2 : 1;
  var colMap        = buildColumnMap(rawData[headerRowIdx], startCol, lastCol, subHeaderInfo, isCampaignInc);

  // Scan ke bawah jika ada baris kosong antara header dan data (fix off-by-one)
  var firstDataRow = headerRowIdx + dataOffset;
  while (firstDataRow < lastRow) {
    var probe  = rawData[firstDataRow];
    var hasVal = false;
    for (var c = startCol; c < Math.min(probe.length, CONFIG.MAX_HEADER_SCAN_COLS); c++) {
      if (String(probe[c] || "").trim() !== "") { hasVal = true; break; }
    }
    if (hasVal) break;
    firstDataRow++;
  }

  var extractedRows = [];
  for (var r = firstDataRow; r < lastRow; r++) {
    var rowData = rawData[r];
    if (isSummaryOrEmptyRow(rowData, colMap, startCol, lastCol)) continue;

    var record     = buildRecord(rowData, colMap, periode, payrollGroup, fileName);
    var hasCodLive = !!colMap["cod_live"];
    var hasLndLive = !!colMap["lnd_live"];
    var hasCod     = !!colMap["cod"];
    var hasLnd     = !!colMap["lnd"];

    if (hasCodLive || hasLndLive) {
      record["fraud_live"] = toNum(record["cod_live"]) + toNum(record["lnd_live"]);
    }
    if (hasCod || hasLnd) {
      record["fraud"] = toNum(record["cod"]) + toNum(record["lnd"]);
    }
    record["total_fraud"] = toNum(record["fraud_live"]) + toNum(record["fraud"]);

    if (isCampaignInc) applyCampaignIncLogic(record);

    applyDefaults(record);
    extractedRows.push(record);
  }

  return extractedRows;
}

// ─────────────────────────────────────────────────────────────
//  SEMUA FUNGSI DI BAWAH INI TIDAK DIUBAH SAMA SEKALI
// ─────────────────────────────────────────────────────────────

var SUB_HEADER_KW = ["invoice","live","pot invoice","lost/damage","lost&damage"];

function detectHeaderRow(rawData, lastRow, lastCol) {
  var lookup   = getAliasLookup();
  var allAlias = Object.keys(lookup);
  var bestScore = 0, bestRowIdx = -1;
  var scanRows = Math.min(lastRow, CONFIG.MAX_HEADER_SCAN_ROWS);
  var scanCols = Math.min(lastCol, CONFIG.MAX_HEADER_SCAN_COLS);

  for (var r = 0; r < scanRows; r++) {
    var row = rawData[r], score = 0;
    for (var c = 0; c < Math.min(row.length, scanCols); c++) {
      var cell = String(row[c] || "").trim().toLowerCase();
      if (!cell) continue;
      for (var ai = 0; ai < allAlias.length; ai++) {
        var kw = allAlias[ai];
        if (kw.length <= 3) { if (cell === kw) { score++; break; } }
        else { if (cell.indexOf(kw) !== -1 || kw.indexOf(cell) !== -1) { score++; break; } }
      }
    }
    if (score > bestScore) { bestScore = score; bestRowIdx = r; }
  }

  if (bestScore < 2) return null;

  var headerRow = rawData[bestRowIdx], startCol = 0;
  for (var c2 = 0; c2 < headerRow.length; c2++) {
    if (String(headerRow[c2] || "").trim() !== "") { startCol = c2; break; }
  }
  return { headerRowIdx: bestRowIdx, startCol: startCol };
}

function detectSubHeaderRow(rawData, headerRowIdx, startCol, lastCol) {
  var result = { hasSubHeader: false, subHeaderMap: {} };
  var subIdx = headerRowIdx + 1;
  if (subIdx >= rawData.length) return result;

  var subRow = rawData[subIdx], count = 0;
  var scanC  = Math.min(subRow.length, CONFIG.MAX_HEADER_SCAN_COLS);
  for (var c = startCol; c < scanC; c++) {
    var val = String(subRow[c] || "").trim().toLowerCase();
    if (!val) continue;
    for (var ki = 0; ki < SUB_HEADER_KW.length; ki++) {
      if (val.indexOf(SUB_HEADER_KW[ki]) !== -1) { count++; result.subHeaderMap[c] = val; break; }
    }
  }
  result.hasSubHeader = (count >= 2);
  return result;
}

function buildColumnMap(headerRow, startCol, lastCol, subHeaderInfo, isCampaignInc) {
  var colMap = {}, lookup = getAliasLookup();
  var scanC  = Math.min(headerRow.length, CONFIG.MAX_HEADER_SCAN_COLS);

  for (var c = startCol; c < scanC; c++) {
    var cell = String(headerRow[c] || "").trim().toLowerCase();
    if (!cell) continue;

    var field = null;
    if (lookup[cell]) { field = lookup[cell]; }
    else {
      for (var kw in lookup) {
        if (kw.length <= 3) continue;
        if (cell.indexOf(kw) !== -1 || kw.indexOf(cell) !== -1) { field = lookup[kw]; break; }
      }
    }
    if (!field) continue;

    if (field === "cod" || field === "lnd") {
      if (subHeaderInfo.hasSubHeader) {
        var subLabel = subHeaderInfo.subHeaderMap[c] || "";
        if (subLabel.indexOf("live") !== -1) {
          field = (field === "cod") ? "cod_live" : "lnd_live";
        }
      }
    }

    if (!colMap[field]) colMap[field] = [];
    colMap[field].push(c);
  }
  return colMap;
}

function isSummaryOrEmptyRow(rowData, colMap, startCol, lastCol) {
  var hasVal = false;
  var scanC  = Math.min(rowData.length, CONFIG.MAX_HEADER_SCAN_COLS);
  for (var c = startCol; c < scanC; c++) {
    if (String(rowData[c] || "").trim() !== "") { hasVal = true; break; }
  }
  if (!hasVal) return true;

  var idVal   = String(getFirstValue(rowData, colMap["id"])   || "").trim();
  var nameVal = String(getFirstValue(rowData, colMap["name"]) || "").trim();
  var HEADER_LABELS = { "no":1,"nama":1,"id":1,"name":1,
    "courier name":1,"staff name":1,"driver id":1,"driver name":1 };
  if (HEADER_LABELS[idVal.toLowerCase()]) return true;
  if (idVal.length <= 2 && !nameVal) return true;
  if (!idVal && !nameVal) return true;
  return false;
}

function buildRecord(rowData, colMap, periode, payrollGroup, fileName) {
  var record = {};
  record["periode"]       = periode;
  record["payroll_group"] = payrollGroup;
  record["source_file"]   = fileName;

  for (var field in COLUMN_MAPPING) {
    if (!colMap[field]) continue;
    if (NUMERIC_COLS[field]) {
      var total = 0, indices = colMap[field];
      for (var i = 0; i < indices.length; i++) total += toNum(rowData[indices[i]]);
      record[field] = total;
    } else {
      record[field] = getFirstValue(rowData, colMap[field]) || "";
    }
  }
  return record;
}

function applyCampaignIncLogic(record) {
  var inc        = toNum(record["incentive"]) || toNum(record["commission"]);
  var totalFraud = toNum(record["total_fraud"]);
  record["commission"] = inc;
  record["subtotal_1"] = inc;
  record["fraud_hold_potong/collect"] = (inc >= totalFraud) ? totalFraud : inc;
}

function applyDefaults(record) {
  for (var i = 0; i < OUTPUT_COLUMNS.length; i++) {
    var col = OUTPUT_COLUMNS[i], val = record[col];
    if (val === undefined || val === null || val === "") {
      record[col] = NUMERIC_COLS[col] ? 0 : "-";
    }
  }
}

function batchWriteToSheet(sheet, rows) {
  var batchSize = CONFIG.BATCH_SIZE;
  for (var i = 0; i < rows.length; i += batchSize) {
    var chunk = rows.slice(i, i + batchSize), matrix = [];
    for (var j = 0; j < chunk.length; j++) {
      var row = chunk[j], arr = [];
      for (var k = 0; k < OUTPUT_COLUMNS.length; k++) arr.push(row[OUTPUT_COLUMNS[k]]);
      matrix.push(arr);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, matrix.length, OUTPUT_COLUMNS.length).setValues(matrix);
    if (rows.length > batchSize) Utilities.sleep(100);
  }
}

function autoFormatOutputSheet(sheet, totalDataRows) {
  if (totalDataRows < 1) return;
  try {
    // Alternating row color — skip jika data terlalu besar (> 30k baris)
    // untuk menghindari timeout saat finalisasi
    if (totalDataRows <= 30000) {
      var colors = [];
      for (var r = 0; r < totalDataRows; r++) {
        colors.push(Array(OUTPUT_COLUMNS.length).fill(r % 2 === 0 ? "#FFFFFF" : "#EEF3FA"));
      }
      sheet.getRange(2, 1, totalDataRows, OUTPUT_COLUMNS.length).setBackgrounds(colors);
    }

    var tglIdx = OUTPUT_COLUMNS.indexOf("tanggal_pembayaran") + 1;
    if (tglIdx > 0) sheet.getRange(2, tglIdx, totalDataRows, 1).setNumberFormat("dd/MM/yyyy");

    for (var col in NUMERIC_COLS) {
      var ci = OUTPUT_COLUMNS.indexOf(col) + 1;
      if (ci > 0) sheet.getRange(2, ci, totalDataRows, 1).setNumberFormat("#,##0");
    }

    if (totalDataRows < 5000) sheet.autoResizeColumns(1, OUTPUT_COLUMNS.length);
  } catch(e) { Logger.log("Format warning: " + e.message); }
}

function writeToMasterLog(logEntry) {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) {
      masterSheet = ss.insertSheet(CONFIG.MASTER_SHEET_NAME);
      var hdrs = ["Periode","Tanggal Run","Jumlah File","Jumlah Payroll",
                  "Link File Google Sheets","Link Folder Periode",
                  "Status Compile","Error Message",
                  "Time Stamp Start","Time Stamp Success","Duration (min)","Last Update"];
      masterSheet.appendRow(hdrs);
      masterSheet.getRange(1,1,1,hdrs.length)
        .setFontWeight("bold").setBackground("#1E3A5F").setFontColor("#FFFFFF");
      masterSheet.setFrozenRows(1);
    }
    masterSheet.appendRow([
      logEntry.periode, logEntry.tanggalRun, logEntry.jumlahFile, logEntry.jumlahPayroll,
      logEntry.linkFile, logEntry.linkFolder, logEntry.status, logEntry.errorMessage,
      logEntry.tsStart, logEntry.tsSuccess, logEntry.durasi, logEntry.lastUpdate
    ]);
    var lastRow    = masterSheet.getLastRow();
    var statusCell = masterSheet.getRange(lastRow, 7);
    if      (logEntry.status === "Success")         statusCell.setBackground("#C6EFCE").setFontColor("#276221");
    else if (logEntry.status === "Partial Success") statusCell.setBackground("#FFEB9C").setFontColor("#9C5700");
    else                                            statusCell.setBackground("#FFC7CE").setFontColor("#9C0006");
  } catch(e) { Logger.log("Master log error: " + e.message); }
}

function getXlsxFiles(folder) {
  var files = [], iter = folder.getFiles();
  var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  while (iter.hasNext()) {
    var f = iter.next();
    if (f.getMimeType() === XLSX_MIME) files.push(f);
  }
  return files;
}

function extractTahun(periode) {
  var m = periode.match(/\d{4}/);
  return m ? m[0] : String(new Date().getFullYear());
}

function getOrCreateOutputFolder(tahun) {
  var name = CONFIG.OUTPUT_FOLDER_PREFIX + " " + tahun;
  // Cari di My Drive (root) — bukan di dalam folder PAYROLL
  var iter = DriveApp.getRootFolder().getFoldersByName(name);
  if (iter.hasNext()) { Logger.log("Folder output sudah ada di My Drive: " + name); return iter.next(); }
  Logger.log("Membuat folder output baru di My Drive: " + name);
  return DriveApp.getRootFolder().createFolder(name);
}

function createOutputSpreadsheet(fileName, folder) {
  var ss = SpreadsheetApp.create(fileName);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch(e) {}
  return ss;
}

function deleteTempFile(tempSS) {
  try { DriveApp.getFileById(tempSS.getId()).setTrashed(true); }
  catch(e) { Logger.log("Gagal hapus temp: " + e.message); }
}

function toNum(val) {
  if (val === null || val === undefined || val === "" || val === "-") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  var n = parseFloat(String(val).replace(/[^\d.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function getFirstValue(rowData, colIndices) {
  if (!colIndices || colIndices.length === 0) return "";
  for (var i = 0; i < colIndices.length; i++) {
    var v = rowData[colIndices[i]];
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}
