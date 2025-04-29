// ====================================================================================
// 📁 Helper: Dropdown, Tanggal, dan Validasi - Resi Automation System
// ====================================================================================


// ------------------------------------------------------------------------
// 🔽 DROPDOWN / DATA VALIDATION
// ------------------------------------------------------------------------

// ✅ Ambil daftar dari sheet tertentu (umum dan khusus periode)
function getListFromSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues().flat();

  if (sheetName === "listPeriode") {
    return data
      .filter(item => item)
      .map(item => {
        if (item instanceof Date) {
          return Utilities.formatDate(item, Session.getScriptTimeZone(), "MMMM yyyy");
        }
        return item.toString();
      });
  }

  return data.filter(item => item).map(item => item.toString());
}

// ✅ Terapkan dropdown list di kolom tertentu
function applyDropdownToColumn(sheet, colIndex, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  const numRows = sheet.getMaxRows();
  sheet.getRange(2, colIndex, numRows - 1).setDataValidation(rule);
}


// ------------------------------------------------------------------------
// 📅 TANGGAL & FORMAT
// ------------------------------------------------------------------------

// ✅ Terapkan date picker hanya ke Kolom L (Tanggal Transaksi)
function applyDatePickerToTanggalTransaksi(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 12, lastRow - 1); // Kolom L
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
}

// ✅ Update format tanggal Indonesia di Kolom M (Tanggal dan Jam Transaksi)
function updateFormattedTanggalTransaksi(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tanggalValues = sheet.getRange(2, 12, lastRow - 1).getValues(); // Kolom L
  const existingFormatted = sheet.getRange(2, 13, lastRow - 1).getValues(); // Kolom M

  const hariIndo = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu"];
  const bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  for (let i = 0; i < tanggalValues.length; i++) {
    const tgl = tanggalValues[i][0];
    const existing = (existingFormatted[i][0] || "").toString().trim();

    const matchJam = existing.match(/(\d{1,2}):(\d{2})$/);
    if (matchJam && matchJam[0] !== "00:00") continue;
    if (!(tgl instanceof Date)) continue;

    const hari = hariIndo[tgl.getDay()];
    const tanggal = tgl.getDate();
    const bulan = bulanIndo[tgl.getMonth()];
    const tahun = tgl.getFullYear();
    const formatted = `${hari}, ${tanggal} ${bulan} ${tahun} 00:00`;

    if (formatted !== existing) {
      sheet.getRange(i + 2, 13).setValue(formatted); // Kolom M
    }
  }
}

// ✅ Sinkronisasi ulang ke kolom M dari formatted string
function syncTanggalJamTransaksiFromFormatted(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const formattedValues = sheet.getRange(2, 13, lastRow - 1).getValues(); // Kolom M
  const targetRange = sheet.getRange(2, 13, lastRow - 1);
  const existingValues = targetRange.getValues();

  for (let i = 0; i < formattedValues.length; i++) {
    const val = formattedValues[i][0];
    const existing = (existingValues[i][0] || "").toString().trim();

    if (val && typeof val === "string" && val.match(/\d{2}:\d{2}$/)) {
      if (existing !== val) {
        targetRange.getCell(i + 1, 1).setValue(val);
      }
    }
  }
}

// ✅ Validasi dan highlight warna pada baris header kolom
function validateAndHighlightHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "Form Responses 1") return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const expectedHeaders = [
    "No", "ID Transaksi", "ID Registrasi", "Email", "Nama Lengkap", "Nomor Telepon",
    "Nama Program", "Session", "Periode Pelaksanaan", "Jumlah Topik yang Diikuti",
    "Jumlah Pembayaran", "Tanggal Transaksi", "Tanggal dan Jam Transaksi", "Metode Pembayaran",
    "Channel Pembayaran",
    "Topik 1", "Topik 2", "Topik 3", "Topik 4", "Topik 5", "Topik 6", "Topik 7", "Topik 8", "Topik 9",
    "Topik 10", "Topik 11", "Topik 12", "Topik 13", "Topik 14", "Topik 15",
    "Nomor Telepon Hashing Otomatis", "Status Resi PDF", "File dalam Folder", "Send Email Status",
    "Keterangan Error", "Folder ID Periode", "Status File di Folder Level 3", "Status Folder Periode di Folder Level 3"
  ];

  for (let i = 0; i < expectedHeaders.length; i++) {
    const actual = (headers[i] || "").toString().trim();
    const expected = expectedHeaders[i];
    const cell = sheet.getRange(1, i + 1);
    cell.setBackground(actual !== expected ? "#f8d7da" : "#d4edda");
  }

  const mismatch = expectedHeaders.some((h, i) => h !== (headers[i] || "").trim());
  if (mismatch) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Urutan header kolom di 'Form Responses 1' tidak sesuai template.\nKolom merah perlu diperiksa."
    );
  }

  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setHorizontalAlignment("center");
  range.setFontWeight("bold");
  range.setFontSize(11);
}

// ✅ Validasi semua topik dan warnai jika tidak sesuai
function validateAllTopicCounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const lastRow = sheet.getLastRow();
  const topicStartCol = 16; // Kolom P
  const topicEndCol = 30;   // Kolom AD

  // ✅ Cari kolom "Keterangan Error" secara dinamis
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.findIndex(h => String(h).toLowerCase().includes("keterangan error")) + 1;


  Logger.log(`🔍 Kolom 'Keterangan Error' ditemukan di kolom nomor: ${statusCol}`);
  if (statusCol === 0) {
    SpreadsheetApp.getUi().alert("❌ Kolom 'Keterangan Error' tidak ditemukan. Pastikan header sudah sesuai.");
    return;
  }

  for (let row = 2; row <= lastRow; row++) {
    validateTopicCount(sheet, row, topicStartCol, topicEndCol, statusCol);
  }
}

function validateTopicCount(sheet, row, topicStartCol, topicEndCol, statusCol) {
  Logger.log(`✅ Menjalankan validasi baris ${row}`);

  const dataRange = sheet.getRange(row, topicStartCol, 1, topicEndCol - topicStartCol + 1);
  const topics = dataRange.getDisplayValues()[0];
  const nonEmptyTopics = topics.filter(t => t.trim() !== "");
  const jumlahTopik = Number(sheet.getRange(row, 10).getValue()); // Kolom J
  const namaLengkap = sheet.getRange(row, 5).getValue(); // Kolom E

  Logger.log(`📊 Baris ${row} - Nama: ${namaLengkap}`);
  Logger.log(`🔢 Jumlah Topik di Kolom J: ${jumlahTopik}`);
  Logger.log(`✏️  Topik terisi: ${nonEmptyTopics.length}`);

  const statusCell = sheet.getRange(row, statusCol);
  const topicCells = sheet.getRange(row, topicStartCol, 1, topicEndCol - topicStartCol + 1);

  if (jumlahTopik && nonEmptyTopics.length > jumlahTopik) {
    statusCell.setValue(`⚠️ Topik melebihi jumlah yang ditentukan untuk ${namaLengkap}`);
    statusCell.setBackground("#f4cccc");
    topicCells.setBackground("#f4cccc");
  } else if (jumlahTopik && nonEmptyTopics.length < jumlahTopik) {
    statusCell.setValue(`⚠️ Topik belum lengkap sesuai jumlah yang ditentukan untuk ${namaLengkap}`);
    statusCell.setBackground("#f4cccc");
    topicCells.setBackground("#f4cccc");
  } else {
    statusCell.clearContent().setBackground(null);

    // 🧹 Reset warna hanya jika tidak ada duplikat (duplikat ditangani terpisah)
    const rawTopics = topics.map(t => t.trim().toLowerCase());
    const duplicates = findDuplicates(rawTopics);
    if (duplicates.length === 0) {
      topicCells.setBackground(null);
    }
  }
}

function forceApplyDatePickerToTanggalTransaksi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const range = sheet.getRange("L2:L1000"); // Sesuaikan batas bawahnya
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  SpreadsheetApp.getUi().alert("✅ Semua kolom L (Tanggal Transaksi) sudah diset sebagai Date Picker!");
}

// ✅ Bersihkan spasi ganda dan trailing space di kolom D (Email), E (Nama), dan F (Telepon)
function cleanEmailNamePhone(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emailRange = sheet.getRange(2, 4, lastRow - 1); // Kolom D
  const nameRange = sheet.getRange(2, 5, lastRow - 1);  // Kolom E
  const phoneRange = sheet.getRange(2, 6, lastRow - 1); // Kolom F

  const emails = emailRange.getValues();
  const names = nameRange.getValues();
  const phones = phoneRange.getValues();

  for (let i = 0; i < lastRow - 1; i++) {
    const cleanedEmail = sanitizeInput(emails[i][0]);
    const cleanedName = sanitizeInput(names[i][0]);
    const cleanedPhone = sanitizeInput(phones[i][0]);

    if (emails[i][0] !== cleanedEmail) emailRange.getCell(i + 1, 1).setValue(cleanedEmail);
    if (names[i][0] !== cleanedName) nameRange.getCell(i + 1, 1).setValue(cleanedName);
    if (phones[i][0] !== cleanedPhone) phoneRange.getCell(i + 1, 1).setValue(cleanedPhone);
  }
}

// 🔧 Utility: Hapus spasi ganda dan trailing
function sanitizeInput(value) {
  if (!value || typeof value !== "string") return value;
  return value.replace(/\s+/g, " ").trim(); // Hapus spasi berlebih dan trailing
}

// ------------------------------------------------------------------------
// ✅ Validasi: Integrity ID Transaksi & Registrasi TANPA kolom snapshot
// ------------------------------------------------------------------------
function validasiIDTrxDanReg(sheet, row = null) {
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];

  const colMap = {
    idTrx: headers.indexOf("ID Transaksi"),
    idReg: headers.indexOf("ID Registrasi"),
    periode: headers.indexOf("Periode Pelaksanaan"),
    session: headers.indexOf("Session"),
    tanggal: headers.indexOf("Tanggal Transaksi"),
    errorNote: headers.indexOf("Keterangan Error")
  };

  const startRow = row ? row : 2;
  const endRow = row ? row : data.length;

  for (let i = startRow; i <= endRow; i++) {
    const rowIndex = i - 1;
    const rowData = data[rowIndex];

    const idTrx = rowData[colMap.idTrx] || "";
    const idReg = rowData[colMap.idReg] || "";

    const periodeNow = rowData[colMap.periode] || "";
    const sessionNow = rowData[colMap.session] || "";
    const tanggalNow = rowData[colMap.tanggal] || "";
    const errorRange = sheet.getRange(i, colMap.errorNote + 1);

    // Reset warna
    sheet.getRange(i, colMap.periode + 1).setBackground(null);
    sheet.getRange(i, colMap.session + 1).setBackground(null);
    sheet.getRange(i, colMap.tanggal + 1).setBackground(null);
    errorRange.setValue("");

    const messages = [];

    const normalizedPeriode = normalizePeriode(periodeNow); // ✅ Normalize ke format Inggris

    // ✅ Ekstrak info dari ID Transaksi jika tersedia
    if (idTrx.startsWith("TC")) {
      const trxPeriode = idTrx.slice(2, 5); // contoh: 425
      const trxSession = idTrx.slice(5, 7); // contoh: 11
      const trxTanggal = idTrx.slice(7, 13); // contoh: 250424

      // → Format periode kolom I menjadi 425
      const periodeMatch = formatPeriodeToCode(periodeNow);
      const sessionMatch = sessionNow.toString().padStart(2, "0");
      const tanggalMatch = formatTanggalToCode(tanggalNow);

      if (periodeMatch && normalizedPeriode !== convertCodeToPeriode(trxPeriode)) {
        messages.push(`Kolom I berubah dari "${convertCodeToPeriode(trxPeriode)}" menjadi "${periodeNow}"`);
        sheet.getRange(i, colMap.periode + 1).setBackground("#f4cccc");
      }

      if (sessionMatch !== trxSession) {
        messages.push(`Kolom H berubah dari "${parseInt(trxSession)}" menjadi "${sessionNow}"`);
        sheet.getRange(i, colMap.session + 1).setBackground("#f4cccc");
      }

      if (tanggalMatch && tanggalMatch !== trxTanggal) {
        messages.push(`Kolom L berubah dari "${convertCodeToTanggal(trxTanggal)}" menjadi "${tanggalNow}"`);
        sheet.getRange(i, colMap.tanggal + 1).setBackground("#f4cccc");
      }
    }

    // ✅ Ekstrak info dari ID Registrasi
    if (idReg.length >= 4) {
      const regPeriode = idReg.slice(0, 3); // contoh: 425
      const periodeMatch = formatPeriodeToCode(periodeNow);
      if (periodeMatch && normalizedPeriode !== convertCodeToPeriode(regPeriode)) {
        messages.push(`Kolom I (untuk ID Registrasi) berubah dari "${convertCodeToPeriode(regPeriode)}" menjadi "${periodeNow}"`);
        sheet.getRange(i, colMap.periode + 1).setBackground("#f4cccc");
      }
    }

    if (messages.length > 0) {
      errorRange.setValue(messages.join(". ") + ". Silakan Generate Ulang ID Transaksi atau kembalikan nilai seperti semula.");
    }
  }
}

// 🔁 Util: Normalize Periode ke format Inggris agar konsisten
function normalizePeriode(val) {
  if (!val || typeof val !== "string") return val;
  const map = {
    Januari: "January", Februari: "February", Maret: "March", April: "April",
    Mei: "May", Juni: "June", Juli: "July", Agustus: "August",
    September: "September", Oktober: "October", November: "November", Desember: "December"
  };
  const parts = val.trim().split(" ");
  if (parts.length !== 2) return val;
  const bulan = map[parts[0]] || parts[0];
  const tahun = parts[1];
  return `${bulan} ${tahun}`;
}

function formatPeriodeToCode(val) {
  if (!val || typeof val !== "string") return null;
  const map = {
    January: "1", February: "2", March: "3", April: "4", May: "5", June: "6",
    July: "7", August: "8", September: "9", October: "10", November: "11", December: "12"
  };
  const parts = normalizePeriode(val).split(" ");
  if (parts.length !== 2) return null;
  const bulan = map[parts[0]] || "0";
  const tahun = parts[1].slice(-2);
  return `${bulan}${tahun}`;
}

function convertCodeToPeriode(code) {
  const bulanMap = ["", "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"];
  const bulan = parseInt(code.slice(0, code.length - 2), 10);
  const tahun = "20" + code.slice(-2);
  return `${bulanMap[bulan]} ${tahun}`;
}

// 🔁 Tanggal → Kode: 24/04/2025 → 250424
function formatTanggalToCode(tglStr) {
  const d = new Date(tglStr);
  if (isNaN(d)) return null;
  const yy = d.getFullYear().toString().slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}${mm}${dd}`;
}

// 🔁 Kode → Tanggal: 250424 → 24/04/2025
function convertCodeToTanggal(code) {
  const yy = "20" + code.slice(0, 2);
  const mm = code.slice(2, 4);
  const dd = code.slice(4, 6);
  return `${dd}/${mm}/${yy}`;
}