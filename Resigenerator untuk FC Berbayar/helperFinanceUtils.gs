// ====================================================================================
// 📁 helperFinanceUtils.gs
// ====================================================================================
// 📌 Fungsi bantu validasi dan formatter untuk email ke Tim Finance
// ====================================================================================

// ✅ Normalisasi Spasi
function normalizeWhitespace(str) {
  return String(str || "").replace(/\s+/g, " ").trim();
}

// ✅ Sanitize Nama File
function sanitize(str) {
  return String(str || "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/–/g, "-")
    .replace(/\s*-\s*/g, " - ")
    .trim();
}

// ✅ Generator Nama File Resi
function generateFileNameResi(row) {
  return `[Receipt TC Phincon Academy] ${sanitize(row.id)} - ${sanitize(row.nama)} - ${sanitize(row.program)} - Session ${sanitize(row.session)} - Lunas`;
}

// ✅ Generator Nama File Bukti Transfer
function generateFileNameBukti(row) {
  return `[Phincon Academy] TC Bukti Transfer - ${sanitize(row.id)} - ${sanitize(row.nama)} - ${sanitize(row.program)} - Session ${sanitize(row.session)} - Lunas`;
}

// ✅ Validasi Data Peserta SUPER KETAT
function validateFinanceRowsStrict(rows) {
  const resiFolder = DriveApp.getFolderById(FINANCE_FOLDER_RESI_ID);
  const buktiFolder = DriveApp.getFolderById(FINANCE_FOLDER_TRANSFER_ID);
  const invalidRows = [];

  rows.forEach(row => {
    const errors = [];

    // 🔎 Validasi 1: Status Bukti Transfer
    if (row.statusBuktiTransfer !== "✅ Berhasil di-upload" && !errors.includes("Bukti Transfer")) {
      errors.push("Bukti Transfer");
    }

    // 🔎 Validasi 2: File Resi harus ada
    const resiName = generateFileNameResi(row);
    const resiExists = resiFolder.getFilesByName(resiName).hasNext();
    if (!resiExists && !errors.includes("Resi")) {
      errors.push("Resi");
    }

    // 🔎 Validasi 3: File Bukti Transfer harus ada
    const buktiName = generateFileNameBukti(row);
    const buktiExists = buktiFolder.getFilesByName(buktiName).hasNext();
    if (!buktiExists && !errors.includes("Bukti Transfer")) {
      errors.push("Bukti Transfer");
    }

    // ⛔️ Kalau ada error, catat peserta
    if (errors.length > 0) {
      invalidRows.push({
        nama: row.nama,
        errors: errors
      });
    }
  });

  return invalidRows;
}

// ✅ Tampilkan Alert jika ada peserta invalid
function showFinanceValidationErrors(invalidRows) {
  if (invalidRows.length === 0) return;

  const list = invalidRows.map(r => `${r.nama} (${r.errors.join(" & ")})`).join("\n");
  SpreadsheetApp.getUi().alert(`❌ Proses dibatalkan.\n\nPeserta berikut belum memenuhi syarat:\n\n${list}`);
}

// ✅ Lookup Nama Program dari Form Responses
function lookupProgramFromFormResponses(idTrx) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxIdTrx = header.findIndex(h => String(h).toLowerCase().includes("id transaksi"));
  const idxProgram = header.findIndex(h => String(h).toLowerCase().includes("program"));

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxIdTrx] || "").trim() === idTrx) {
      return normalizeWhitespace(data[i][idxProgram] || "");
    }
  }
  return "";
}
