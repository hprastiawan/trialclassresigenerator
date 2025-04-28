// ====================================================================================
// 📄 Formatter Khusus Sheet "Kirim ke Tim Finance"
// ====================================================================================


// ✅ Auto resize semua kolom berdasarkan isi + styling header
function financeUpdateNameEmailAndPaymentFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emailRange = sheet.getRange(2, 4, lastRow - 1); // Kolom D
  const nameRange = sheet.getRange(2, 5, lastRow - 1);  // Kolom E

  const emails = emailRange.getValues();
  const names = nameRange.getValues();

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i][0]?.toString().trim().toLowerCase() || "";
    emailRange.getCell(i + 1, 1).setValue(email);
    const isValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    emailRange.getCell(i + 1, 1).setBackground(isValid ? null : "#f4cccc");

    const name = names[i][0]?.toString().trim() || "";
    nameRange.getCell(i + 1, 1).setValue(toTitleCase(name));
  }
}

// ✅ Center align untuk kolom Session dan Jumlah Topik
function financeAlignCenterColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const columnsToCenter = [6, 8]; // Kolom Session (F), Jumlah Topik (H)
  columnsToCenter.forEach(col => {
    sheet.getRange(2, col, lastRow - 1).setHorizontalAlignment("center");
  });
}


// ✅ Format kapitalisasi nama & validasi email untuk sheet Kirim ke Tim Finance
function financeUpdateNameEmailAndPaymentFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emailRange = sheet.getRange(2, 4, lastRow - 1); // Kolom D = Email
  const nameRange = sheet.getRange(2, 5, lastRow - 1);  // Kolom E = Nama Lengkap

  const emails = emailRange.getValues();
  const names = nameRange.getValues();

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i][0]?.toString().trim().toLowerCase() || "";
    emailRange.getCell(i + 1, 1).setValue(email);
    const isValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    emailRange.getCell(i + 1, 1).setBackground(isValid ? null : "#f4cccc");

    const name = names[i][0]?.toString().trim() || "";
    nameRange.getCell(i + 1, 1).setValue(toTitleCase(name));
  }
}



// ✅ Utility: Format kapitalisasi tiap kata (Title Case)
function toTitleCase(str) {
  return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}


// ====================================================================================
// 📂 Helper: Buat Folder Tanggal Jika Belum Ada di SHEET: KIRIM KE TIM FINANCE
// ====================================================================================

// Fungsi untuk membuat folder tanggal (misal: "7 April 2025") dalam folder periode
function createFolderTanggalIfNeeded(parentFolderId, tanggalString) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const dateObj = new Date(tanggalString);
    if (isNaN(dateObj)) return "";

    const formatter = Utilities.formatDate;
    const tanggalFormatted = formatter(dateObj, Session.getScriptTimeZone(), "d MMMM yyyy");
    const folders = parentFolder.getFoldersByName(tanggalFormatted);

    if (folders.hasNext()) {
      return folders.next().getId(); // Folder sudah ada
    } else {
      const newFolder = parentFolder.createFolder(tanggalFormatted);
      return newFolder.getId(); // Folder baru dibuat
    }
  } catch (e) {
    Logger.log("❌ Error createFolderTanggalIfNeeded: " + e);
    return "";
  }
}

// ✅ Auto resize semua kolom berdasarkan isi + styling header (khusus sheet Kirim ke Tim Finance)
function financeAutoResizeAllColumnsSmart(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let col = 1; col <= lastCol; col++) {
    const dataRange = sheet.getRange(2, col, Math.max(0, lastRow - 1));
    const colValues = dataRange.getValues().flat().filter(v => v !== "" && v !== null);
    const headerText = headers[col - 1] || "";

    if (colValues.length > 0) {
      sheet.autoResizeColumn(col);
      const currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, currentWidth + 20);
    } else {
      const estimatedWidth = Math.max(120, headerText.length * 9);
      sheet.setColumnWidth(col, estimatedWidth);
    }
  }

  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  headerRange.setWrap(true);
  headerRange.setBackground("#d9ead3");

  sheet.setFrozenRows(1);
  // ✅ Atur align center khusus untuk kolom L (Status Pembayaran)
  if (lastRow > 1) {
    const statusPembayaranRange = sheet.getRange(2, 12, lastRow - 1, 1);
    statusPembayaranRange.setHorizontalAlignment("center");
    statusPembayaranRange.setVerticalAlignment("middle");
  }
}
