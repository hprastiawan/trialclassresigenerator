// ====================================================================================
// 📁 Helper: Formatter dan Utilitas Data - Resi Automation System
// ====================================================================================


// ------------------------------------------------------------------------
// ✨ AUTO FORMAT & STYLING
// ------------------------------------------------------------------------

// ✅ Auto resize semua kolom berdasarkan isi + styling header
function autoResizeAllColumnsSmart(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let col = 1; col <= lastCol; col++) {
    const dataRange = sheet.getRange(2, col, lastRow - 1);
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

  // ✅ Atur lebar kolom AJ, AK, AL jika sheet memiliki cukup kolom
  const safeCols = sheet.getLastColumn();
  if (safeCols >= 38) {
    sheet.setColumnWidth(36, 280); // Kolom AJ
    sheet.setColumnWidth(37, 200); // Kolom AK
    sheet.setColumnWidth(38, 240); // Kolom AL
  }

  // ✅ Atur lebar kolom AI agar pesan error terbaca jelas
  if (safeCols >= 35) {
    sheet.setColumnWidth(35, 380); // Kolom AI = Keterangan Error
  }


  // sheet.setColumnWidth(36, 280); // Kolom AJ - Folder ID Periode
  // sheet.setColumnWidth(37, 200); // Kolom AK - Status File di Folder
  // sheet.setColumnWidth(38, 240); // Kolom AL - Status Folder Periode

  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  headerRange.setWrap(true);
  headerRange.setBackground("#d9ead3");
}


// ✅ Center align untuk kolom H, J, N, dan AG
function alignCenterSpecificColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const columnsToCenter = [8, 10, 14, 33]; // Kolom H, J, N, AG
  columnsToCenter.forEach(col => {
    sheet.getRange(2, col, lastRow - 1).setHorizontalAlignment("center");
  });
}


// ------------------------------------------------------------------------
// 🔢 AUTO NUMBER & HASH
// ------------------------------------------------------------------------

// ✅ Auto numbering di kolom A jika Nama & Email terisi
function updateAutoNumbering(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const noCol = 1;
  const emailCol = 4;
  const namaCol = 5;

  for (let i = 1; i < values.length; i++) {
    const email = (values[i][emailCol - 1] || "").toString().trim();
    const nama = (values[i][namaCol - 1] || "").toString().trim();

    if (email && nama) {
      sheet.getRange(i + 1, noCol).setValue(i);
    } else {
      sheet.getRange(i + 1, noCol).clearContent();
    }
  }
}

// ✅ Fungsi hash nomor telepon (output: **********123)
function hashPhoneNumber(phone) {
  if (!phone) return '';
  const digits = phone.toString().replace(/\D/g, '');
  const last3 = digits.slice(-3);
  return '*'.repeat(13 - last3.length) + last3;
}

// ✅ Update kolom AE berdasarkan kolom F (Nomor Telepon)
function updatePhoneHashing(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const phoneRange = sheet.getRange(2, 6, lastRow - 1); // Kolom F
  const hashRange = sheet.getRange(2, 31, lastRow - 1); // Kolom AE
  const phoneValues = phoneRange.getValues();
  const hashValues = hashRange.getValues();

  for (let i = 0; i < phoneValues.length; i++) {
    const phone = phoneValues[i][0];
    const newHash = hashPhoneNumber(phone);
    if (newHash !== hashValues[i][0]) {
      hashRange.getCell(i + 1, 1).setValue(newHash);
    }
  }
}


// ------------------------------------------------------------------------
// 🧹 FORMAT TEXT: EMAIL, NAMA, METODE
// ------------------------------------------------------------------------

// ✅ Format kapitalisasi tiap kata (Title Case)
function toTitleCase(str) {
  return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

// ✅ Format dan validasi Email, Nama Lengkap, Metode Pembayaran
function updateNameEmailAndPaymentFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emailRange = sheet.getRange(2, 4, lastRow - 1);   // Kolom D
  const nameRange = sheet.getRange(2, 5, lastRow - 1);    // Kolom E
  const methodRange = sheet.getRange(2, 14, lastRow - 1); // Kolom N

  const emails = emailRange.getValues();
  const names = nameRange.getValues();
  const methods = methodRange.getValues();

  for (let i = 0; i < emails.length; i++) {
    const rawEmail = emails[i][0] ? emails[i][0].toString().trim() : "";
    const lowerEmail = rawEmail.toLowerCase();
    emailRange.getCell(i + 1, 1).setValue(lowerEmail);
    const isValidEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(lowerEmail);
    emailRange.getCell(i + 1, 1).setBackground(isValidEmail ? null : "#f4cccc");

    const rawName = names[i][0] ? names[i][0].toString().trim() : "";
    nameRange.getCell(i + 1, 1).setValue(toTitleCase(rawName));

    const rawMethod = methods[i][0] ? methods[i][0].toString().trim() : "";
    methodRange.getCell(i + 1, 1).setValue(toTitleCase(rawMethod));
  }
}

// ------------------------------------------------------------------------
// 🧹 MENGATUR FORMAT LIST TOPIK DI FILE RESI
// ------------------------------------------------------------------------

// ✅ Fungsi untuk format topik ke bentuk multi-kolom (auto-adaptif)
function formatListTopikToMulticolumn(topikArray) {
  const cleaned = topikArray.filter(t => t && String(t).trim() !== "");
  const total = cleaned.length;

  let colCount = 1;
  if (total <= 5) colCount = 1;
  else if (total <= 10) colCount = 2;
  else colCount = 3;

  const rows = Math.ceil(total / colCount);
  const columns = [];

  for (let i = 0; i < colCount; i++) {
    columns.push(cleaned.slice(i * rows, (i + 1) * rows));
  }

  const resultLines = [];
  for (let i = 0; i < rows; i++) {
    const lineParts = [];
    for (let j = 0; j < colCount; j++) {
      const item = columns[j][i];
      const index = i + 1 + j * rows;
      if (item) {
        lineParts.push(`${index}. ${item}`.padEnd(35)); // Sesuaikan padding jika perlu
      }
    }
    resultLines.push(lineParts.join(" ")); // Em space untuk jarak kolom
  }

  return resultLines.join("\n");
}
