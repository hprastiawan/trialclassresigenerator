// ====================================================================================
// 📄 Resi Generator PDF - Resi Quick Class & Email Automation System
// ====================================================================================


// ------------------------------------------------------------------------
// ✅ Konfigurasi Setup Awal: Slide ID, Folder Output, Sheet, dan Header Index
// ------------------------------------------------------------------------
function getResiSetup() {
  const slideTemplateId = '1C8rohAmNeqb6VyL5bErBKPYSy69yN2kB8v5Yui2nAcY'; // Template Google Slides
  const folderOutputId = '1LwOj_U3zqf8YYFUIbEXvScQARQR9nq8R'; // Folder "Resi FC Berbayar Hasil Generate"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getDisplayValues();
  const headersRow = data[0];
  const getCol = name => headersRow.indexOf(name);

  const headers = {
    trxCol: getCol("ID Transaksi"),
    regCol: getCol("ID Registrasi"),
    emailCol: getCol("Email"),
    nameCol: getCol("Nama Lengkap"),
    phoneCol: getCol("Nomor Telepon Hashing Otomatis"),
    tglTextCol: getCol("Tanggal dan Jam Transaksi"),
    channelCol: getCol("Channel Pembayaran"),
    programCol: getCol("Nama Program"),
    periodeCol: getCol("Periode Pelaksanaan"),
    sessionCol: getCol("Session"),
    jmlBayarCol: getCol("Jumlah Pembayaran"),
    metodeCol: getCol("Metode Pembayaran"),
    statusResiCol: getCol("Status Resi PDF"),
    statusFileCol: getCol("File dalam Folder"),
    sendEmailStatusCol: getCol("Send Email Status"),
    jmlTopikCol: getCol("Jumlah Topik yang Diikuti"),
    topikStartCol: getCol("Topik 1") // Kolom P (index ke-16)
  };

  return {
    sheet,
    headers,
    slideTemplateId,
    folderOutputId
  };
}


// ------------------------------------------------------------------------
// ✅ Generate Nama File Resi berdasarkan format standar
// ------------------------------------------------------------------------
function generateResiFileName(row, h) {
  const sanitize = str => String(str).replace(/[\\/:*?"<>|]/g, "").trim();
  return `[Receipt TC Phincon Academy] ${sanitize(row[h.trxCol])} - ${sanitize(row[h.nameCol])} - ${sanitize(row[h.programCol])} - Session ${sanitize(row[h.sessionCol])} - Lunas`;
}


// ------------------------------------------------------------------------
// ✅ Fungsi Utama: Generate Resi untuk Baris Aktif
// ------------------------------------------------------------------------
function generateResiPDFforCurrentRow() {
  const { sheet, headers, folderOutputId, slideTemplateId } = getResiSetup();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();

  // ⛔️ Validasi jika memilih header
  if (row === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ✅ Ambil data baris aktif
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ⛔️ Validasi kelengkapan ID Transaksi & Registrasi
  if (!data[headers.trxCol] || !data[headers.regCol]) {
    return ui.alert("🚫 Data belum lengkap untuk baris ini");
  }

  // ✅ Siapkan variabel
  headers.rowIndex = row - 2;
  const fileName = generateResiFileName(data, headers);
  const outputFolder = DriveApp.getFolderById(folderOutputId);
  const existing = outputFolder.getFilesByName(fileName);

  // ⚠️ Jika file sudah ada → konfirmasi overwrite
  if (existing.hasNext()) {
    const confirm = ui.alert(`❗File \"${fileName}\" sudah ada. Mau ganti?`, ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.NO) return;
    existing.next().setTrashed(true);
  }

  // ✅ Buat file PDF
  createResiPDF(data, headers, slideTemplateId, outputFolder, fileName, sheet);

  // ✅ Feedback ke user
  SpreadsheetApp.getActiveSpreadsheet().toast(`1 file berhasil digenerate`, "Progress", 3);
  ui.alert(`✅ File berhasil digenerate untuk: ${data[headers.nameCol]}`);
}

function generateResiPDFFromSelection() {
  const { sheet, headers, slideTemplateId, folderOutputId } = getResiSetup();
  const ui = SpreadsheetApp.getUi();
  const outputFolder = DriveApp.getFolderById(folderOutputId);

  const selection = sheet.getActiveRangeList();
  if (!selection) return;

  const selectedRows = new Set();
  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i >= 2) selectedRows.add(i);
    }
  });

  const rowIndexes = [...selectedRows].sort((a, b) => a - b);
  let count = 0;

  rowIndexes.forEach(row => {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!rowData[headers.trxCol] || !rowData[headers.regCol]) return;

    headers.rowIndex = row - 2;
    const fileName = generateResiFileName(rowData, headers);
    const existing = outputFolder.getFilesByName(fileName);

    // ✅ Update kolom "File dalam Folder"
    let fileStatus = "-";
    if (existing.hasNext()) {
      fileStatus = "Ada";
    } else {
      fileStatus = rowData[headers.statusResiCol] === "✅ PDF Generated" ? "Pernah dihapus" : "Belum Ada";
    }
    sheet.getRange(row, headers.statusFileCol + 1).setValue(fileStatus);

    // ✅ Hanya generate ulang jika "Pernah dihapus" atau "Belum Ada"
    if (fileStatus === "Pernah dihapus" || fileStatus === "Belum Ada") {
      createResiPDF(rowData, headers, slideTemplateId, outputFolder, fileName, sheet);
      count++;
      SpreadsheetApp.getActiveSpreadsheet().toast(`${count} file berhasil digenerate...`, "Progress", 3);
    }
  });

  ui.alert(`✅ ${count} file resi berhasil digenerate dari baris terpilih`);
}


// ------------------------------------------------------------------------
// ✅ Fungsi Inti: Membuat file PDF berdasarkan baris data
// ------------------------------------------------------------------------
function createResiPDF(row, h, slideTemplateId, outputFolder, fileName, sheet) {
  const slideCopy = DriveApp.getFileById(slideTemplateId).makeCopy(`Resi - ${row[h.nameCol]}`);
  const presentation = SlidesApp.openById(slideCopy.getId());
  const slide = presentation.getSlides()[0];

  // ✅ Ambil data list topik dari kolom P–AD
  const topik = row.slice(h.topikStartCol, h.topikStartCol + 15);
  const listTopikFormatted = formatListTopikToMulticolumn(topik);

  // ✅ Format angka manual ke bentuk 9.500.000,-
  const formatManual = val => {
    if (!val || val === "-" || val === "0" || val === "Rp 0") return "-";
    const cleaned = val.toString().replace(/[^0-9]/g, '');
    const formatted = cleaned.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return `${formatted},-`;
  };

  // ✅ Mapping tag placeholder di template dengan data sheet
  const replacements = {
    '<<tanggaltrx>>': row[h.tglTextCol],
    '<<idtrx>>': row[h.trxCol],
    '<<channel>>': row[h.channelCol],
    '<<idreg>>': row[h.regCol],
    '<<namapeserta>>': row[h.nameCol],
    '<<email>>': row[h.emailCol],
    '<<notlp>>': row[h.phoneCol],
    '<<namaprog>>': row[h.programCol],
    '<<periode>>': formatPeriode(row[h.periodeCol]),
    '<<session>>': row[h.sessionCol],
    '<<jmlbayar>>': formatManual(row[h.jmlBayarCol]),
    '<<metodebayar>>': row[h.metodeCol],
    '<<hargaprog>>': formatManual(row[h.jmlBayarCol]), // Harga = jml bayar
    '<<jmltopik>>': row[h.jmlTopikCol],
    '<<listtopik>>': listTopikFormatted
  };


// ------------------------------------------------------------------------
// ✅ Fungsi Utama: Cetak Resi untuk Seluruh Data
// ------------------------------------------------------------------------
function generateResiPDFForAll() {
  const { sheet, headers, slideTemplateId, folderOutputId } = getResiSetup();
  const ui = SpreadsheetApp.getUi();
  const data = sheet.getDataRange().getDisplayValues();
  const outputFolder = DriveApp.getFolderById(folderOutputId);

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 1;

    // ⛔️ Lewati jika data belum lengkap
    if (!row[headers.trxCol] || !row[headers.regCol]) continue;

    const fileName = generateResiFileName(row, headers);
    const found = outputFolder.getFilesByName(fileName);

    // ✅ Tentukan status berdasarkan kondisi aktual file di Google Drive
    let fileStatus = "-";
    const statusResi = row[headers.statusResiCol];
    const statusFileNow = row[headers.statusFileCol];

    if (statusResi === "✅ PDF Generated" && found.hasNext()) {
      fileStatus = "Ada";
      sheet.getRange(rowIndex, headers.statusFileCol + 1).setValue(fileStatus);
      continue; // ✅ Lewati jika file sudah ada
    } else if (statusResi === "✅ PDF Generated" && !found.hasNext()) {
      fileStatus = "Pernah dihapus";
    } else {
      fileStatus = "Belum dibuat";
    }

    // 📝 Update kolom "File dalam Folder"
    sheet.getRange(rowIndex, headers.statusFileCol + 1).setValue(fileStatus);

    if (fileStatus === "Pernah dihapus" || fileStatus === "Belum dibuat") {
      headers.rowIndex = i;
      createResiPDF(row, headers, slideTemplateId, outputFolder, fileName, sheet);
      count++;
      SpreadsheetApp.getActiveSpreadsheet().toast(`${count} file berhasil digenerate...`, "Progress", 3);
    }
  }

  ui.alert(`✅ Proses selesai.\n${count} file resi berhasil digenerate dari seluruh data.`);
}


  // ✅ Replace seluruh tag di slide
  for (const [tag, value] of Object.entries(replacements)) {
    slide.replaceAllText(tag, value);
  }

  // ✅ Simpan & export ke PDF
  presentation.saveAndClose();
  const blob = DriveApp.getFileById(slideCopy.getId()).getAs(MimeType.PDF);
  blob.setName(fileName);
  outputFolder.createFile(blob);
  DriveApp.getFileById(slideCopy.getId()).setTrashed(true); // Hapus file .gslides setelah PDF dibuat

  // ✅ Update status di sheet
  if (typeof h.rowIndex !== 'undefined') {
    const rowInSheet = h.rowIndex + 2;
    sheet.getRange(rowInSheet, h.statusResiCol + 1).setValue("✅ PDF Generated");
    sheet.getRange(rowInSheet, h.statusFileCol + 1).setValue("Ada");
  }
}

// ✅ Format tanggal periode ke bentuk "April 2025"
function formatPeriode(val) {
  if (!(val instanceof Date)) return String(val); // fallback jika bukan Date
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return `${months[val.getMonth()]} ${val.getFullYear()}`;
}

// ------------------------------------------------------------------------
// ✅ Validasi Ulang: Cek file resi berdasarkan status di sheet
// ------------------------------------------------------------------------
function validateResiFileExistence() {
  const { sheet, headers, folderOutputId } = getResiSetup();
  const data = sheet.getDataRange().getDisplayValues();
  const outputFolder = DriveApp.getFolderById(folderOutputId);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fileStatus = row[headers.statusFileCol];
    const idTrx = row[headers.trxCol];
    const reg = row[headers.regCol];

    // Lewati jika data belum lengkap
    if (!idTrx || !reg) continue;

    const fileName = generateResiFileName(row, headers);
    const found = outputFolder.getFilesByName(fileName);

    if (fileStatus === "Ada" && !found.hasNext()) {
      sheet.getRange(i + 1, headers.statusFileCol + 1).setValue("Pernah dihapus");
    }

    if (!fileStatus || fileStatus.trim() === "") {
      sheet.getRange(i + 1, headers.statusFileCol + 1).setValue("Belum dibuat");
    }
  }
}

// ------------------------------------------------------------------------
// ✅ Fungsi Utama: Generate Resi berdasarkan filter Periode tertentu
// ------------------------------------------------------------------------
// ✅ Fungsi Utama: Generate Resi berdasarkan filter Periode tertentu
function generateResiPDFFilteredByPeriode(periodeTarget) {
  const { sheet, headers, slideTemplateId, folderOutputId } = getResiSetup();
  const data = sheet.getDataRange().getDisplayValues();
  const outputFolder = DriveApp.getFolderById(folderOutputId);
  const ui = SpreadsheetApp.getUi();

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowInSheet = i + 1;

    // ⛔️ Lewati jika periode tidak sesuai
    if (row[headers.periodeCol] !== periodeTarget) continue;

    // ⛔️ Lewati jika ID tidak lengkap
    if (!row[headers.trxCol] || !row[headers.regCol]) continue;

    const fileName = generateResiFileName(row, headers);
    const existingFile = outputFolder.getFilesByName(fileName);
    const statusResi = row[headers.statusResiCol];

    // ✅ Tentukan status file
    let fileStatus = "-";
    if (statusResi === "✅ PDF Generated" && existingFile.hasNext()) {
      fileStatus = "Ada";
    } else if (statusResi === "✅ PDF Generated") {
      fileStatus = "Pernah dihapus";
    } else {
      fileStatus = "Belum dibuat";
    }

    // 📝 Update kolom AG (File dalam Folder)
    sheet.getRange(rowInSheet, headers.statusFileCol + 1).setValue(fileStatus);

    // ✅ Generate ulang jika file belum ada atau pernah dihapus
    if (fileStatus === "Belum dibuat" || fileStatus === "Pernah dihapus") {
      headers.rowIndex = i;
      createResiPDF(row, headers, slideTemplateId, outputFolder, fileName, sheet);

      // 🟢 Tambahkan update kolom AF setelah PDF berhasil dibuat
      sheet.getRange(rowInSheet, headers.statusResiCol + 1).setValue("✅ PDF Generated");
      sheet.getRange(rowInSheet, headers.statusFileCol + 1).setValue("Ada");

      count++;
      SpreadsheetApp.getActiveSpreadsheet().toast(`${count} file berhasil digenerate...`, "Progress", 3);
    }
  }

  ui.alert(`✅ Selesai! ${count} file resi berhasil digenerate untuk periode: ${periodeTarget}`);
}
