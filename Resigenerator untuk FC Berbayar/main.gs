// ====================================================================================
// 📁 Resi Quick Class & Email Automation System - Phincon Academy
// ====================================================================================

// ✅ Global Constants - Folder IDs
const FOLDER_OUTPUT_QUICKCLASS = "1WGGxu1ZiECUUiI_Cr-97sBpj1os-cfRy"; // Folder utama output (Generated) Resi dan Bukti Transfer FC Berbayar
const FOLDER_RESI_QUICKCLASS = "1LwOj_U3zqf8YYFUIbEXvScQARQR9nq8R"; // Folder Resi FC Berbayar Hasil Generate
const FOLDER_TRANSFER_QUICKCLASS = "10mDu1jxvA4CxO-vuFkfLYq0GdCUjZEEY"; // Folder Bukti Transfer FC Berbayar Hasil Generate
const FOLDER_MERGED_RESI_TRANSFER = "1axbicwtUkVLQm-vQvguA94XRnAdbdTpT"; // Folder gabungan Resi & Bukti Transfer FC Berbayar

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses 1");

  SpreadsheetApp.flush();
  Utilities.sleep(300);

  // ✅ Tambah menu dulu agar selalu muncul
  createResiAutomationMenus();

  // ✅ Proses lainnya
  validateAndHighlightHeaders();
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment("center");
  autoResizeAllColumnsSmart(sheet);
  updateAutoNumbering(sheet);
  updateNameEmailAndPaymentFormatting(sheet);

  const namaProgramList = getListFromSheet("listNamaProgram");
  applyDropdownToColumn(sheet, 7, namaProgramList);

  const periodeList = getListFromSheet("listPeriode");
  applyDropdownToColumn(sheet, 9, periodeList);

  const channelList = getListFromSheet("channelPembayaran");
  applyDropdownToColumn(sheet, 15, channelList);

  const topicList = getListFromSheet("listTopic");
  for (let col = 16; col <= 30; col++) {
    applyDropdownToColumn(sheet, col, topicList);
  }

  applyDatePickerToTanggalTransaksi(sheet);        // ✅ Validasi hanya untuk Kolom L
  updateFormattedTanggalTransaksi(sheet);
  syncTanggalJamTransaksiFromFormatted(sheet);
  highlightDuplicateTopik();
  validateAllTopicCounts();
  updatePhoneHashing(sheet);
  alignCenterSpecificColumns(sheet);
  cleanEmailNamePhone(sheet);
  validasiIDTrxDanReg(sheet);
  
  showReadyConfirmation();
}


function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Form Responses 1") return;

  // ✅ Re-check header jika baris ke-1 disentuh
  if (e.range.getRow() === 1) {
    validateAndHighlightHeaders();
    return;
  }

  const editedCol = e.range.getColumn();
  const row = e.range.getRow();
  const topicStartCol = 16; // Kolom P
  const topicEndCol = 30;   // Kolom AD
  const jumlahTopikCol = 10; // Kolom J

  // ✅ Deteksi kolom "Keterangan Error" secara dinamis
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.findIndex(h => String(h).toLowerCase().includes("keterangan error")) + 1;

  if (statusCol === 0) return;

  // Jika kolom Topik yang diedit
  if (editedCol >= topicStartCol && editedCol <= topicEndCol) {
    highlightDuplicateTopik();
    validateTopicCount(sheet, row, topicStartCol, topicEndCol, statusCol);
  }

  // Jika kolom Jumlah Topik yang Diikuti (Kolom J) diedit
  if (editedCol === jumlahTopikCol) {
    validateTopicCount(sheet, row, topicStartCol, topicEndCol, statusCol);
  }

  // ✅ Update Tanggal dan Jam Transaksi jika Kolom L (Tanggal Transaksi) diubah
  if (editedCol === 12) {
    applyDatePickerToTanggalTransaksi(sheet);
    updateFormattedTanggalTransaksi(sheet);

    const tanggal = sheet.getRange(row, 12).getValue(); // kolom L
    const existing = sheet.getRange(row, 13).getValue(); // kolom M

    if (tanggal instanceof Date) {
      let jam = "00:00";

      // Cek apakah existing (kolom M) sudah ada jam
      if (typeof existing === "string") {
        const match = existing.match(/\d{1,2}:\d{2}$/);
        if (match) jam = match[0]; // pakai jam yang lama
      }

      const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
      const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                      'Agustus', 'September', 'Oktober', 'November', 'Desember'];

      const hari = days[tanggal.getDay()];
      const tgl = tanggal.getDate();
      const bulan = months[tanggal.getMonth()];
      const tahun = tanggal.getFullYear();

      const finalFormatted = `${hari}, ${tgl} ${bulan} ${tahun} ${jam}`;
      sheet.getRange(row, 13).setValue(finalFormatted); // Kolom M
    }
  }

  // // ✅ Update Tanggal dan Jam Transaksi jika Kolom L (Tanggal Transaksi) diubah
  // if (editedCol === 12) {
  //   applyDatePickerToTanggalTransaksi(sheet);
  //   updateFormattedTanggalTransaksi(sheet);
  //   syncTanggalJamTransaksiFromFormatted(sheet);
  // }

  // ✅ Jika kolom Nomor Telepon (F) diedit, update hashing ke kolom AE
  if (editedCol === 6) {
    const phone = sheet.getRange(row, 6).getValue();
    const hash = hashPhoneNumber(phone);
    sheet.getRange(row, 31).setValue(hash); // Kolom AE
  }

  const centerCols = [8, 10, 14, 33]; // H, J, N, AG
  if (centerCols.includes(editedCol)) {
    sheet.getRange(row, editedCol).setHorizontalAlignment("center");
  }

  if ([4, 5, 6].includes(editedCol)) {
  const original = sheet.getRange(row, editedCol).getValue();
  const cleaned = sanitizeInput(original);
    if (original !== cleaned) {
      sheet.getRange(row, editedCol).setValue(cleaned);
    }
  }

  const triggerCols = [8, 9, 12]; // Kolom H, I, L
    if (triggerCols.includes(e.range.getColumn()) && e.range.getRow() > 1) {
    validasiIDTrxDanReg(sheet, e.range.getRow());
  }


  updateAutoNumbering(sheet);
  updateNameEmailAndPaymentFormatting(sheet);
}


function refreshResiStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

  SpreadsheetApp.flush();
  Utilities.sleep(300);

  try {
    validateAndHighlightHeaders();
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment("center");
    autoResizeAllColumnsSmart(sheet);
    updateAutoNumbering(sheet);
    updateNameEmailAndPaymentFormatting(sheet);

    const namaProgramList = getListFromSheet("listNamaProgram");
    applyDropdownToColumn(sheet, 7, namaProgramList); // Kolom G

    const periodeList = getListFromSheet("listPeriode");
    applyDropdownToColumn(sheet, 9, periodeList); // Kolom I

    const channelList = getListFromSheet("channelPembayaran");
    applyDropdownToColumn(sheet, 15, channelList); // Kolom O

    const topicList = getListFromSheet("listTopic");
    for (let col = 16; col <= 30; col++) {
      applyDropdownToColumn(sheet, col, topicList);
    }

    applyDatePickerToTanggalTransaksi(sheet);        // ✅ Validasi hanya untuk Kolom L
    updateFormattedTanggalTransaksi(sheet);
    syncTanggalJamTransaksiFromFormatted(sheet);
    highlightDuplicateTopik();
    validateAllTopicCounts();
    updatePhoneHashing(sheet);
    alignCenterSpecificColumns(sheet);
    cleanEmailNamePhone(sheet);

    // 🔁 Buat folder berdasarkan Periode Pelaksanaan (hanya untuk data yang sudah valid)
    createFoldersPerPeriode(true);
    validateResiFileExistence(); // ✅ Tambahan: Cek ulang file yang "Ada", tapi hilang
    validasiIDTrxDanReg(sheet);

    SpreadsheetApp.getUi().alert("✅ Halaman berhasil di-refresh");
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Gagal refresh halaman: " + err.message);
    console.error("refreshResiStatus error:", err);
  }
}
