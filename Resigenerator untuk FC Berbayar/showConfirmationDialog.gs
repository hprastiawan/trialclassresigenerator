// ====================================================================================
// 📁 Dialog & Konfirmasi - Resi Quick Class & Email Automation System
// ====================================================================================

// ✅ Show confirmation alert setelah semua proses onOpen selesai
function showReadyConfirmation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("✅ Dokumen ini sudah siap digunakan", "Phincon Academy", 5);
  Logger.log("✅ showReadyConfirmation executed");
}

// ✅ Konfirmasi Menu Refresh untuk Resi Quick Class
function showRefreshResiConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🔁 Refresh Halaman",
    "Apakah Kamu yakin ingin me-refresh halaman ini?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    refreshResiStatus(); // ✅ Fungsi utama untuk refresh semua tampilan
  }
}

// ✅ Konfirmasi Generate ID Transaksi & ID Registrasi
function showGenerateIdTrxReg() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ Konfirmasi",
    "Apakah Kamu yakin ingin generate ID Transaksi & Registrasi?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    generateIdTrxAndReg(true); // ✅ Jalankan dengan alert ringkasan
  }
}

// ------------------------------------------------------------------------
// ⚠️ CONFIRMATION LIST / MANAJEMEN DIREKTORI
// ------------------------------------------------------------------------


// ✅ Konfirmasi: Buat Folder per Periode
function showCreateFolderPerPeriode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "📁 Buat Folder per Periode",
    "Apakah kamu ingin membuat folder berdasarkan Periode Pelaksanaan?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    createFoldersPerPeriode(true);
  }
}

// ✅ Konfirmasi: Bersihkan Folder Kosong
function showEraseEmptyFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🗑️ Bersihkan Folder Kosong",
    "Apakah kamu yakin ingin menghapus folder yang kosong dari folder 'Resi & Bukti Transfer FC Berbayar'?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    eraseEmptyFoldersFromMergedFolder();
  }
}

// ✅ Konfirmasi: Generate Summary Peserta
function showGenerateSummaryPeserta() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "📊 Generate Summary Peserta",
    "Apakah kamu ingin membuat atau memperbarui sheet ringkasan peserta berdasarkan Periode Pelaksanaan?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    generateSummaryPesertaSheet();
  }
}

// ------------------------------------------------------------------------
// 📕 KONFIRMASI: Generate Resi untuk Baris Aktif
// ------------------------------------------------------------------------

// ✅ Fungsi dialog konfirmasi sebelum generate resi 1 baris
function showResiBarisAktif() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const activeRow = sheet.getActiveCell().getRow();
  const ui = SpreadsheetApp.getUi();

  // ⛔️ Validasi: Tidak boleh memilih baris header
  if (activeRow === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ❗️Validasi: Tidak memilih baris yang valid
  if (activeRow < 2) {
    ui.alert("‼️ Pilih salah satu baris data terlebih dahulu");
    return;
  }

  const nama = sheet.getRange(activeRow, 5).getValue(); // Kolom E = Nama Lengkap
  const response = ui.alert("⚠️ Konfirmasi", `Apakah Kamu yakin ingin membuat resi untuk ${nama}?`, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateResiPDFforCurrentRow(); // Fungsi utama ada di resiGeneratorPDF.gs
  }
}

// ------------------------------------------------------------------------
// 📘📗 KONFIRMASI: Generate Resi untuk Beberapa Baris Terpilih
// ------------------------------------------------------------------------
function showResiBeberapaBarisAktif() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();
  if (!selection) return;

  const selectedRows = new Set();
  let headerSelected = false;

  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i === 1) headerSelected = true;
      if (i >= 2) selectedRows.add(i);
    }
  });

  if (headerSelected) {
    SpreadsheetApp.getUi().alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 2) {
    SpreadsheetApp.getUi().alert("‼️ Pilih minimal 2 baris data terlebih dahulu");
    return;
  }

  const namaList = rowIndexes.map(row => sheet.getRange(row, 5).getValue()).filter(n => n).join(", ");
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", `Apakah Kamu yakin ingin membuat resi untuk ${rowIndexes.length} peserta berikut?\n\n${namaList}`, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateResiPDFFromSelection();
  }
}

// ✅ Konfirmasi: Cetak Resi untuk Seluruh Data
function showResiSeluruhData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ Konfirmasi",
    "Apakah kamu yakin ingin membuat Resi untuk seluruh data yang ada?\n\nFile hanya akan digenerate ulang jika belum ada atau pernah dihapus.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    generateResiPDFForAll();
  }
}

function showGenerateResiByPeriode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "🗓️ Generate Resi Berdasarkan Periode",
    "Masukkan Periode Pelaksanaan (Contoh: April 2025):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const periode = response.getResponseText().trim();
    if (periode) generateResiPDFFilteredByPeriode(periode);
  }
}

// ------------------------------------------------------------------------
// 📧 KONFIRMASI: Kirim Resi untuk 1 Peserta (Baris Aktif)
// ------------------------------------------------------------------------
function showKirimEmailkePesertaBarisAktif() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const activeRow = sheet.getActiveCell().getRow();
  const ui = SpreadsheetApp.getUi();

  // ⛔️ Validasi baris header
  if (activeRow === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ⛔️ Validasi jika tidak memilih baris valid
  if (activeRow < 2) {
    ui.alert("⛔️ Pilih salah satu baris data terlebih dahulu");
    return;
  }

  const email = sheet.getRange(activeRow, 4).getValue(); // Kolom D = Email
  const nama = sheet.getRange(activeRow, 5).getValue();  // Kolom E = Nama Lengkap
  const trxId = sheet.getRange(activeRow, 2).getValue(); // Kolom B = ID Transaksi
  const regId = sheet.getRange(activeRow, 3).getValue(); // Kolom C = ID Registrasi

  // ⛔️ Validasi jika data wajib belum lengkap
  if (!email || !trxId || !regId) {
    ui.alert("❌ Data belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi");
    return;
  }

  const response = ui.alert(
    "⚠️ Konfirmasi Kirim Email",
    `Apakah Kamu yakin ingin mengirim resi ke peserta berikut?\n\n👤 Nama: ${nama}\n📧 Email: ${email}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    KirimEmailkePesertaBarisAktif(); // ✅ Fungsi utama untuk kirim email
  }
}

// ------------------------------------------------------------------------
// 📧 KONFIRMASI: Kirim Resi untuk Beberapa Peserta (Baris Terpilih)
// ------------------------------------------------------------------------
function showKirimEmailkeBeberapaPeserta() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();
  const ui = SpreadsheetApp.getUi();

  if (!selection) return;

  const selectedRows = new Set();
  let headerSelected = false;

  // ✅ Loop setiap range yang diseleksi dan kumpulkan nomor baris unik
  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i === 1) headerSelected = true;
      if (i >= 2) selectedRows.add(i);
    }
  });

  // ⛔️ Tidak boleh memilih header
  if (headerSelected) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 2) {
    ui.alert("‼️ Pilih minimal 2 baris data terlebih dahulu");
    return;
  }

  // ✅ Validasi data wajib per baris (email, ID Transaksi, ID Registrasi)
  const incomplete = rowIndexes.filter(row => {
    const email = sheet.getRange(row, 4).getValue(); // Kolom D = Email
    const trxId = sheet.getRange(row, 2).getValue(); // Kolom B = ID Transaksi
    const regId = sheet.getRange(row, 3).getValue(); // Kolom C = ID Registrasi
    return !(email && trxId && regId);
  });

  if (incomplete.length > 0) {
    ui.alert("❌ Beberapa baris belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi di semua baris yang dipilih");
    return;
  }

  // ✅ Ambil daftar nama peserta untuk preview konfirmasi
  const namaList = rowIndexes
    .map(row => sheet.getRange(row, 5).getValue()) // Kolom E = Nama Lengkap
    .filter(n => n)
    .join(", ");

  // 🔔 Konfirmasi akhir ke pengguna
  const response = ui.alert(
    "⚠️ Konfirmasi Kirim Email",
    `Apakah Kamu yakin ingin mengirim resi ke ${rowIndexes.length} peserta berikut?\n\n${namaList}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    KirimEmailkeBeberapaPeserta(); // ✅ Fungsi utama kirim email ke beberapa peserta
  }
}

// ------------------------------------------------------------------------
// 📬 KONFIRMASI: Kirim Resi untuk Seluruh Peserta
// ------------------------------------------------------------------------
function showKirimEmailkeSemuaPeserta() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();
  const ui = SpreadsheetApp.getUi();

  // ✅ Hitung jumlah peserta valid (email, ID Transaksi, ID Registrasi lengkap)
  const validRows = data.slice(1).filter(row =>
    row[1] && row[2] && row[3] // Kolom B (ID Transaksi), C (ID Registrasi), D (Email)
  );
  const totalValid = validRows.length;

  // ⛔️ Tidak ada data yang valid
  if (totalValid === 0) {
    ui.alert("❌ Tidak ada data peserta yang valid untuk dikirimi email.");
    return;
  }

  // ✅ Konfirmasi pengiriman
  const response = ui.alert(
    "⚠️ Konfirmasi Kirim Email Massal",
    `Apakah Kamu yakin ingin mengirim email ke seluruh peserta yang valid?\n\nJumlah peserta: ${totalValid}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    KirimEmailkeSemuaPeserta(); // ✅ Fungsi utama pengiriman email
  }
}

// ------------------------------------------------------------------------
// 📬 KONFIRMASI: Kirim Email berdasarkan Periode
// ------------------------------------------------------------------------
function showKirimEmailPesertaByPeriode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "📬 Kirim Email berdasarkan Periode",
    "Ketik periode yang ingin dikirim emailnya (misalnya: April 2025):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const periodeTarget = response.getResponseText().trim();
  if (!periodeTarget) {
    ui.alert("❗ Periode tidak boleh kosong.");
    return;
  }

  KirimEmailPesertaByPeriode(periodeTarget); // ✅ Fungsi utama di bawah
}

// ------------------------------------------------------------------------
// 📩 KONFIRMASI: Get Data Peserta untuk Tim Finance
// ------------------------------------------------------------------------
function showGetDataPeserta() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "⚠️ Konfirmasi Get Data Peserta",
    "Apakah Kamu yakin ingin memuat data peserta yang sudah valid dan mengirimnya ke sheet 'Kirim ke Tim Finance'?\n\nData hanya akan dimuat untuk peserta yang:\n• File resinya sudah tersedia (Kolom AG = 'Ada')\n• Email resi sudah berhasil dikirim (Kolom AH = 'Sending completed ✅')",
    ui.ButtonSet.OK_CANCEL
  );

  // ✅ Lanjutkan jika user klik OK
  if (response === ui.Button.OK) {
    loadDataKeFinance(); // Fungsi utama untuk generate isi sheet
  }
}

// ------------------------------------------------------------------------
// 📤 KONFIRMASI: Upload Bukti Transfer
// ------------------------------------------------------------------------
function showUploadBuktiTransferSidebar() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kirim ke Tim Finance");
  const range = sheet.getActiveRange();
  const rowIndex = range.getRow();

  if (rowIndex === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih.\nSilakan pilih baris data peserta.");
    return;
  }

  const namaLengkap = sheet.getRange(rowIndex, 5).getValue(); // Kolom E = Nama
  const response = ui.alert(
    "📤 Upload Bukti Transfer",
    `Apakah kamu yakin ingin mengupload bukti transfer untuk ${namaLengkap}?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    openUploadDialogForActiveRow(); // Fungsi ini ada di uploadBuktiTransferPeserta.gs
  }
}

// ====================================================================================
// 📁 showConfirmationDialog.gs
// ====================================================================================
// 📌 Konfirmasi sebelum 🚀 Kirim Email ke Tim Finance
// ====================================================================================

// ✅ Fungsi konfirmasi sebelum kirim email ke Tim Finance
function showKirimEmailKeTimFinance() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "🚀 Kirim Email ke Tim Finance",
    "Ketikkan Tanggal Transaksi yang ingin diproses. Format: dd/MM/yy (Contoh: 17/04/25)",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const tanggalStr = response.getResponseText().trim();
    if (!tanggalStr) {
      ui.alert("❌ Input tidak boleh kosong. Proses dibatalkan.");
      return;
    }
    sendFinanceEmailByTanggal(tanggalStr);
  } else {
    ui.alert("❌ Proses dibatalkan.");
  }
}



// // ✅ Konfirmasi: Kirim Email ke Tim Finance
// function showKirimEmailKeTimFinance() {
//   const ui = SpreadsheetApp.getUi();
//   const response = ui.prompt(
//     "🚀 Kirim Email ke Tim Finance",
//     "Masukkan Tanggal Transaksi yang akan dikirim ke Tim Finance (Contoh: 17/04/2025):",
//     ui.ButtonSet.OK_CANCEL
//   );

//   if (response.getSelectedButton() === ui.Button.OK) {
//     const tanggal = response.getResponseText().trim();
//     if (!tanggal) {
//       ui.alert("❗️Tanggal tidak boleh kosong.");
//       return;
//     }

//     // ✅ Simpan tanggal ke Properties lalu panggil fungsi utama
//     PropertiesService.getScriptProperties().setProperty("tanggalTransaksiKeFinance", tanggal);
//     triggerSendEmailKeFinance(); // fungsi pemicu statis (akan dibuat di file baru)
//   }
// }

function triggerSendEmailKeFinance() {
  const tanggal = PropertiesService.getScriptProperties().getProperty("tanggalTransaksiKeFinance");
  if (!tanggal) return;

  sendFinanceEmailByTanggal(tanggal); // fungsi utama, nanti ditaruh di file `sendFinanceEmailHandler.gs`
}


