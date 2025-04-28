// ====================================================================================
// 📁 Menu Handler - Resi Quick Class & Email Automation System
// ====================================================================================

// ✅ Fungsi untuk menampilkan menu utama di Google Sheets
function createResiAutomationMenus() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🧾 Resi Automation")
    .addItem("🔄 Refresh Halaman", "showRefreshResiConfirmation")
    .addItem("🔢 Generate ID Transaksi & Registrasi", "showGenerateIdTrxReg")
    .addSubMenu(
      ui.createMenu("🗃️ Manajemen Direktori")
        .addItem("📁 Buat Folder per Periode", "showCreateFolderPerPeriode")
        .addItem("🗑️ Bersihkan Folder kosong", "showEraseEmptyFolder")
        .addItem("📊 Generate Summary Peserta", "showGenerateSummaryPeserta")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("🧾 Buat Resi")
        .addItem("📕 Untuk Baris Ini", "showResiBarisAktif")
        .addItem("📘📗 Untuk Beberapa Baris", "showResiBeberapaBarisAktif")
        .addItem("📥 Generate Resi per Periode", "showGenerateResiByPeriode")
        .addItem("📚 Cetak Resi untuk Seluruh Data", "showResiSeluruhData")
    )
    .addSubMenu(
      ui.createMenu("📮 Kirim Email ke Peserta")
        .addItem("👤 Kirim Resi untuk 1 Peserta", "showKirimEmailkePesertaBarisAktif")
        .addItem("👥 Kirim Resi untuk Beberapa Peserta", "showKirimEmailkeBeberapaPeserta")
        .addItem("👥👥 Kirim Resi untuk Seluruh Peserta", "showKirimEmailkeSemuaPeserta")
        .addItem("📬 Kirim Email per Periode", "showKirimEmailPesertaByPeriode")
    )
    .addSubMenu(
      ui.createMenu("📬 Kirim Email ke Tim Finance")
        .addItem("🧲 Get Data Peserta", "showGetDataPeserta")
        .addItem("📤 Upload Bukti Transfer", "showUploadBuktiTransferSidebar")
        .addItem("🚀 Kirim Email ke Tim Finance", "showKirimEmailKeTimFinance")
    )
    .addToUi();
}
