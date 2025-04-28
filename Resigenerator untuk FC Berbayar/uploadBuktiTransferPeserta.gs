// ====================================================================================
// 📤 Upload Bukti Transfer - Sidebar - Resi Quick Class & Email Automation System
// ====================================================================================

// ✅ Fungsi untuk membuka sidebar upload bukti transfer dari baris aktif
function openUploadDialogForActiveRow() {
  const html = HtmlService.createHtmlOutputFromFile("uploadBuktiTransferPesertaWeb")
    .setWidth(400)
    .setHeight(300)
    .setTitle("Upload Bukti Transfer");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ✅ Decode base64 → buat blob → teruskan ke fungsi upload utama
function uploadBase64File(base64Data, fileName, mimeType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kirim ke Tim Finance");
  const rowIndex = sheet.getActiveCell().getRow();

  // ⛔️ Jika baris header atau tidak valid
  if (rowIndex < 2) {
    SpreadsheetApp.getUi().alert("⛔️ Pilih baris data peserta terlebih dahulu.");
    return;
  }

  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  return uploadBuktiTransferFromDialog(blob, rowIndex);
}

// ✅ Upload bukti transfer ke Drive & update status
function uploadBuktiTransferFromDialog(blob, rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Kirim ke Tim Finance");
  const formSheet = ss.getSheetByName("Form Responses 1");

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const clean = txt => String(txt || "").trim().replace(/\s+/g, " ");

  const idTransaksi = clean(row[1]);   // Kolom B
  const nama = clean(row[4]);          // Kolom E
  const session = clean(row[5]);       // Kolom F

  // 🔍 Ambil nama program dari Form Responses 1
  const formData = formSheet.getDataRange().getValues();
  const header = formData[0];
  const idxTrx = header.indexOf("ID Transaksi");
  const idxNama = header.indexOf("Nama Lengkap");
  const idxProgram = header.indexOf("Nama Program");

  let namaProgram = "Program Tidak Ditemukan";
  for (let i = 1; i < formData.length; i++) {
    if (clean(formData[i][idxTrx]) === idTransaksi && clean(formData[i][idxNama]) === nama) {
      namaProgram = formData[i][idxProgram] || namaProgram;
      break;
    }
  }

  // ✅ Format nama file
  const fileNama = `[Phincon Academy] TC Bukti Transfer - ${idTransaksi} - ${nama} - ${namaProgram} - Session ${session} - Lunas`;

  // ✅ Simpan ke Drive folder
  const folder = DriveApp.getFolderById("10mDu1jxvA4CxO-vuFkfLYq0GdCUjZEEY"); // Bukti Transfer FC Berbayar
  folder.createFile(blob.setName(fileNama));

  // ✅ Update status di kolom M (kolom ke-13)
  sheet.getRange(rowIndex, 13).setValue("✅ Berhasil di-upload");

  // ✅ Alert sukses
  SpreadsheetApp.getUi().alert("📤 Upload Bukti Bayar", `1 file berhasil diupload:\n${fileNama}`, SpreadsheetApp.getUi().ButtonSet.OK);
}
