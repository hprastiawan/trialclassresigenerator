// ====================================================================================
// 📄 Data Loader untuk Tim Finance - Resi Quick Class & Email Automation System
// ====================================================================================

// ✅ Fungsi bantu hilangkan spasi berlebih (double space, spasi awal/akhir)
function normalizeWhitespace(str) {
  return String(str || "").replace(/\s+/g, " ").trim();
}

// ✅ Fungsi utama untuk memuat dan menggabungkan data valid ke sheet "Kirim ke Tim Finance"
function loadDataKeFinance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");
  const financeSheetName = "Kirim ke Tim Finance";
  const financeSheet = ss.getSheetByName(financeSheetName) || ss.insertSheet(financeSheetName);

  const formData = formSheet.getDataRange().getValues();
  const headers = formData[0];

  // ✅ Ambil indeks penting dari sheet Form Responses 1
  const IDX = {
    TRX: headers.indexOf("ID Transaksi"),
    REG: headers.indexOf("ID Registrasi"),
    EMAIL: headers.indexOf("Email"),
    NAMA: headers.indexOf("Nama Lengkap"),
    SESSION: headers.indexOf("Session"),
    PERIODE: headers.indexOf("Periode Pelaksanaan"),
    TOPIK: headers.indexOf("Jumlah Topik yang Diikuti"),
    TGL: headers.indexOf("Tanggal dan Jam Transaksi"),
    JUMLAH: headers.indexOf("Jumlah Pembayaran"),
    CHANNEL: headers.indexOf("Channel Pembayaran"),
    STATUS_RESI: headers.indexOf("File dalam Folder"),
    EMAIL_SENT: headers.indexOf("Send Email Status"),
    FOLDER_ID_PERIODE: headers.indexOf("Folder ID Periode")
  };

  if (Object.values(IDX).includes(-1)) {
    SpreadsheetApp.getUi().alert("❌ Kolom penting tidak ditemukan di Form Responses 1.");
    return;
  }

  // ✅ Header akhir sesuai urutan ekspektasi
  const outputHeaders = [
    "No", "ID Transaksi", "ID Registrasi", "Email", "Nama Lengkap", "Session",
    "Periode Pelaksanaan", "Jumlah Topik yang Diikuti", "Tanggal dan Jam Transaksi",
    "Jumlah Pembayaran", "Channel Pembayaran", "Status Pembayaran",
    "Status Bukti Transfer", "Folder ID Periode", "Folder ID Tanggal", "Status Kirim ke Tim Finance"
  ];

  // ✅ Tulis header jika sheet masih kosong
  if (financeSheet.getLastRow() === 0) {
    financeSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  }

  // ✅ Ambil data existing dari sheet Finance (jika ada)
  let existingData = [];
  const lastRow = financeSheet.getLastRow();
  if (lastRow > 1) {
    existingData = financeSheet.getRange(2, 1, lastRow - 1, outputHeaders.length).getValues();
  }

  const existingMap = new Map();
  existingData.forEach((row, i) => {
    const trxId = normalizeWhitespace(row[1]);
    const key = JSON.stringify(row.slice(1, 11).map(normalizeWhitespace));
    existingMap.set(trxId, { index: i + 2, key, data: row });
  });

  const rowsToWrite = [];

  // ✅ Proses semua baris dari Form Responses 1
  for (let i = 1; i < formData.length; i++) {
    const row = formData[i];
    const fileStatus = normalizeWhitespace(row[IDX.STATUS_RESI]);
    const emailStatus = normalizeWhitespace(row[IDX.EMAIL_SENT]);
    if (fileStatus !== "Ada" || !emailStatus.startsWith("Sending completed ✅")) continue;

    const trxId = normalizeWhitespace(row[IDX.TRX]);

    // ✅ Format Periode Pelaksanaan jadi "April 2025"
    const rawPeriode = row[IDX.PERIODE];
    const periodeFormatted = rawPeriode instanceof Date
      ? Utilities.formatDate(rawPeriode, Session.getScriptTimeZone(), "MMMM yyyy")
      : Utilities.formatDate(new Date(rawPeriode), Session.getScriptTimeZone(), "MMMM yyyy");

    const cleanData = [
      trxId,
      normalizeWhitespace(row[IDX.REG]),
      normalizeWhitespace(row[IDX.EMAIL]),
      normalizeWhitespace(row[IDX.NAMA]),
      normalizeWhitespace(row[IDX.SESSION]),
      periodeFormatted,
      normalizeWhitespace(row[IDX.TOPIK]),
      normalizeWhitespace(row[IDX.TGL]),
      normalizeWhitespace(row[IDX.JUMLAH]),
      row[IDX.CHANNEL] // ❗️Ambil persis tanpa title case
    ];

    const keyString = JSON.stringify(cleanData);
    const existing = existingMap.get(trxId);

    // ✅ Ambil status manual jika sudah pernah diinput
    const statusBukti = existing?.data?.[12] || "Belum di upload";
    const folderIdPeriode = normalizeWhitespace(row[IDX.FOLDER_ID_PERIODE] || existing?.data?.[13] || "");
    const tanggalTransaksi = normalizeWhitespace(row[IDX.TGL]);
    const folderIdTanggal = folderIdPeriode ? createFolderTanggalIfNeeded(folderIdPeriode, tanggalTransaksi) : "";
    const statusKirim = existing?.data?.[15] || "Belum dikirim ke Tim Finance";

    const newRow = [
      "", ...cleanData, "Lunas", statusBukti, folderIdPeriode, folderIdTanggal, statusKirim
    ];

    // ✅ Update jika berubah, skip jika sama, insert jika baru
    if (existing) {
      if (existing.key !== keyString) {
        financeSheet.getRange(existing.index, 1, 1, outputHeaders.length).setValues([newRow]);
      }
    } else {
      rowsToWrite.push(newRow);
    }
  }

  // ✅ Tambahkan baris baru jika ada
  if (rowsToWrite.length > 0) {
    const start = financeSheet.getLastRow() + 1;
    rowsToWrite.forEach((r, i) => r[0] = start + i); // No
    financeSheet.getRange(start, 1, rowsToWrite.length, outputHeaders.length).setValues(rowsToWrite);
  }

  // ✅ Styling & validasi
  financeAutoResizeAllColumnsSmart(financeSheet);
  financeAlignCenterColumns(financeSheet);
  financeUpdateNameEmailAndPaymentFormatting(financeSheet);

  // ✅ Isi ulang kolom "No" berdasarkan urutan baris
  const finalLastRow = financeSheet.getLastRow();
  if (finalLastRow > 1) {
    const noValues = Array.from({ length: finalLastRow - 1 }, (_, i) => [i + 1]);
    const noRange = financeSheet.getRange(2, 1, noValues.length, 1); // kolom A
    noRange.setValues(noValues);
    noRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  }

  SpreadsheetApp.getUi().alert("✅ Data berhasil dimuat ke sheet 'Kirim ke Tim Finance'");
}
