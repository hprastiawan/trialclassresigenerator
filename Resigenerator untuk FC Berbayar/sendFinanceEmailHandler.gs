// ====================================================================================
// 📁 sendFinanceEmailHandler.gs
// ====================================================================================
// 📌 Fungsi utama 🚀 Kirim Email ke Tim Finance
// ====================================================================================

// ✅ Fungsi utama setelah input tanggal
function sendFinanceEmailByTanggal(tanggalStr) {
  try {
    const rows = getFinanceRowsByTanggal(tanggalStr);

    if (rows.length === 0) {
      SpreadsheetApp.getUi().alert(`❌ Tidak ditemukan data peserta untuk tanggal transaksi: ${tanggalStr}`);
      return;
    }

    // 🔥 Validasi Super Ketat
    const invalidRows = validateFinanceRowsStrict(rows);
    if (invalidRows.length > 0) {
      showFinanceValidationErrors(invalidRows);
      throw new Error("Validasi gagal, proses kirim email dihentikan.");
    }

    // ✅ Step 6: Copy file ke Folder Tanggal
    const folderTanggalUrl = copyFinanceFilesToTanggalFolder(rows);

    // ✅ Step 7-8: Kirim Email ke Finance
    sendFinanceEmail(rows, tanggalStr, folderTanggalUrl);

    // ✅ Step 9: Update status berhasil kirim
    markFinanceStatusAsSent(rows);

    // ✅ Step 10-11: Toast dan Ringkasan Alert
    showFinanceSuccessToast(rows.length, tanggalStr);

  } catch (error) {
    Logger.log("❌ ERROR di sendFinanceEmailByTanggal(): " + error);
    SpreadsheetApp.getUi().alert(`❌ Terjadi error saat proses kirim email.\n\n${error.message}`);
  }
}

// ✅ Ambil semua peserta berdasarkan tanggal input
function getFinanceRowsByTanggal(tanggalStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Kirim ke Tim Finance");
  const data = sheet.getDataRange().getValues();

  const result = [];
  const [ddInput, mmInput, yyInput] = tanggalStr.split("/");
  const tanggalInputFormat = `${parseInt(ddInput)} ${getMonthName(parseInt(mmInput))} 20${yyInput}`;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tanggalCell = row[8]; // Kolom I: Tanggal dan Jam Transaksi

    if (!tanggalCell) continue;

    const tanggalParts = tanggalCell.split(",");
    if (tanggalParts.length < 2) continue;

    const tanggalWithoutDay = tanggalParts[1].trim().split(" ")[0] + " " + tanggalParts[1].trim().split(" ")[1] + " " + tanggalParts[1].trim().split(" ")[2];

    if (tanggalWithoutDay !== tanggalInputFormat) continue;

    const idTrx = String(row[1] || "").trim();
    const nama = String(row[4] || "").trim();
    const statusBukti = row[12]; // Kolom M
    const program = lookupProgramFromFormResponses(idTrx);

    result.push({
      rowIndex: i + 1,
      id: idTrx,
      nama: nama,
      program: program,
      session: String(row[5] || "").trim(),
      tanggal: tanggalWithoutDay,
      jamTransaksi: tanggalCell,
      jumlah: row[9],
      channel: row[10],
      status: row[11],
      statusBuktiTransfer: statusBukti,
      periodeFolderId: row[13],
      folderTanggalId: row[14]
    });
  }

  return result;
}

// ✅ Helper ambil Nama Bulan Indonesia
function getMonthName(monthNumber) {
  const months = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  return months[monthNumber];
}

// ====================================================================================
// 📌 Tambahan Step 6-11 (NEW)
// ====================================================================================

// ✅ Step 6: Salin file Resi & Bukti Transfer ke Folder Tanggal
function copyFinanceFilesToTanggalFolder(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resiFolder = DriveApp.getFolderById(FINANCE_FOLDER_RESI_ID);
  const buktiFolder = DriveApp.getFolderById(FINANCE_FOLDER_TRANSFER_ID);
  const folderTanggal = DriveApp.getFolderById(rows[0].folderTanggalId);

  let totalCopied = 0;

  rows.forEach((row, index) => {
    const resiName = generateFileNameResi(row);
    const buktiName = generateFileNameBukti(row);

    const folderFiles = folderTanggal.getFiles(); // 🔍 Ambil semua file di Folder Tanggal
    const existingFileNames = [];
    while (folderFiles.hasNext()) {
      existingFileNames.push(folderFiles.next().getName());
    }

    // 🔎 Cek Resi
    if (!existingFileNames.includes(resiName)) {
      const resiFile = resiFolder.getFilesByName(resiName);
      if (resiFile.hasNext()) {
        folderTanggal.createFile(resiFile.next().getBlob());
        totalCopied++;
      }
    }

    // 🔎 Cek Bukti Transfer
    if (!existingFileNames.includes(buktiName)) {
      const buktiFile = buktiFolder.getFilesByName(buktiName);
      if (buktiFile.hasNext()) {
        folderTanggal.createFile(buktiFile.next().getBlob());
        totalCopied++;
      }
    }

    // 📢 Toast Progress
    ss.toast(`📂 Menyalin file... (${index + 1}/${rows.length} peserta)`, "Progress Copy File", 2);
    SpreadsheetApp.flush();
  });

  // 📢 Toast Selesai
  ss.toast(`✅ Total ${totalCopied} file berhasil disalin ke folder tanggal`, "Copy Selesai", 4);

  return folderTanggal.getUrl();
}


// ✅ Step 7-8: Kirim Email ke Finance
function sendFinanceEmail(rows, tanggalStr, folderTanggalUrl) {
  const jumlahPeserta = rows.length;
  const folderTanggal = DriveApp.getFolderById(rows[0].folderTanggalId);

  const subject = `[Phincon Academy] Dana Masuk Pembayaran Trial Class - ${formatTanggalIndo(tanggalStr)} 🔔`;
  const body = buildFinanceEmailBody(rows, jumlahPeserta, folderTanggalUrl);

  const mailOptions = {
    htmlBody: body,
    name: FINANCE_EMAIL_NAME
  };

  if (jumlahPeserta <= MAX_LAMPIRAN_FINANCE) {
    const attachments = [];
    const files = folderTanggal.getFiles();
    while (files.hasNext()) {
      attachments.push(files.next().getBlob());
    }
    mailOptions.attachments = attachments;
  }

  MailApp.sendEmail({
    to: FINANCE_EMAIL_TO,
    subject: subject,
    ...mailOptions
  });
}

// ✅ Step 9: Update kolom P (Status Kirim ke Finance)
function markFinanceStatusAsSent(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Kirim ke Tim Finance");

  rows.forEach(row => {
    sheet.getRange(row.rowIndex, 16).setValue("✅ Berhasil terkirim"); // Kolom P
  });
}

// ✅ Step 10-11: Toast & Ringkasan
function showFinanceSuccessToast(jumlahPeserta, tanggalStr) {
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${jumlahPeserta} peserta berhasil dikirim ke Finance`, "📤 Email Finance", 5);

  const message = 
    `✅ Proses Kirim Selesai!\n\n` +
    `📅 Tanggal Transaksi: ${formatTanggalIndo(tanggalStr)}\n` +
    `👥 Total Peserta: ${jumlahPeserta}\n` +
    `✅ Berhasil Terkirim: ${jumlahPeserta}\n` +
    `❌ Gagal Terkirim: 0`;

  SpreadsheetApp.getUi().alert(message);
}

// ====================================================================================
// 📌 Fungsi bantu Format Tanggal Indo (dd/MM/yy ➔ 17 April 2025)
// ====================================================================================
function formatTanggalIndo(dateStr) {
  const months = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  const [dd, mm, yy] = dateStr.split("/");
  return `${parseInt(dd)} ${months[parseInt(mm, 10) - 1]} 20${yy}`;
}

// ====================================================================================
// 📌 Fungsi bantu Format Rupiah
// ====================================================================================
function formatRupiah(nominal) {
  if (!nominal || nominal === "-" || nominal === "Rp 0") return "-";
  const cleaned = nominal.toString().replace(/[^0-9]/g, '');
  return `Rp ${parseInt(cleaned).toLocaleString('id-ID')},-`;
}