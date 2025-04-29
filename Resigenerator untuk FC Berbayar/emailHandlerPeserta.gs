// ====================================================================================
// 📩 Email Handler Peserta - Resi Quick Class & Email Automation System
// ====================================================================================


// ✅ Fungsi Utama: Kirim email ke peserta dari baris aktif (1 peserta)
function KirimEmailkePesertaBarisAktif() {
  const { sheet, headers, folderOutputId } = getResiSetup(); // Ambil setup resi
  const row = sheet.getActiveRange().getRow(); // Baris aktif saat ini
  const ui = SpreadsheetApp.getUi();

  // ⛔️ Validasi jika memilih header
  if (row === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ✅ Ambil seluruh data baris aktif
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ✅ Ambil nilai-nilai penting
  const email = rowData[headers.emailCol];
  const trxId = rowData[headers.trxCol];
  const regId = rowData[headers.regCol];
  const name = rowData[headers.nameCol];

  // ❌ Validasi kelengkapan data
  if (!email || !trxId || !regId) {
    sheet.getRange(row, headers.sendEmailStatusCol + 1)
      .setValue("Sending failed ❌: Data tidak lengkap");
    ui.alert("❌ Data belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi");
    return;
  }

  // ✅ Cek apakah file resi PDF tersedia
  const folder = DriveApp.getFolderById(folderOutputId);
  const fileName = generateResiFileName(rowData, headers); // Gunakan nama file standar
  const files = folder.getFilesByName(fileName);

  // ❌ File tidak ditemukan
  if (!files.hasNext()) {
    sheet.getRange(row, headers.sendEmailStatusCol + 1)
      .setValue("Sending failed ❌: File Resi tidak ditemukan di GDrive");
    ui.alert(`🚫 File resi "${fileName}" tidak ditemukan di folder`);
    return;
  }

  // ✅ Kirim email
  try {
    const pdf = files.next().getAs(MimeType.PDF);
    TemplateEmailkePeserta(rowData, headers, pdf); // Fungsi kirim email dengan template
    sheet.getRange(row, headers.sendEmailStatusCol + 1)
      .setValue("Sending completed ✅");
    ui.alert(`✅ Email berhasil dikirim ke ${name}`);
  } catch (err) {
    const msg = err.message || "Unknown error";
    sheet.getRange(row, headers.sendEmailStatusCol + 1)
      .setValue("Sending failed ❌: " + msg);
    ui.alert(`❌ Gagal mengirim email: ${msg}`);
  }
}

// ------------------------------------------------------------------------
// 📩 KIRIM EMAIL: Untuk Beberapa Peserta (Baris Terpilih)
// ------------------------------------------------------------------------
function KirimEmailkeBeberapaPeserta() {
  const { sheet, headers, folderOutputId } = getResiSetup(); // ✅ Setup awal
  const selection = sheet.getActiveRangeList();
  const ui = SpreadsheetApp.getUi();

  if (!selection) {
    ui.alert("⛔️ Tidak ada baris yang dipilih");
    return;
  }

  const folder = DriveApp.getFolderById(folderOutputId);
  const ranges = selection.getRanges();

  let processed = 0;
  let success = 0;
  let failed = 0;

  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowIndex = startRow + i;
      const row = values[i];

      const name = row[headers.nameCol];
      const email = row[headers.emailCol];
      const trxId = row[headers.trxCol];
      const regId = row[headers.regCol];

      SpreadsheetApp.getActive().toast(`📨 Mengirim ke ${name} (${processed + 1})...`);
      Utilities.sleep(300); // 🔄 Jeda pengiriman email agar tidak overload

      // ⛔️ Validasi data wajib
      if (!email || !trxId || !regId) {
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: Data tidak lengkap");
        failed++;
        processed++;
        continue;
      }

      // ✅ Ambil file resi dari Drive
      const fileName = generateResiFileName(row, headers);
      const files = folder.getFilesByName(fileName);

      if (!files.hasNext()) {
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: File Resi tidak ditemukan");
        failed++;
        processed++;
        continue;
      }

      try {
        const pdf = files.next().getAs(MimeType.PDF);
        TemplateEmailkePeserta(row, headers, pdf); // ✅ Kirim email dengan template
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending completed ✅");
        success++;
      } catch (err) {
        const msg = err.message || "Unknown error";
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: " + msg);
        failed++;
      }

      processed++;
    }
  }

  // 📊 Ringkasan akhir ke pengguna
  ui.alert(
    `📬 Proses Kirim Email Selesai\n\n✅ Berhasil: ${success}\n❌ Gagal: ${failed}\n📦 Total Diproses: ${processed}`
  );
}


// ------------------------------------------------------------------------
// 📬 FUNGSI UTAMA: Kirim Email ke Seluruh Peserta (semua data)
// ------------------------------------------------------------------------
function KirimEmailkeSemuaPeserta() {
  const { sheet, headers, folderOutputId } = getResiSetup();
  const folder = DriveApp.getFolderById(folderOutputId);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues(); // Lewati header

  let processed = 0;
  let success = 0;
  let failed = 0;

  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2; // Karena header di baris ke-1
    const row = data[i];
    const name = row[headers.nameCol];
    const email = row[headers.emailCol];
    const trxId = row[headers.trxCol];
    const regId = row[headers.regCol];

    SpreadsheetApp.getActive().toast(`📨 Mengirim ke ${name} (${processed + 1})...`);
    Utilities.sleep(300); // Jeda agar tidak timeout

    // ❌ Cek data tidak lengkap
    if (!email || !trxId || !regId) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: Data tidak lengkap");
      failed++;
      processed++;
      continue;
    }

    // ✅ Cek apakah file PDF tersedia
    const fileName = generateResiFileName(row, headers);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: File tidak ditemukan");
      failed++;
      processed++;
      continue;
    }

    try {
      const pdf = files.next().getAs(MimeType.PDF);
      TemplateEmailkePeserta(row, headers, pdf); // Kirim email dengan template
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending completed ✅");
      success++;
    } catch (err) {
      const msg = err.message || "Unknown error";
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: " + msg);
      failed++;
    }

    processed++;
  }

  // ✅ Tampilkan ringkasan
  SpreadsheetApp.getUi().alert(
    `📧 Ringkasan Pengiriman Email\n\n✅ Berhasil: ${success}\n❌ Gagal: ${failed}\n📦 Total diproses: ${processed} peserta`
  );
}


// ------------------------------------------------------------------------
// 📧 Fungsi Utama: Kirim Email ke Peserta berdasarkan Periode
// ------------------------------------------------------------------------
function KirimEmailPesertaByPeriode(periodeTarget) {
  const { sheet, headers, folderOutputId } = getResiSetup();
  const ui = SpreadsheetApp.getUi();
  const data = sheet.getDataRange().getDisplayValues();
  const folder = DriveApp.getFolderById(folderOutputId);

  let count = 0;
  let success = 0;
  let failed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 1;
    const periode = row[headers.periodeCol];

    if (periode !== periodeTarget) continue;

    const email = row[headers.emailCol];
    const trxId = row[headers.trxCol];
    const regId = row[headers.regCol];
    const name = row[headers.nameCol];

    if (!email || !trxId || !regId) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: Data tidak lengkap");
      failed++;
      continue;
    }

    const fileName = generateResiFileName(row, headers);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: File tidak ditemukan");
      failed++;
      continue;
    }

    try {
      const pdf = files.next().getAs(MimeType.PDF);
      TemplateEmailkePeserta(row, headers, pdf);
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending completed ✅");
      success++;
    } catch (err) {
      const msg = err.message || "Unknown error";
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: " + msg);
      failed++;
    }

    count++;
    SpreadsheetApp.getActiveSpreadsheet().toast(`📨 Kirim ke ${name}...`, "Progress", 3);
    Utilities.sleep(300); // ⏳ jeda antar kirim
  }

  ui.alert(`✅ Proses Kirim Email Selesai untuk Periode: ${periodeTarget}\n\n✅ ${success} berhasil\n❌ ${failed} gagal\n📦 Total diproses: ${count}`);
}

