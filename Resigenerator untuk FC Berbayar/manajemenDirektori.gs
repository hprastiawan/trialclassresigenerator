// ====================================================================================
// 📁 Manajemen Direktori - Resi Quick Class & Email Automation System
// ====================================================================================

/**
 * ✅ Fungsi utama untuk membuat folder berdasarkan kolom "Periode Pelaksanaan"
 * Folder akan dibuat di dalam folder Level 2: "Resi & Bukti Transfer FC Berbayar"
 * ID folder disimpan di kolom AJ, dan status folder disimpan di kolom AL
 */
function createFoldersPerPeriode(updateStatus = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();
  const parentFolder = DriveApp.getFolderById(FOLDER_MERGED_RESI_TRANSFER);
  const existingFolders = parentFolder.getFolders();

  const folderMap = new Map();
  while (existingFolders.hasNext()) {
    const folder = existingFolders.next();
    folderMap.set(folder.getName().toLowerCase(), folder.getId());
  }

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const rawPeriode = data[i][8]; // Kolom I
    const folderIdCell = sheet.getRange(row, 36); // Kolom AJ
    const folderStatusCell = sheet.getRange(row, 38); // Kolom AL

    if (!rawPeriode) continue;

    let periodeStr = "";
    if (rawPeriode instanceof Date) {
      periodeStr = Utilities.formatDate(rawPeriode, Session.getScriptTimeZone(), "MMMM yyyy");
    } else if (typeof rawPeriode === "string") {
      periodeStr = rawPeriode.trim();
    } else {
      continue;
    }

    const folderKey = periodeStr.toLowerCase();

    if (folderMap.has(folderKey)) {
      const folderId = folderMap.get(folderKey);
      if (updateStatus) {
        if (!folderIdCell.getValue()) folderIdCell.setValue(folderId);
        folderStatusCell.setValue("✅ Folder ditemukan"); // ✅ selalu ditulis
      }
    } else {
      const newFolder = parentFolder.createFolder(periodeStr);
      folderMap.set(folderKey, newFolder.getId());
      folderIdCell.setValue(newFolder.getId());
      folderStatusCell.setValue("✅ Folder dibuat");
      Logger.log(`📁 Folder baru dibuat: ${periodeStr}`);
    }
  }

  SpreadsheetApp.getUi().alert("✅ Folder per Periode berhasil diproses.");
}



/**
 * 🗑️ Hapus semua folder kosong dari dalam folder Level 2
 * Target: folder di dalam FOLDER_MERGED_RESI_TRANSFER
 * Juga bersihkan kolom AJ (Folder ID Periode) dan AL (Status Folder Periode di Folder Level 3)
 * Menampilkan log baris dan nama folder yang dibersihkan
 */
function eraseEmptyFoldersFromMergedFolder() {
  const parentFolder = DriveApp.getFolderById(FOLDER_MERGED_RESI_TRANSFER);
  const folders = parentFolder.getFolders();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();

  const toDeleteMap = new Map(); // folderId → folderName
  let countDeleted = 0;

  // 🔍 Scan folder kosong di Drive
  while (folders.hasNext()) {
    const folder = folders.next();
    const files = folder.getFiles();
    const subFolders = folder.getFolders();

    // ✅ Cek apakah folder benar-benar kosong (tidak ada file dan folder di dalamnya)
    if (!files.hasNext() && !subFolders.hasNext()) {
      folder.setTrashed(true);
      toDeleteMap.set(folder.getId(), folder.getName());
      Logger.log(`🗑️ Folder kosong dihapus: ${folder.getName()}`);
      countDeleted++;
    }
  }

  // 🧹 Bersihkan dari Spreadsheet (Kolom AJ dan AL) jika folder ID cocok dengan yang dihapus
  const affectedRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const folderId = data[i][35]; // Kolom AJ - Folder ID Periode

    if (folderId && toDeleteMap.has(folderId)) {
      sheet.getRange(row, 36).clearContent(); // Kolom AJ
      sheet.getRange(row, 38).clearContent(); // Kolom AL
      affectedRows.push(`• Baris ${row}: ${toDeleteMap.get(folderId)}`);
    }
  }

  // 🧾 Ringkasan hasil
  const summary = affectedRows.length > 0
    ? `✅ ${countDeleted} folder kosong telah dipindahkan ke Trash.\n\n📋 Folder & data yang dibersihkan:\n${affectedRows.join("\n")}`
    : `✅ Tidak ada folder kosong yang ditemukan.`;

  SpreadsheetApp.getUi().alert(summary);
}

// ====================================================================================
// 📄 Summary Peserta Sheet Generator
// ====================================================================================

/**
 * ✅ Buat atau perbarui sheet "Summary Peserta" dengan data peserta dan file per periode
 */
function generateSummaryPesertaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");
  const summarySheetName = "Summary Peserta";

  // Buat atau bersihkan sheet summary
  let sheet = ss.getSheetByName(summarySheetName);
  if (!sheet) {
    sheet = ss.insertSheet(summarySheetName);
  } else {
    sheet.clear();
  }

  const data = formSheet.getDataRange().getValues();
  const folderMap = new Map();
  const parentFolder = DriveApp.getFolderById(FOLDER_MERGED_RESI_TRANSFER);
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    folderMap.set(f.getName().toLowerCase(), f);
  }

  // Kumpulkan data jumlah peserta per periode
  const summaryMap = new Map();
  for (let i = 1; i < data.length; i++) {
    const rawPeriode = data[i][8]; // Kolom I
    if (!rawPeriode) continue;

    let periodeStr = "";
    if (rawPeriode instanceof Date) {
      periodeStr = Utilities.formatDate(rawPeriode, Session.getScriptTimeZone(), "MMMM yyyy");
    } else if (typeof rawPeriode === "string") {
      periodeStr = rawPeriode.trim();
    }
    if (!periodeStr) continue;

    const key = periodeStr;
    const current = summaryMap.get(key) || { peserta: 0, file: 0, url: "" };
    current.peserta += 1;

    const folder = folderMap.get(periodeStr.toLowerCase());
    if (folder) {
      current.url = folder.getUrl();
      const files = folder.getFiles();
      let count = 0;
      while (files.hasNext()) {
        files.next();
        count++;
      }
      current.file = count;
    }

    summaryMap.set(key, current);
  }

  // Tulis ke sheet
  const headers = ["No", "Periode Pelaksanaan", "Jumlah Peserta", "Jumlah File", "Link Folder"];
  sheet.appendRow(headers);

  const sorted = Array.from(summaryMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  sorted.forEach(([periode, { peserta, file, url }], idx) => {
    sheet.appendRow([idx + 1, periode, peserta, file, url]);
  });

  autoResizeAllColumnsSmart(sheet);
  
  // ✅ Styling tambahan khusus header untuk sheet Summary Peserta
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment("center");
  headerRange.setWrap(true);
  headerRange.setBackground("#d9ead3");

  SpreadsheetApp.getUi().alert("✅ Sheet 'Summary Peserta' berhasil diperbarui!");
}

