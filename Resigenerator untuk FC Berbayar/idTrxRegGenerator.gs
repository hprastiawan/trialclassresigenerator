// ====================================================================================
// 📁 ID Generator - Transaksi & Registrasi - Resi Quick Class Automation
// ====================================================================================

// ✅ Fungsi utama untuk generate ID Transaksi & ID Registrasi
function generateIdTrxAndReg(showAlert = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const values = sheet.getDataRange().getValues();
  const channelList = getListFromSheet("channelPembayaran");

  const trxChanges = [];
  const regChanges = [];
  let newlyGenerated = 0;

  for (let i = 1; i < values.length; i++) {
    const row = i + 1;
    const [
      , existingTrx, existingReg, , , , , session, periode, , , tanggalRaw, , , channel
    ] = values[i];

    if (!periode || !session || !tanggalRaw || !channel) continue;


    // ====================================================================================
    // 🔹 Siapkan komponen dasar dari Periode, Tanggal, dan Channel
    // ====================================================================================
    const datePeriode = new Date(periode); // handle baik string maupun date
    const month = (datePeriode.getMonth() + 1).toString(); // bulan (1-12)
    const yy = datePeriode.getFullYear().toString().slice(-2); // ambil 2 digit akhir tahun
    const periodeCode = `${month}${yy}`; // contoh: "425"
    const trxDate = Utilities.formatDate(new Date(tanggalRaw), Session.getScriptTimeZone(), "yyMMdd");
    const channelIndex = String(channelList.findIndex(c => c === channel) + 1).padStart(2, "0");

    // ====================================================================================
    // 🔹 Generate ID Transaksi
    // ====================================================================================
    const trxPrefix = `TC${periodeCode}${session}${trxDate}${channelIndex}`;
    const trxOrder = getOrderForTrx(values, i, periode, session, tanggalRaw, channel);
    const trxFinal = `${trxPrefix}${trxOrder}`;

    if (!existingTrx) {
      sheet.getRange(row, 2).setValue(trxFinal); // Kolom B
      newlyGenerated++;
    } else if (existingTrx !== trxFinal) {
      sheet.getRange(row, 2).setValue(trxFinal);
      trxChanges.push(values[i][4]); // Nama
    }

    // ====================================================================================
    // 🔹 Generate ID Registrasi
    // ====================================================================================
    const regSuffix = (existingReg && existingReg.length >= 8)
      ? existingReg.slice(4, 8)
      : generateRandomCode(4);

    const regOrder = getOrderForReg(values, i, periode);
    const regFinal = `${periodeCode}${regSuffix}${regOrder}`;

    if (!existingReg) {
      sheet.getRange(row, 3).setValue(regFinal); // Kolom C
      newlyGenerated++;
    } else if (existingReg !== regFinal) {
      sheet.getRange(row, 3).setValue(regFinal);
      regChanges.push(values[i][4]); // Nama
    }
  }

  // ====================================================================================
  // 🔔 Tampilkan ringkasan alert (jika diminta)
  // ====================================================================================
  if (showAlert) {
    let alertMsg = "";

    if (newlyGenerated > 0 && regChanges.length === 0 && trxChanges.length === 0) {
      alertMsg = "✅ ID Transaksi & Registrasi berhasil dibuat dan diupdate";
    } else if (regChanges.length > 0 || trxChanges.length > 0) {
      const list = [];
      regChanges.forEach(n => list.push(`✅ ID Registrasi atas nama ${n} mengalami perubahan`));
      trxChanges.forEach(n => list.push(`✅ ID Transaksi atas nama ${n} mengalami perubahan`));
      alertMsg = list.join("\n");
    } else {
      alertMsg = "ℹ️ ID Transaksi / ID Registrasi tidak ada perubahan dari data yang ada saat ini";
    }

    SpreadsheetApp.getUi().alert(alertMsg);
  }
}

// ====================================================================================
// 🔧 Utilitas & Helper Function
// ====================================================================================

// ✅ Ambil angka bulan dari nama bulan (e.g. "April 2025" → 4)
function getMonthNumberFromPeriode(periode) {
  const monthMap = {
    "Januari": 1, "Februari": 2, "Maret": 3, "April": 4, "Mei": 5, "Juni": 6,
    "Juli": 7, "Agustus": 8, "September": 9, "Oktober": 10, "November": 11, "Desember": 12
  };
  const periodeStr = String(periode || "");
  return monthMap[periodeStr.split(" ")[0]] || 0;
  //return monthMap[periode.split(" ")[0]] || 0;
}

// ✅ Hitung urutan ID Transaksi berdasarkan kombinasi periode + session + tanggal transaksi
function getOrderForTrx(values, currentIndex, periode, session, tanggal) {
  let count = 0;
  const refTanggal = new Date(tanggal).getTime();

  for (let i = 1; i < currentIndex; i++) {
    const [,,,,,, , sess, peri, , , tgl] = values[i];
    if (!sess || !peri || !tgl) continue;

    const sameSession = sess == session;
    const samePeriode = new Date(peri).getTime() === new Date(periode).getTime();
    const sameTanggal = new Date(tgl).getTime() === refTanggal;

    if (sameSession && samePeriode && sameTanggal) count++;
  }

  return String(count + 1).padStart(4, "0");
}

// ✅ Hitung urutan ID Registrasi berdasarkan periode pelaksanaan saja
function getOrderForReg(values, currentIndex, periode) {
  let count = 0;
  const refDate = new Date(periode).getTime();

  for (let i = 1; i < currentIndex; i++) {
    const peri = values[i][8]; // Kolom I
    if (!peri) continue;

    const compDate = new Date(peri).getTime();
    if (compDate === refDate) count++;
  }

  return String(count + 1).padStart(4, "0");
}


// ✅ Buat kode acak alfanumerik (default: 4 karakter)
function generateRandomCode(length = 4) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  return Array.from({ length }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
}