// ====================================================================================
// 📁 constantsEmailFinance.gs
// ====================================================================================
// 📌 Tempat menyimpan ID folder, batas lampiran, dan info email default
// ====================================================================================

// ====================================================================================
// 📂 ID FOLDER GOOGLE DRIVE
// ====================================================================================

// ✅ Folder tempat file RESI yang digenerate disimpan
const FINANCE_FOLDER_RESI_ID = "1LwOj_U3zqf8YYFUIbEXvScQARQR9nq8R";

// ✅ Folder tempat file BUKTI TRANSFER yang diupload peserta
const FINANCE_FOLDER_TRANSFER_ID = "10mDu1jxvA4CxO-vuFkfLYq0GdCUjZEEY";

// ====================================================================================
// 📎 BATAS LAMPIRAN EMAIL
// ====================================================================================

// ✅ Maksimum jumlah peserta untuk melampirkan file langsung (jika lebih, kirim link saja)
const MAX_LAMPIRAN_FINANCE = 5;

// ====================================================================================
// 📧 KONFIGURASI EMAIL TUJUAN
// ====================================================================================

// ✅ Email utama tujuan pengiriman notifikasi
const FINANCE_EMAIL_TO = "hendra.prastiawan2@gmail.com";

// ✅ Email tambahan untuk CC (bisa dikosongkan jika tidak ada)
const FINANCE_EMAIL_CC = "";

// ✅ Email tambahan untuk BCC (bisa dikosongkan jika tidak ada)
const FINANCE_EMAIL_BCC = "";

// ✅ Nama pengirim yang akan tampil di email
const FINANCE_EMAIL_NAME = "Phincon Academy System";
