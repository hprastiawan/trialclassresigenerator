
# 📑 Trial Class Resi Generator (Google Apps Script)

Sistem otomatisasi berbasis Google Apps Script untuk mengelola pembuatan, pengiriman, dan pelaporan **resi pembayaran Trial Class** melalui Google Sheets dan Google Drive.  
Dilengkapi dengan fitur pembuatan ID, pembuatan resi, pengiriman email ke peserta, pengumpulan bukti transfer, hingga pengiriman email ke tim keuangan (Finance).

---

## 🚀 Fitur Utama

### 🔹 Generate ID Transaksi & ID Registrasi
- Membuat ID Transaksi unik dan ID Registrasi otomatis berdasarkan data peserta.
- Menu: **Generate ID Transaksi & Registrasi**

### 🔹 Manajemen Direktori Google Drive
- Membuat struktur folder penyimpanan otomatis:
  - Folder Periode
  - Folder Tanggal Transaksi
- Menu: **Manajemen Direktori → Buat Folder Periode/Tanggal**

### 🔹 Buat Resi Pembayaran
- Menghasilkan file resi pembayaran (PDF) untuk peserta berdasarkan pilihan:
  - Untuk 1 Baris
  - Untuk Beberapa Baris
  - Per Periode
  - Seluruh Data
- Menu: **Buat Resi**

### 🔹 Kirim Email ke Peserta
- Mengirimkan email otomatis berisi resi ke email peserta.
- Pilihan kirim:
  - Kirim Resi untuk 1 Peserta
  - Kirim Resi untuk Beberapa Peserta
  - Kirim Resi untuk Semua Peserta
  - Kirim Email per Periode
- Menu: **Kirim Email ke Peserta**

### 🔹 Upload Bukti Transfer Peserta
- Upload file bukti pembayaran peserta ke Google Drive.
- Update status bukti transfer di Sheet.
- Menu: **Kirim Email ke Tim Finance → Upload Bukti Transfer**

### 🔹 Kirim Email ke Tim Finance
- Validasi ketat file resi dan bukti transfer.
- Menyalin file Resi & Bukti Transfer ke Folder Tanggal.
- Mengirimkan email rekap keuangan Trial Class ke tim Finance:
  - Jika peserta ≤ 5: file dikirim sebagai lampiran.
  - Jika peserta > 5: kirim link folder Google Drive.
- Update status pengiriman di Sheet.
- Menu: **Kirim Email ke Tim Finance → Kirim Email ke Tim Finance**

---

## 📋 Struktur Menu

Berikut tampilan struktur menu utama di Google Sheets:

```plaintext
Resi Automation
├── Refresh Halaman
├── Generate ID Transaksi & Registrasi
├── Manajemen Direktori
│   ├── Buat Folder Periode
│   └── Buat Folder Tanggal
├── Buat Resi
│   ├── Untuk Baris Ini
│   ├── Untuk Beberapa Baris
│   ├── Generate Resi per Periode
│   └── Cetak Resi untuk Seluruh Data
├── Kirim Email ke Peserta
│   ├── Kirim Resi untuk 1 Peserta
│   ├── Kirim Resi untuk Beberapa Peserta
│   ├── Kirim Resi untuk Seluruh Peserta
│   └── Kirim Email per Periode
└── Kirim Email ke Tim Finance
    ├── Get Data Peserta
    ├── Upload Bukti Transfer
    └── 🚀 Kirim Email ke Tim Finance
```

---

## 🛠️ Flow Proses Kirim Email ke Finance

1. Klik **Kirim Email ke Tim Finance → Kirim Email ke Tim Finance**.
2. Input Tanggal Transaksi (format: `dd/MM/yy`).
3. Sistem akan:
   - Validasi status bukti transfer dan ketersediaan file.
   - Menyalin file ke folder tanggal jika belum ada.
   - Mengirim email ke Finance.
4. Setelah sukses:
   - Update status "✅ Berhasil terkirim" di Sheet.
   - Tampilkan alert ringkasan pengiriman.

---

## 📦 Teknologi yang Digunakan
- **Google Apps Script**
- **Google Sheets**
- **Google Drive API**
- **MailApp Service** untuk pengiriman email

---

## 🖌️ Desain Template Email
- Header banner visual
- Judul "Notifikasi Dana Masuk Trial Class"
- Tabel peserta dengan informasi lengkap
- Lampiran file atau link folder Google Drive
- Footer profesional (Logo, Alamat, Link Social Media)

---

## ✅ Status
🟢 **Project aktif dan berjalan dengan baik.**  
Terus diperbarui untuk meningkatkan ketepatan validasi, pengelolaan file, dan otomatisasi pengiriman.
