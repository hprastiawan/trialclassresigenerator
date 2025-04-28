// ====================================================================================
// 📌 Fungsi pembuat HTML Body Email ke Finance
// ====================================================================================
function buildFinanceEmailBody(rows, jumlahPeserta, folderTanggalUrl) {
  const bannerUrl = "https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Email%20Header%20Banner%20-%20Phincon%20Academy/Email%20Banner%20Phincon%20Academy%202025.png";
  const logoFooterUrl = "https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Phincon%20Academy%20-%20Logo%20Footer.png";
  const linkedinIcon = "https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/linkedin.png";
  const instagramIcon = "https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/instagram.png";
  const whatsappIcon = "https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/whatsapp.png";

  const tableRows = rows.map(row => `
    <tr>
      <td>${row.id}</td>
      <td>${row.nama}</td>
      <td>${row.program}</td>
      <td>Session ${row.session}</td>
      <td>${row.jamTransaksi}</td>
      <td>${formatRupiah(row.jumlah)}</td>
      <td>${row.channel}</td>
      <td>${row.status}</td>
    </tr>
  `).join("");

  const attachmentInfo = jumlahPeserta <= MAX_LAMPIRAN_FINANCE
    ? `<p style="margin-top: 20px;">Terlampir bukti transfer dan resi pembayaran peserta.</p>`
    : `<p style="margin-top: 20px;">Berikut pranala Resi dan Bukti Transfer dari peserta di atas: <a href="${folderTanggalUrl}" target="_blank">📂 Google Drive Folder</a></p>`;

  return `
    <div style="font-family: Arial, sans-serif; background-color: #f9f9f9; padding: 40px; max-width: 860px; margin: auto; border-radius: 10px;">
      <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border-radius: 10px; padding: 30px;">
        <tr>
          <td>
            <img src="${bannerUrl}" alt="Header Banner Phincon Academy" style="width: 100%; max-width: 800px; border-radius: 8px;" />
          </td>
        </tr>
        <tr>
          <td style="padding-top: 30px;">
            <h2 style="text-align: center; color: #333;">Notifikasi Dana Masuk Trial Class</h2>
            <p>Halo Finance Team,</p>
            <p>Berikut adalah data dana masuk untuk pembayaran Trial Class untuk total <strong>${jumlahPeserta}</strong> peserta:</p>
            <table width="100%" border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; font-size: 14px; margin-top: 20px;">
              <thead style="background-color: #f3f3f3;">
                <tr>
                  <th>ID Transaksi</th>
                  <th>Nama Lengkap</th>
                  <th>Program</th>
                  <th>Session</th>
                  <th>Tanggal & Jam</th>
                  <th>Jumlah</th>
                  <th>Channel</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>

            ${attachmentInfo}

            <p style="margin-top: 40px;">Best regards,<br><strong>Phincon Academy Team</strong></p>
          </td>
        </tr>
        <tr>
          <td style="padding: 30px; background-color: #f3f3f3; text-align: center; border-radius: 8px;">
            <img src="${logoFooterUrl}" alt="Phincon Academy Footer Logo" style="height: 40px; margin-bottom: 10px;" />
            <div style="font-size: 13px; color: #555; line-height: 1.5; margin-top: 10px;">
              Gandaria 8 Office Tower, 8th Floor<br/>
              Jl. Arteri Pd. Indah No.10, Kebayoran Lama, Jakarta Selatan<br/>
              Jakarta 12240, Indonesia
            </div>
            <div style="margin-top: 15px;">
              <a href="https://www.linkedin.com/school/phincon-academy/" target="_blank">
                <img src="${linkedinIcon}" alt="LinkedIn" style="height: 24px; margin: 0 6px;" />
              </a>
              <a href="https://www.instagram.com/phinconacademy" target="_blank">
                <img src="${instagramIcon}" alt="Instagram" style="height: 24px; margin: 0 6px;" />
              </a>
              <a href="https://api.whatsapp.com/send/?phone=6281119970372" target="_blank">
                <img src="${whatsappIcon}" alt="WhatsApp" style="height: 24px; margin: 0 6px;" />
              </a>
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;
}