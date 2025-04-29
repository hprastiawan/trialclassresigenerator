// ====================================================================================
// 📩 Template Email ke Peserta - Resi Quick Class & Email Automation System
// ====================================================================================

// ✅ Fungsi: Format angka ke bentuk mata uang Rupiah
function formatRupiah(amount) {
  if (!amount || amount === "-" || amount === "0" || amount === "Rp 0") return "-";
  const cleaned = amount.toString().replace(/[^0-9]/g, '');
  const formatted = cleaned.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return `Rp ${formatted},-`;
}

// ✅ Fungsi Utama: Template email untuk peserta
function TemplateEmailkePeserta(row, headers, blob) {
  const name = row[headers.nameCol];
  const email = row[headers.emailCol];
  const trxId = row[headers.trxCol];
  const regId = row[headers.regCol];
  const tgl = row[headers.tglTextCol];
  const channel = row[headers.channelCol];
  const jumlah = row[headers.jmlBayarCol];
  const program = row[headers.programCol];
  const session = row[headers.sessionCol];
  const status = "Lunas"; // 📌 Status selalu Lunas untuk sistem ini
  const tipe = "TC";

  const bannerUrl = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Email%20Header%20Banner%20-%20Phincon%20Academy/Email%20Banner%20Phincon%20Academy%202025.png';
  const logoFooter = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Phincon%20Academy%20-%20Logo%20Footer.png';
  const linkedinIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/linkedin.png';
  const instagramIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/instagram.png';
  const whatsappIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/whatsapp.png';

  // ✅ Ambil topik dari kolom P sampai kolom terakhir topik (maks 15 kolom)
  const topikList = row.slice(headers.topikStartCol, headers.topikStartCol + 15).filter(t => t);
  const topikTable = topikList.length > 0
    ? `<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-size: 14px;">${
        topikList.map((t, i) => `
        <tr style="background-color: ${i % 2 === 0 ? '#f8f8f8' : '#ffffff'};">
          <td style="border: 1px solid #ddd;">Topik ${i + 1}</td>
          <td style="border: 1px solid #ddd;">${t}</td>
        </tr>`).join('')
      }</table>`
    : '<p>(Belum ada topik diisi)</p>';

  const subject = `[Phincon Academy] ${trxId} - ${name} - TC ${program} (Session ${session}) - Lunas - Berhasil diterima 🎉`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; background-color: #f9f9f9 !important; padding: 30px; max-width: 640px; margin: auto; border-radius: 10px;">
      <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#ffffff" style="background-color: #ffffff !important; padding: 20px; border-radius: 8px; color: #333 !important;">
        <tr>
          <td><img src="${bannerUrl}" alt="Phincon Academy Banner" style="width: 100%; max-width: 600px; height: auto; border-radius: 8px;" /></td>
        </tr>
        <tr>
          <td style="padding-top: 20px;">
            <h2 style="color: #333; text-align: center;">Pembayaran TC Course Kamu Telah Berhasil Diterima</h2>
            <p>Hi <strong>${name}</strong>,</p>
            <p>Selamat, pembayaran kamu sudah berhasil kami terima.</p>
            <p style="color: #21a366; font-weight: bold;">Detail Pembayaran</p>
            <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse: collapse; font-size: 14px;">
              <tr style="background-color: #f3fdf6;"><td><strong>ID Transaksi</strong></td><td><strong>${trxId}</strong></td></tr>
              <tr><td>Nama Program</td><td>${tipe} ${program}</td></tr>
              <tr><td>Channel Pembayaran</td><td>${channel}</td></tr>
              <tr><td>Tanggal Transaksi</td><td>${tgl}</td></tr>
              <tr><td>Jumlah</td><td><strong>${formatRupiah(jumlah)}</strong></td></tr>
              <tr><td>Status Pembayaran</td><td><strong>${status}</strong></td></tr>
            </table>
            <br>
            <p style="color: #21a366; font-weight: bold;">Topik yang Diikuti</p>
            ${topikTable}
            <p style="margin-top: 20px;">Silakan temukan bukti pembayaran Kamu pada lampiran email ini.</p>
            <p>Terima kasih atas kepercayaan Kamu kepada Phincon Academy.</p>
            <p style="margin-top: 30px;">Best regards,<br><br><strong>Phincon Academy Team</strong></p>
          </td>
        </tr>
        <tr>
          <td style="padding: 30px; background-color: #f3f3f3; text-align: center; border-radius: 8px;">
            <img src="${logoFooter}" alt="Phincon Academy Logo" style="height: 40px; margin-bottom: 10px;" />
            <div style="font-size: 13px; color: #555; line-height: 1.4;">
              Gandaria 8 Office Tower, 8th Floor<br>
              Jl. Arteri Pd. Indah No.10, RT.9/RW.6 Kby. Lama Utara,<br>
              Kec. Kby. Lama, Kota Jakarta Selatan<br>
              Jakarta 12240, Indonesia
            </div>
            <div style="margin-top: 10px;">
              <a href="https://www.linkedin.com/school/phincon-academy/" target="_blank"><img src="${linkedinIcon}" style="height: 24px; margin: 0 6px;" /></a>
              <a href="https://www.instagram.com/phinconacademy" target="_blank"><img src="${instagramIcon}" style="height: 24px; margin: 0 6px;" /></a>
              <a href="https://api.whatsapp.com/send/?phone=6281119970372" target="_blank"><img src="${whatsappIcon}" style="height: 24px; margin: 0 6px;" /></a>
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;

  // ✅ Kirim email ke peserta DEV
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [blob],
    name: ' TEST EMAIL Phincon Academy',
    bcc: 'hendra.prastiawan4@gmail.com, academy@phincon.com, phinconacademy@gmail.com'
  });

  // // ✅ Kirim email ke peserta PROD
  // MailApp.sendEmail({
  //   to: email,
  //   subject: subject,
  //   htmlBody: htmlBody,
  //   attachments: [blob],
  //   name: 'Phincon Academy',
  //   bcc: 'academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com, phinconacademy@gmail.com'
  // });
}
