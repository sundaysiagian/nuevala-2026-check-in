  function doGet(e) {
    try {
      const action      = e.parameter.action;
      const namaPenukar = e.parameter.namaPenukar;
      let result;

      if (action === 'qr') {
        result = processQrCodeData(e.parameter.qrData, namaPenukar);
      } else if (action === 'code') {
        result = processAlphanumericCode(e.parameter.code, namaPenukar);
      } else {
        result = { status: 'error', message: 'Action tidak dikenal.' };
      }

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  function doPost(e) {
    try {
      const data   = JSON.parse(e.postData.contents);
      let   result;

      if (data.action === 'qr') {
        result = processQrCodeData(data.qrData, data.namaPenukar);
      } else if (data.action === 'code') {
        result = processAlphanumericCode(data.code, data.namaPenukar);
      } else {
        result = { status: 'error', message: 'Action tidak dikenal.' };
      }

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ============================================================
  // HELPERS
  // ============================================================

  /**
   * Menghasilkan kode alfanumerik unik sepanjang 4 karakter.
   *
   * Karakter yang dipakai sengaja dipilih untuk menghindari ambiguitas visual:
   * huruf O dan angka 0 mirip, huruf I dan angka 1 juga mirip — jadi dikeluarkan.
   * Tujuannya biar panitia dan pembeli ga salah baca kode saat hari-H.
   *
   * @param {Set} existingCodes - Kumpulan kode yang sudah ada di spreadsheet,
   *                              supaya kode yang dihasilkan dijamin unik (tidak bentrok).
   * @returns {string} Kode unik 4 karakter, huruf besar semua.
   */
  function generateUniqueCode(existingCodes) {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // karakter ambigu dihapus: 0, O, 1, I
    let code, attempts = 0;
    do {
      code = '';
      for (let i = 0; i < 4; i++) code += chars[Math.floor(Math.random() * chars.length)];
      attempts++;
    } while (existingCodes.has(code) && attempts < 1000);
    return code;
  }

  /**
   * Mengambil gambar banner tiket dari Google Drive berdasarkan tier.
   *
   * Gambar disimpan di Google Drive (bukan di-embed langsung ke kode)
   * supaya lebih gampang diganti tanpa perlu utak-atik kode.
   * Cara ganti gambarnya ada di bagian FAQ di README.
   *
   * @param {string} tier - Jenis tiket: 'GOLD', 'SILVER', atau 'BRONZE' (case-insensitive).
   * @returns {Blob|null} Blob gambar siap pakai, atau null kalau gagal diambil.
   */
  function getTicketImageBlob(tier) {
    const urls = {
      'GOLD':   'https://drive.google.com/uc?export=view&id=1fMIsw6tT6oPX9946AcN7OWz3tJbsz-7w',
      'SILVER': 'https://drive.google.com/uc?export=view&id=10MLsrAeCnAhFWfmNJ514GiwS7lnytk5q',
      'BRONZE': 'https://drive.google.com/uc?export=view&id=1wGA35mtbN-W2DmbVXQumJtSBLFPqomIL'
    };
    const url = urls[tier.toUpperCase()];
    if (!url) return null;
    try {
      return UrlFetchApp.fetch(url).getBlob().setName('ticket-banner.jpg');
    } catch(e) {
      Logger.log('Gagal mengambil gambar untuk tier ' + tier + ': ' + e);
      return null;
    }
  }

  /**
   * Membuat isi HTML untuk email tiket yang dikirimkan ke pembeli.
   *
   * Fungsi ini menghasilkan dua blok utama:
   *   1. Kartu tiket — menampilkan info acara, nama pembeli, show, jumlah, dan total
   *   2. Blok QR code + kode alfanumerik — untuk penukaran tiket di pintu masuk
   *
   * Warna dan label tiket berbeda-beda tergantung tier (Gold, Silver, Bronze).
   * Informasi acara (tanggal, tempat, kontak) ditulis langsung di sini —
   * kalau ada perubahan, cari dan edit bagian tabel di dalam fungsi ini.
   *
   * @param {string}  nama      - Nama lengkap pembeli.
   * @param {string}  show      - Show yang dipilih (misal: "Show 1 – Siang").
   * @param {string}  jumlah    - Jumlah tiket yang dipesan.
   * @param {string}  total     - Total pembayaran (akan diformat ke Rupiah).
   * @param {string}  tier      - Jenis tiket: 'GOLD', 'SILVER', atau 'BRONZE'.
   * @param {string}  kode      - Kode alfanumerik 4 karakter untuk verifikasi manual.
   * @param {boolean} hasImage  - true kalau blob gambar banner berhasil diambil.
   * @returns {string} String HTML lengkap untuk dikirim via MailApp.
   */
  function buildEmailHtml(nama, show, jumlah, total, tier, kode, hasImage) {
    // Konfigurasi tampilan per tier: warna latar, warna aksen, dan label
    const cfg = {
      GOLD:   { bg: '#070605', accent: '#ECC976', label: 'Exclusive Gold Access' },
      SILVER: { bg: '#2E3137', accent: '#72B0B6', label: 'Standard Silver Access' },
      BRONZE: { bg: '#413539', accent: '#D66E40', label: 'Bronze Access' }
    }[tier.toUpperCase()] || { bg: '#413539', accent: '#D66E40', label: 'Access' };

    // Banner gambar hanya ditampilkan kalau blob-nya berhasil diambil
    const imgHtml = hasImage
      ? `<img src="cid:ticketBanner" width="400" style="display:block;width:100%;height:180px;object-fit:cover;" alt="${tier} Ticket">`
      : '';

    return `
  <div style="font-family:'Segoe UI',Arial,sans-serif;background:#f0f0f0;padding:30px;">

    <div style="max-width:400px;margin:0 auto;background:${cfg.bg};border-radius:15px;overflow:hidden;border-left:8px solid ${cfg.accent};box-shadow:0 10px 30px rgba(0,0,0,0.3);">
      ${imgHtml}
      <div style="padding:20px;color:white;">
        <span style="color:${cfg.accent};text-transform:uppercase;font-size:11px;letter-spacing:3px;display:block;margin-bottom:5px;">${cfg.label}</span>
        <h2 style="margin:0 0 4px;font-size:28px;font-weight:300;color:white;">NUEVALA</h2>
        <p style="font-size:13px;opacity:0.7;margin:0 0 16px;">Voices Beyond The Walls</p>
        <p style="font-size:13px;opacity:0.9;margin:0 0 14px;">Dear <strong>${nama}</strong>, pembayaran Anda telah dikonfirmasi.</p>

        <!-- Informasi acara — ubah di sini kalau ada perubahan tanggal, tempat, atau kontak -->
        <table style="width:100%;font-size:13px;color:white;border-collapse:collapse;">
          <tr><td style="opacity:0.6;padding:3px 12px 3px 0;white-space:nowrap;">Tanggal</td><td>Sabtu, 16 Mei 2026</td></tr>
          <tr><td style="opacity:0.6;padding:3px 12px 3px 0;white-space:nowrap;">Tempat</td><td>Aula Barat ITB, Jl. Ganesha No. 10</td></tr>
          <tr><td style="opacity:0.6;padding:3px 12px 3px 0;white-space:nowrap;">Show</td><td><strong>${show}</strong></td></tr>
          <tr><td style="opacity:0.6;padding:3px 12px 3px 0;white-space:nowrap;">Jumlah</td><td><strong>${jumlah} tiket</strong></td></tr>
          <tr><td style="opacity:0.6;padding:3px 12px 3px 0;white-space:nowrap;">Total</td><td><strong>${formatRupiah(total)}</strong></td></tr>
        </table>
      </div>
      <div style="border-top:1px dashed rgba(255,255,255,0.2);padding:12px 20px;">
        <table style="width:100%;"><tr>
        <td style="color:white;font-size:11px;">Pertanyaan? WhatsApp <strong>081263153382</strong> (Indah)</td>
        </tr></table>
      </div>
    </div>

    <div style="max-width:400px;margin:20px auto;background:white;border-radius:10px;padding:20px;text-align:center;">
      <p style="color:#555;font-size:13px;margin:0 0 12px;">Tunjukkan QR Code ini saat penukaran tiket:</p>
      <img src="cid:qrcode" alt="QR Code" width="180" height="180" style="display:block;margin:0 auto;">
      <p style="color:#555;font-size:13px;margin:18px 0 8px;">Jika QR Code tidak terbaca, gunakan <strong>kode verifikasi</strong> ini:</p>
      <div style="font-size:34px;font-weight:bold;letter-spacing:10px;color:#111;margin:8px 0;font-family:monospace;">${kode}</div>
      <p style="color:#bbb;font-size:11px;margin:6px 0 0;">Simpan kode ini sebagai cadangan.</p>
    </div>

    <div style="max-width:400px;margin:0 auto;text-align:center;color:#888;font-size:12px;line-height:2;">
      <p>Info lebih lanjut: <strong>@psmitbconcert</strong></p>
      <p style="margin-top:6px;">Sampai jumpa di Nuevala! ^^</p>
    </div>

  </div>

  </div>`;
  }

  /**
   * Memformat angka menjadi format Rupiah (Rp1.000.000,00).
   * Dipakai untuk menampilkan kolom total pembayaran di dalam email.
   *
   * @param {string|number} value - Nilai yang akan diformat (boleh mengandung karakter non-angka).
   * @returns {string} String dalam format Rupiah, atau nilai aslinya kalau parsing gagal.
   */
  function formatRupiah(value) {
    const num = parseFloat(String(value).replace(/[^0-9]/g, ''));
    if (isNaN(num)) return value;
    return 'Rp' + num.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ============================================================
  // SEND EMAILS
  // ============================================================

  /**
   * Fungsi utama pengiriman email tiket — dipanggil via menu Pengiriman Email > Kirim Email Tiket.
   *
   * Alur kerjanya:
   *   1. Baca semua data dari spreadsheet aktif
   *   2. Untuk setiap baris, cek dua syarat: pembayaran sudah diverifikasi (checkbox = true)
   *      dan email belum pernah dikirim (Status Pengiriman bukan 'Terkirim')
   *   3. Kalau belum punya kode alfanumerik, generate kode baru yang unik
   *   4. Buat QR code via API pihak ketiga (qrserver.com)
   *   5. Kirim email HTML dengan banner tier, QR code, dan kode alfanumerik
   *   6. Tandai baris sebagai 'Terkirim' atau 'Gagal' di kolom Status Pengiriman
   *
   * Catatan penting:
   *   - Google membatasi 100 email per hari untuk akun biasa. Kalau lebih, cicil di hari berbeda.
   *   - Kalau QR code gagal dibuat (layanan down), email tetap terkirim — kode alfanumerik tetap berfungsi.
   *   - Kalau status suatu baris adalah 'Gagal', kosongin kolom Status Pengiriman-nya
   *     lalu jalankan ulang fungsi ini — baris tersebut akan diproses kembali.
   */
  function sendEmailsFromSheet() {
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const values = sheet.getDataRange().getValues();
    const header = values[0];

    // Pemetaan nama kolom ke indeks — kalau nama kolom di spreadsheet berubah,
    // perbarui string-string di bawah ini agar tetap cocok
    const col = {
      email:       header.indexOf('Email untuk pengiriman tiket'),
      nama:        header.indexOf('Nama Lengkap'),
      show:        header.indexOf('Pilih Show'),
      jenis:       header.indexOf('Jenis Tiket'),
      jumlah:      header.indexOf('Jumlah Tiket yang Dipesan'),
      total:       header.indexOf('Total Pembayaran'),
      bukti:       header.indexOf('Validasi Bukti Pembayaran'),
      statusKirim: header.indexOf('Status Pengiriman'),
      kodeAlfa:    header.indexOf('Kode Alfanumerik')
    };

    // Validasi awal: pastikan semua kolom yang dibutuhkan ada di spreadsheet
    const missing = Object.entries(col).filter(([_, v]) => v === -1).map(([k]) => k);
    if (missing.length) {
      SpreadsheetApp.getUi().alert('Kolom tidak ditemukan:\n' + missing.join('\n'));
      return;
    }

    // Kumpulkan semua kode yang sudah ada supaya kode baru tidak bentrok
    const existingCodes = new Set(
      values.slice(1)
        .map(r => String(r[col.kodeAlfa]).trim().toUpperCase())
        .filter(c => c.length === 4)
    );

    let sentCount = 0, failCount = 0;

    for (let i = 1; i < values.length; i++) {
      const row         = values[i];
      const email       = String(row[col.email]).trim();
      const nama        = String(row[col.nama]).trim();
      const show        = String(row[col.show]).trim();
      const jenis       = String(row[col.jenis]).trim().toUpperCase();
      const jumlah      = String(row[col.jumlah]).trim();
      const total       = String(row[col.total]).trim();
      const bukti       = row[col.bukti];           // true kalau checkbox sudah dicentang bendahara
      const statusKirim = String(row[col.statusKirim]).trim();

      // Lewati baris yang: emailnya kosong, pembayarannya belum diverifikasi, atau sudah terkirim
      if (!email || bukti !== true || statusKirim === 'Terkirim') continue;

      try {
        // Ambil kode yang sudah ada, atau generate baru kalau belum ada
        let kode = String(row[col.kodeAlfa]).trim().toUpperCase();
        if (kode.length !== 4) {
          kode = generateUniqueCode(existingCodes);
          existingCodes.add(kode);
          sheet.getRange(i + 1, col.kodeAlfa + 1).setValue(kode);
          SpreadsheetApp.flush(); // simpan ke spreadsheet sekarang, jangan tunggu loop selesai
        }

        // Ambil gambar banner sesuai tier
        const tierBlob = getTicketImageBlob(jenis);

        // Generate QR code dari layanan eksternal
        // Data yang di-encode: nama, email, show, jumlah, total, dan kode alfanumerik
        const qrData = `NAMA=${nama}|EMAIL=${email}|SHOW=${show}|TIKET=${jumlah}|TOTAL=${total}|KODE=${kode}`;
        const qrBlob = UrlFetchApp.fetch(
          `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(qrData)}`
        ).getBlob().setName('qrcode.png');

        // Gambar di-embed langsung ke dalam email (inline), bukan sebagai lampiran
        const inlineImages = { qrcode: qrBlob };
        if (tierBlob) inlineImages.ticketBanner = tierBlob;

        MailApp.sendEmail({
          to: email,
          subject: 'Konfirmasi Pembelian Tiket NUEVALA 2026',
          htmlBody: buildEmailHtml(nama, show, jumlah, total, jenis, kode, !!tierBlob),
          inlineImages
        });

        sheet.getRange(i + 1, col.statusKirim + 1).setValue('Terkirim');
        SpreadsheetApp.flush();
        sentCount++;
        Logger.log(`Terkirim ke ${email} | tier: ${jenis} | kode: ${kode}`);

      } catch(e) {
        sheet.getRange(i + 1, col.statusKirim + 1).setValue('Gagal');
        SpreadsheetApp.flush();
        failCount++;
        Logger.log(`Gagal untuk ${email}: ${e}`);
      }
    }

    SpreadsheetApp.getUi().alert(`Selesai.\nTerkirim: ${sentCount}\nGagal: ${failCount}`);
  }

  // ============================================================
  // SHARED VALIDATION LOGIC
  // ============================================================

  /**
   * Mengambil data lengkap dari spreadsheet aktif beserta pemetaan kolomnya.
   * Fungsi ini dipakai bersama oleh validasi QR code maupun kode alfanumerik
   * supaya tidak ada duplikasi kode yang sama di dua tempat.
   *
   * @returns {{ sheet: Sheet, values: Array[], col: Object }}
   *   - sheet:  objek sheet aktif (untuk operasi tulis)
   *   - values: semua baris data termasuk header
   *   - col:    pemetaan nama kolom ke indeks angka
   */
  function getSheetData() {
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const col = {
      email:       header.indexOf('Email untuk pengiriman tiket'),
      nama:        header.indexOf('Nama Lengkap'),
      show:        header.indexOf('Pilih Show'),
      jenis:       header.indexOf('Jenis Tiket'),
      jumlah:      header.indexOf('Jumlah Tiket yang Dipesan'),
      total:       header.indexOf('Total Pembayaran'),
      statusValid: header.indexOf('Status Validasi'),
      waktuValid:  header.indexOf('Waktu Validasi'),
      kodeAlfa:    header.indexOf('Kode Alfanumerik')
    };
    return { sheet, values, col };
  }

  /**
   * Memvalidasi satu baris tiket dan menulis hasilnya ke spreadsheet.
   *
   * Kalau tiket belum pernah divalidasi:
   *   - Kolom Status Validasi diisi 'Valid'
   *   - Kolom Waktu Validasi diisi waktu sekarang
   *   - Kolom Keterangan diisi nama panitia yang memvalidasi
   *   - Baris diberi warna hijau sebagai penanda visual
   *
   * Kalau tiket sudah pernah divalidasi sebelumnya:
   *   - Tidak ada perubahan data di spreadsheet
   *   - Fungsi mengembalikan status 'already_valid' beserta info waktu validasi sebelumnya
   *   - Ini yang muncul sebagai kotak kuning di halaman scanner
   *
   * @param {Sheet}   sheet       - Objek sheet aktif.
   * @param {Array[]} values      - Semua data baris dari spreadsheet.
   * @param {Object}  col         - Pemetaan nama kolom ke indeks.
   * @param {number}  i           - Indeks baris yang akan divalidasi (0-based dari array values).
   * @param {string}  namaPenukar - Nama panitia yang melakukan validasi, dicatat di kolom Keterangan.
   * @returns {{ status: string, message: string, ticketType: string }}
   *   status bisa berupa: 'success', 'already_valid'
   */
  function validateRow(sheet, values, col, i, namaPenukar) {
    const row    = values[i];
    const nama   = String(row[col.nama]).trim();
    const show   = String(row[col.show]).trim();
    const jumlah = String(row[col.jumlah]).trim();
    const jenis  = String(row[col.jenis]).trim().toUpperCase();

    // Cek langsung dari spreadsheet (bukan dari cache values) untuk menghindari kondisi balapan
    // kalau dua panitia scan tiket yang sama di waktu yang hampir bersamaan
    const currentStatus = sheet.getRange(i + 1, col.statusValid + 1).getValue();

    if (currentStatus === 'Valid') {
      const existingTime = sheet.getRange(i + 1, col.waktuValid + 1).getValue();
      const formatted = existingTime instanceof Date
        ? Utilities.formatDate(existingTime, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss')
        : String(existingTime);
      return {
        status: 'already_valid',
        message: `Tiket ${nama} (${show}, ${jumlah} tiket) sudah divalidasi pada ${formatted}.`,
        ticketType: jenis
      };
    }

    const now = new Date();
    const keteranganCol = sheet.getLastRow() > 0
      ? values[0].indexOf('Keterangan')
      : -1;

    sheet.getRange(i + 1, col.statusValid + 1).setValue('Valid');
    sheet.getRange(i + 1, col.waktuValid + 1).setValue(now);
    if (keteranganCol !== -1) {
      sheet.getRange(i + 1, keteranganCol + 1).setValue('Ditukar oleh: ' + namaPenukar);
    }
    // Warna hijau sebagai penanda visual baris yang sudah divalidasi
    sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground('#d4edda');
    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: `✓ ${jumlah} tiket ${show} untuk ${nama} berhasil divalidasi.\nDitukar oleh: ${namaPenukar}`,
      ticketType: jenis
    };
  }

  // ============================================================
  // QR CODE VALIDATION
  // ============================================================

  /**
   * Memproses data hasil scan QR code dari halaman scanner (Index.html).
   *
   * QR code yang di-scan akan menghasilkan string dengan format:
   *   NAMA=...|EMAIL=...|SHOW=...|TIKET=...|TOTAL=...|KODE=...
   *
   * Fungsi ini mem-parsing string tersebut, mencari baris yang cocok
   * di spreadsheet berdasarkan kombinasi email + nama + show + jumlah + total,
   * lalu memanggil validateRow() untuk menyelesaikan proses validasi.
   *
   * @param {string} qrData      - String hasil decode QR code.
   * @param {string} namaPenukar - Nama panitia yang melakukan scan.
   * @returns {{ status: string, message: string, ticketType: string }}
   *   status bisa berupa: 'success', 'already_valid', 'not_found', 'error'
   */
  function processQrCodeData(qrData, namaPenukar) {
    if (!qrData || typeof qrData !== 'string' || !qrData.trim()) {
      return { status: 'error', message: 'Data QR kosong atau tidak valid.' };
    }
    if (!namaPenukar || !namaPenukar.trim()) {
      return { status: 'error', message: 'Nama penukar harus diisi.' };
    }

    const { sheet, values, col } = getSheetData();
    const missing = Object.entries(col).filter(([_, v]) => v === -1).map(([k]) => k);
    if (missing.length) return { status: 'error', message: 'Kolom tidak ditemukan: ' + missing.join(', ') };

    // Parsing string QR code jadi objek key-value
    const parsed = {};
    qrData.split('|').forEach(item => {
      const eq = item.indexOf('=');
      if (eq !== -1) parsed[item.substring(0, eq).trim()] = item.substring(eq + 1).trim();
    });

    const { NAMA, EMAIL, SHOW, TIKET, TOTAL } = parsed;
    if (!NAMA || !EMAIL || !SHOW || !TIKET || !TOTAL) {
      return { status: 'error', message: 'Format QR tidak sesuai atau data tidak lengkap.' };
    }

    // Cari baris yang cocok berdasarkan kombinasi 5 field — lebih aman dari sekadar email saja
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (
        String(row[col.email]).trim()  === EMAIL &&
        String(row[col.nama]).trim()   === NAMA  &&
        String(row[col.show]).trim()   === SHOW  &&
        String(row[col.jumlah]).trim() === TIKET &&
        String(row[col.total]).trim()  === TOTAL
      ) {
        return validateRow(sheet, values, col, i, namaPenukar.trim());
      }
    }

    return { status: 'not_found', message: `Data tidak ditemukan: ${NAMA} | ${EMAIL} | ${SHOW}` };
  }

  /**
   * Memproses verifikasi tiket via kode alfanumerik 4 karakter (jalur manual).
   *
   * Dipakai ketika QR code tidak bisa dibaca (kamera bermasalah, printout blur, dsb).
   * Pembeli cukup tunjukkan kode 4 karakter dari email mereka, panitia input secara manual.
   *
   * Pencarian dilakukan dengan membandingkan kode yang diinput (diubah ke huruf besar)
   * dengan kolom Kode Alfanumerik di setiap baris spreadsheet.
   *
   * @param {string} code        - Kode 4 karakter yang diinput panitia.
   * @param {string} namaPenukar - Nama panitia yang melakukan verifikasi.
   * @returns {{ status: string, message: string, ticketType: string }}
   *   status bisa berupa: 'success', 'already_valid', 'not_found', 'error'
   */
  function processAlphanumericCode(code, namaPenukar) {
    if (!code || typeof code !== 'string' || !code.trim()) {
      return { status: 'error', message: 'Kode kosong atau tidak valid.' };
    }
    if (!namaPenukar || !namaPenukar.trim()) {
      return { status: 'error', message: 'Nama penukar harus diisi.' };
    }

    const upperCode = code.trim().toUpperCase();
    if (upperCode.length !== 4) {
      return { status: 'error', message: 'Kode harus tepat 4 karakter.' };
    }

    const { sheet, values, col } = getSheetData();
    const missing = Object.entries(col).filter(([_, v]) => v === -1).map(([k]) => k);
    if (missing.length) return { status: 'error', message: 'Kolom tidak ditemukan: ' + missing.join(', ') };

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][col.kodeAlfa]).trim().toUpperCase() === upperCode) {
        return validateRow(sheet, values, col, i, namaPenukar.trim());
      }
    }

    return { status: 'not_found', message: `Kode "${upperCode}" tidak ditemukan.` };
  }

  // ============================================================
  // BOILERPLATE
  // ============================================================

  /**
   * Dipanggil otomatis oleh Google Apps Script setiap kali spreadsheet dibuka.
   * Fungsi ini menambahkan menu "Pengiriman Email" ke menu bar spreadsheet
   * supaya panitia ticketing bisa mengirim email cukup dari spreadsheet,
   * tanpa perlu buka Apps Script editor sama sekali.
   */
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Pengiriman Email')
      .addItem('Kirim Email Tiket', 'sendEmailsFromSheet')
      .addToUi();
  }
