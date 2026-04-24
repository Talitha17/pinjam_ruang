function doGet(e) {
  // 1. Menangkap parameter "page" dari URL (misal: ?page=peminjam)
  // Jika tidak ada parameter page, maka default-nya kosong / buka index
  var page = '';
  if (e && e.parameter && e.parameter.page) {
    page = e.parameter.page;
  }
  
  // ==========================================
  // ROUTING HALAMAN
  // ==========================================
  
  // A. Jika URL memuat parameter ?page=peminjam
  if (page === 'peminjam') {
    const template = HtmlService.createTemplateFromFile('peminjam'); // Nama file Peminjam.html
    
    // Menangkap filter tanggal jika ada
    var startFilter = '';
    if (e.parameter.start) {
      startFilter = e.parameter.start;
    }
    
    template.title = "Daftar Peminjam";
    template.pageActive = page; // <- Kirim variabel halaman aktif
    template.webAppUrl = ScriptApp.getService().getUrl();
    template.startParam = startFilter; 
    
    // PANGGIL FUNGSI: Dapatkan data peminjam dari sheet (sesuaikan dengan nama fungsi Anda)
    // Pastikan getPeminjam() mengembalikan Array of Objects
    template.peminjam = getPeminjam(startFilter);
    
    return template.evaluate()
        .setTitle('Daftar Peminjam')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // B. (Opsional) Jika URL memuat parameter ?page=cek (Untuk halaman form meminjam)
    else if (page === 'cek') {
    // 1. Buatkan file HTML baru di menu Apps Script bernama "FormCek.html"
    const template = HtmlService.createTemplateFromFile('cek'); 
    
    // 2. Tangkap ID ruangan yang di-klik user
    var idRuangan = e.parameter.id;
    
    
    template.title = "Cek & Pinjam Ruangan";
    template.pageActive = page; // <- Kirim variabel halaman aktif
    template.id_ruangan = idRuangan; // Kirim id ke HTML agar bisa ditampilkan atau diproses
    template.webAppUrl = ScriptApp.getService().getUrl();
  

    // MISAL: Mengirimkan daftar Jam Operasional ke file slot_jam.html
    template.jam_operasional = ["08:00", "09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00"];
    
    // Hari ini sebagai default (YYYY-MM-DD). Bisa disesuaikan kalau dari input form / URL
    template.tanggal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Array riwayat peminjaman untuk di-filter (Misal ambil dari fungsi Anda sendiri)
    template.dipinjam = []; 

    return template.evaluate()
        .setTitle('Cek Peminjaman')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // C. Jika URL tidak memiliki parameter page (HALAMAN BERANDA / DEFAULT)
  else {
    const template = HtmlService.createTemplateFromFile('index'); // Nama file index.html
    
    template.title = "Daftar Ruangan Tersedia";
    template.pageActive = ''; // <- Kosong berarti beranda (Home)
    template.webAppUrl = ScriptApp.getService().getUrl();
    
    // PANGGIL FUNGSI: Dapatkan data ruangan (yang kodenya kita bahas di awal)
    template.ruangan = getRuangan(); 
    
    return template.evaluate()
        .setTitle('Peminjaman Ruangan')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getRuangan() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Ruangan');
  const range = sh.getDataRange();

  const values = range.getValues();
  const rich = range.getRichTextValues(); // ← WAJIB

  const header = values[0];
  const rows = values.slice(1);

  const idxRuang = header.indexOf('id_ruangan');
  const idxNama = header.indexOf('nama_ruangan');
  const idxJenis = header.indexOf('jenis_ruangan');
  const idxFoto = header.indexOf('foto');

  const hasil = [];

  // Looping untuk setiap baris data
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    
    // Karena array 'rows' dimulai dari data (tanpa header),
    // maka indeks pasangannya di array 'rich' (yang masih memiliki header) adalah i + 1
    const richRowIndex = i + 1;
    
    // Ambil sel RichText untuk foto
    const richCell = rich[richRowIndex][idxFoto];
    
    // AMBIL LINK DARI HYPERLINK (tambahkan pengecekan null safety)
    const richLink = richCell ? richCell.getLinkUrl() : null;

    hasil.push({
      id_ruangan: r[idxRuang],
      nama: r[idxNama],
      jenis_ruangan: r[idxJenis],
      foto: convertDrive(richLink)
    });
  }

  return hasil;
}

function convertDrive(url) {
  if (!url) return '';
  const id = url.match(/[-\w]{25,}/);
  if (!id) return '';
  return "https://lh3.googleusercontent.com/d/" + id[0];
}


/* =========================================================
   Fungsi Menerima Data Peminjam dengan Fitur Filter PER MINGGU
   ========================================================= */
function getPeminjam(filterTanggal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ambil Nama Ruangan
  const shRuangan = ss.getSheetByName('Ruangan');
  const dataRuangan = shRuangan.getDataRange().getValues();
  const headerRuangan = dataRuangan[0];
  const idxRefIdRuangan = headerRuangan.indexOf('id_ruangan');
  const idxRefNamaRuangan = headerRuangan.indexOf('nama_ruangan');

  const mapRuangan = {};
  for (let i = 1; i < dataRuangan.length; i++) {
    const barisRuangan = dataRuangan[i];
    const id = barisRuangan[idxRefIdRuangan];
    const nama = barisRuangan[idxRefNamaRuangan];
    if (id) {
      mapRuangan[id] = nama; 
    }
  }

  // 2. Ambil Data Peminjam
  const sh = ss.getSheetByName('Peminjam');
  const data = sh.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const idxId = header.indexOf('id_ruangan'); 
  const idxNama = header.indexOf('nama');
  const idxNip = header.indexOf('nip');
  const idxOper = header.indexOf('operator');
  const idxTel = header.indexOf('telepon');
  const idxRuang = header.indexOf('id_ruangan');
  const idxPinjam = header.indexOf('tanggal_pinjam');
  const idxKembali = header.indexOf('tanggal_kembali');
  const idxKet = header.indexOf('keterangan');
  const idxStatus = header.indexOf('status');

  const hasil = [];
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    
    // ========================================================
    // LOGIKA FILTER TANGGAL: RENTANG WAKTU 1 MINGGU (7 HARI)
    // ========================================================
    if (filterTanggal && filterTanggal !== '') {
      const pinjamValue = r[idxPinjam];
      
      // Jika kosong, abaikan
      if (!pinjamValue) continue; 
      
      const tglSheet = new Date(pinjamValue);
      if (isNaN(tglSheet.getTime())) continue; 
      
      // Standarisasi tgl dari Spreadsheet ke jam 00:00 (Awal Hari)
      const waktuDataSheet = new Date(tglSheet.getFullYear(), tglSheet.getMonth(), tglSheet.getDate()).getTime();

      // Cerna inputan Kalender HTML (contoh: "2026-03-09")
      const bagianTanggal = filterTanggal.split('-');
      const tahunInput = parseInt(bagianTanggal[0], 10);
      const bulanInput = parseInt(bagianTanggal[1], 10) - 1; // Bulan di Javascript dimulai dari 0
      const hariInput = parseInt(bagianTanggal[2], 10);
      
      // TETAPKAN BATAS BAWAH (Hari H yang dipilih)
      const batasBawahMutlak = new Date(tahunInput, bulanInput, hariInput).getTime();
      
      // TETAPKAN BATAS ATAS (Hari ke-7 / Maju 6 hari dari hari pertama)
      // Rumus: 6 Hari * 24 Jam * 60 Menit * 60 Detik * 1000 Milidetik
      const intervalWaktu = 6 * 24 * 60 * 60 * 1000;
      const batasAtasMutlak = batasBawahMutlak + intervalWaktu;
      
      // SELEKSI: Jika jadwal di sheet ini terjadi SEBELUM hari Senin pilihan, ATAU SETELAH hari Minggu (Batas Akhir)...
      if (waktuDataSheet < batasBawahMutlak || waktuDataSheet > batasAtasMutlak) {
        continue; // LEWATI! (Jangan masukkan datanya ke tampilan HTML)
      }
    }
    // ========================================================

    const idRuanganPinjam = r[idxRuang];
    const namaRuanganDitemukan = mapRuangan[idRuanganPinjam] || "Ruangan Tidak Dikenali (" + idRuanganPinjam + ")";

    hasil.push({
      id_ruangan: r[idxId],
      nama: r[idxNama],
      nip: r[idxNip],
      operator: r[idxOper],
      telepon: r[idxTel],
      id_ruangan: idRuanganPinjam,
      nama_ruangan: namaRuanganDitemukan, 
      tanggal_pinjam: r[idxPinjam],
      tanggal_kembali: r[idxKembali],
      keterangan: r[idxKet],
      status: r[idxStatus]
    });
  }

  return hasil; 
}



/* ==================================================
   Fungsi Menerima Form Submit dari Halaman HTML 
   ================================================== */
function simpanPeminjam(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPeminjam = ss.getSheetByName('Peminjam');
  
  // 1. MEMBUAT ID PEMINJAM OTOMATIS (AUTO-INCREMENT)
  const barisTerakhir = shPeminjam.getLastRow();
  let idBaru = 1; 
  
  if (barisTerakhir > 1) { 
    const idTerakhir = shPeminjam.getRange(barisTerakhir, 1).getValue();
    if (!isNaN(idTerakhir) && idTerakhir !== "") {
      idBaru = Number(idTerakhir) + 1;
    } else {
      idBaru = barisTerakhir;
    }
  }
  
  // 2. MENGATUR FORMAT WAKTU PINJAM & KEMBALI DARI INPUT MANUAL
  // Menggabungkan tanggal dan jam menjadi format utuh
  const TglPinjamGabungan = formObject.tgl_pinjam + " " + formObject.jam_pinjam;
  const TglKembaliGabungan = formObject.tgl_kembali + " " + formObject.jam_kembali;
  
  // 3. MENYUSUN DATA AGAR PAS DENGAN KOLOM SPREADSHEET (A SAMPAI J)
  const barisBaru = [
    idBaru,                   // Kolom A (Ke-1) : id_peminjam 
    formObject.nama,          // Kolom B (Ke-2) : nama
    formObject.nip,           // Kolom C (Ke-3) : nip
    formObject.operator,      // Kolom D (Ke-4) : operator
    formObject.telepon,       // Kolom E (Ke-5) : telepon
    formObject.id_ruangan,    // Kolom F (Ke-6) : id_ruangan
    TglPinjamGabungan,        // Kolom G (Ke-7) : tanggal_pinjam (Sekarang dari input manual)
    TglKembaliGabungan,       // Kolom H (Ke-8) : tanggal_kembali
    formObject.keterangan,    // Kolom I (Ke-9) : keterangan
    "Menunggu"                // Kolom J (Ke-10): status
  ];
  
  // 4. MENYIMPAN KE SPREADSHEET
  shPeminjam.appendRow(barisBaru);
  
  return true; 
}

/* ==================================================
   Fungsi untuk Tombol Gambar Verifikasi di Spreadsheet
   ================================================== */
function verifikasiPeminjaman() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. Pastikan admin sedang berada di Sheet "Peminjam"
  if (sheet.getName() !== 'Peminjam') {
    SpreadsheetApp.getUi().alert("Peringatan", "Harap buka Sheet 'Peminjam' terlebih dahulu untuk melakukan verifikasi.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 2. Ambil Nomor Baris (Row) yang sedang dipilih / disorot admin
  const barisAktif = sheet.getActiveCell().getRow();
  
  // 3. Jangan izinkan verifikasi pada baris Header (Baris 1)
  if (barisAktif === 1) {
    SpreadsheetApp.getUi().alert("Awas", "Anda memilih baris Header. Silakan klik/pilih salah satu sel di baris data anggota yang ingin diverifikasi terlebih dahulu.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 4. Pastikan baris yang dipilih tidak kosong (Bisa cek dari kolom F / ID Ruangan)
  const idRuangan = sheet.getRange(barisAktif, 6).getValue();
  if (!idRuangan) {
    SpreadsheetApp.getUi().alert("Kosong", "Baris yang Anda pilih kosong atau tidak memiliki ID Ruangan.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 5. Eksekusi Pengubahan Status 
  // Berdasarkan Screenshot Anda, kolom "status" ada di kolom J (Kolom ke-10)
  const kolomStatus = 10; 
  
  // Mengubah teks di sel status menjadi 'terverifikasi'
  sheet.getRange(barisAktif, kolomStatus).setValue('terverifikasi');
  
  // Tampilkan pesan sukses
  SpreadsheetApp.getUi().alert("Sukses!", "Peminjaman pada baris ke-" + barisAktif + " berhasil MENDAPATKAN STATUS TERVERIFIKASI.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ==================================================
   Fungsi Wajib untuk Menggabungkan / Memanggil File HTML Terpisah
   ================================================== */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ==================================================
   FITUR PINTAR: Mencari Jam yang Sudah Dibooking
   ================================================== */
function getJamTerpakai(idRuanganCari, tanggalDicari) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Peminjam');
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1);
  
  const jamSudahLaku = [];
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const idR = String(r[5]); // id_ruangan ada di Kolom F (indeks ke-5)
    const tglP = r[6];        // tanggal_pinjam ada di Kolom G (indeks ke-6)
    const status = String(r[9]).toLowerCase(); // status di Kolom J
    
    // Jangan menghitung bookingan yang batal, ditolak, atau sudah selesai
    if (status === 'batal' || status === 'ditolak' || status === 'selesai') continue;
    
    // Jika ID ruangannya cocok dan ada nilai isian tanggalnya
    if (idR === String(idRuanganCari) && tglP) {
      
      // Amankan pembacaan tanggal (Dari string text jadi Objek Waktu)
      let objTgl;
      if (typeof tglP === 'string') {
         objTgl = new Date(tglP.replace(/-/g, "/"));
      } else {
         objTgl = new Date(tglP);
      }
      
      if (isNaN(objTgl.getTime())) continue; // Lewati kalau rusak
      
      const ptTahun = objTgl.getFullYear();
      const ptBulan = ("0" + (objTgl.getMonth() + 1)).slice(-2);
      const ptHari = ("0" + objTgl.getDate()).slice(-2);
      const strTanggal = ptTahun + "-" + ptBulan + "-" + ptHari;
      
      // Jika ternyata tangalnya SAMA dengan kalender yang dipencet pengguna
      if (strTanggal === tanggalDicari) {
         // Curi jamnya (contoh: "08:00") dan kumpulkan
         const textJam = ("0" + objTgl.getHours()).slice(-2) + ":" + ("0" + objTgl.getMinutes()).slice(-2);
         jamSudahLaku.push(textJam);
      }
    }
  }
  
  return jamSudahLaku; // Kirim daftar jam mati ke HTML
}



