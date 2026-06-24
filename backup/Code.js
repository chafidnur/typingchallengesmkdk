// Konfigurasi Database (Ganti dengan ID Spreadsheet Anda)
const DB_ID = "1bIKv1sFJcI4Z7ohMzyHFd3cvME-MKs8TPpnHCKtP61k";

// ====================================================
// FUNGSI UTAMA (ROUTING SISTEM SPA)
// ====================================================

// Fungsi DoGet (Entry point Web App)
function doGet(e) {
  // PERBAIKAN KRITIS: Selalu muat "Index" sebagai wadah utama.
  // Jangan langsung memuat halaman lain agar CSS dan JS tetap terbawa.
  let html = HtmlService.createTemplateFromFile("Index");
  return html.evaluate()
             .setTitle("Typing Challenge SMK Duta Karya")
             .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
             .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk Include file terpisah (Login, Utils, dll) ke dalam Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Fungsi untuk menyuplai HTML dinamis saat pindah halaman (digunakan oleh Utils.html)
function getHtmlTemplate(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================================================
// DATABASE OPERATION - VERSI 1.0
// ====================================================

// Fungsi Test Koneksi Database (Membaca Sheet 'setting')
function getAppConfig() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheet = ss.getSheetByName("setting");
  
  // Pengaman jika sheet belum ada
  if (!sheet) return {}; 
  
  const data = sheet.getDataRange().getValues();
  let config = {};
  
  // Skip baris 1 (Header), mulai dari baris 2
  for (let i = 1; i < data.length; i++) {
    config[data[i][1]] = data[i][2]; // setting_key = setting_value
  }
  return config;
}

// Fungsi Autentikasi Login Dasar (Versi 1.0)
function prosesLogin(username, password) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  
  // Validasi jika sheet 'users' belum dibuat di Spreadsheet
  if (!sheetUser) {
    return { status: "error", message: "Sistem Error: Sheet 'users' tidak ditemukan di database!" };
  }

  const dataUser = sheetUser.getDataRange().getValues();
  
  // Skip baris 1 (Header), mulai cari dari baris 2 (index 1)
  for(let i = 1; i < dataUser.length; i++) {
    
    // PERBAIKAN KRITIS: Paksa tipe data menjadi String dan hilangkan spasi kosong
    // Ini mencegah error jika di Excel password diset angka 1234, tapi JS membaca "1234"
    let dbUser = String(dataUser[i][2]).trim(); 
    let dbPass = String(dataUser[i][3]).trim(); 
    
    let inputUser = String(username).trim();
    let inputPass = String(password).trim();
    
    if(dbUser === inputUser && dbPass === inputPass) {
      // Jika cocok, kembalikan data user
      return {
        status: "success",
        kd_user: dataUser[i][0], // Kolom A: kd_user
        nama: dataUser[i][5],    // Kolom F: nama
        role: dataUser[i][6]     // Kolom G: role (ADMIN/GURU/SISWA)
      };
    }
  }
  
  // Jika loop selesai dan tidak ada yang cocok
  return { status: "error", message: "Username atau Password salah!" };
}

// ====================================================
// FUNGSI GAME & TYPING TEST
// ====================================================

function getKalimatAcak() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName("kalimat");
    
    if(sheet) {
       const data = sheet.getDataRange().getValues();
       // Pastikan ada isinya lebih dari sekadar header (baris 1)
       if(data.length > 1) {
          // Pilih baris acak (dari index 1 sampai akhir)
          let randomIndex = Math.floor(Math.random() * (data.length - 1)) + 1;
          
          // Sesuaikan dengan letak kolom 'kalimat' di Sheet Anda.
          // Misal kalau kalimat ada di kolom D (index 3):
          let teksKalimat = String(data[randomIndex][3]).trim();
          
          if(teksKalimat !== "") return teksKalimat;
       }
    }
  } catch(e) {
    // Abaikan error, langsung jalankan fallback di bawah
  }
  
  // FALLBACK: Jika Sheet 'kalimat' belum dibuat / kosong, munculkan teks ini
  return "Informatika SMK Duta Karya menyiapkan generasi digital yang kompeten dalam merakit masa depan dengan baris kode dan teknologi kecerdasan buatan.";
}

// ====================================================
// FUNGSI VALIDASI DATA MASTER
// ====================================================

/**
 * Mengecek apakah NIS sudah terdaftar di sistem.
 * Return: TRUE jika NIS belum ada (aman), FALSE jika dobel.
 */
function cekNisAman(nisBaru) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  
  if (!sheetUser) return false;

  const dataUser = sheetUser.getDataRange().getValues();
  let nisTarget = String(nisBaru).trim();

  // Asumsi urutan kolom di sheet 'users':
  // [0]kd_user, [1]fkd_kelas, [2]username, [3]password, [4]nis
  for(let i = 1; i < dataUser.length; i++) {
    let nisDb = String(dataUser[i][4]).trim(); 
    
    if(nisDb === nisTarget) {
      return false; // Gagal: NIS sudah digunakan!
    }
  }
  
  return true; // Sukses: NIS belum terdaftar
}

// Contoh fungsi saat Anda/Admin menambahkan user baru
function simpanUserBaru(dataSiswa) {
  // Panggil validasi terlebih dahulu
  if (!cekNisAman(dataSiswa.nis)) {
    return { status: "error", message: "NIS " + dataSiswa.nis + " sudah terdaftar di database!" };
  }
}

// ====================================================
// FUNGSI PENYIMPANAN SKOR & UPDATE STATISTIK
// ====================================================

function simpanSkorPemain(kd_user, wpm, accuracy, salah, total_karakter, durasi) {
  const ss = SpreadsheetApp.openById(DB_ID);
  
  // 1. Simpan Riwayat Permainan ke Sheet 'score'
  const sheetScore = ss.getSheetByName("score");
  if(sheetScore) {
    let now = new Date();
    // [kd_score, fkd_user, fkd_event, wpm, accuracy, total_benar, total_salah, durasi_detik, exp_didapat, nilai, rank_room, rank_kelas, rank_sekolah, tgl_main, created_at]
    let totalBenar = total_karakter - salah;
    let expDidapat = Math.round(wpm * 1.5); // Contoh rumus sederhana V1.0
    let kdScore = "SCR-" + now.getTime(); // Generate ID unik
    
    sheetScore.appendRow([
      kdScore, kd_user, "TRAINING", wpm, accuracy, totalBenar, salah, durasi, expDidapat, 0, "-", "-", "-", now, now
    ]);
  }

  // 2. Update WPM Terbaik dan Akurasi Terbaik di Sheet 'users'
  const sheetUser = ss.getSheetByName("users");
  if(sheetUser) {
    const dataUser = sheetUser.getDataRange().getValues();
    
    for(let i = 1; i < dataUser.length; i++) {
      if(dataUser[i][0] === kd_user) {
        let bestWpm = Number(dataUser[i][12]) || 0; // Kolom M: total_wpm_terbaik
        let bestAcc = Number(dataUser[i][13]) || 0; // Kolom N: total_accuracy_terbaik
        
        let targetWpm = Number(wpm);
        let targetAcc = Number(accuracy);

        // Hanya update jika skor saat ini lebih tinggi dari rekor terbaiknya
        if(targetWpm > bestWpm) {
          sheetUser.getRange(i + 1, 13).setValue(targetWpm);
        }
        if(targetAcc > bestAcc) {
          sheetUser.getRange(i + 1, 14).setValue(targetAcc);
        }
        
        break; // Hentikan pencarian jika user sudah ketemu dan diupdate
      }
    }
  }
  
  return true;
}

// ====================================================
// FUNGSI PENYIMPANAN SKOR & UPDATE STATISTIK
// ====================================================

function simpanSkorPemain(kd_user, wpm, accuracy, salah, total_karakter, durasi) {
  const ss = SpreadsheetApp.openById(DB_ID);
  
  // 1. Simpan Riwayat Permainan ke Sheet 'score'
  const sheetScore = ss.getSheetByName("score");
  if(sheetScore) {
    let now = new Date();
    let totalBenar = total_karakter - salah;
    let expDidapat = Math.round(wpm * 1.5); 
    let kdScore = "SCR-" + now.getTime(); 
    
    // Format: [kd_score, fkd_user, fkd_event, wpm, accuracy, total_benar, total_salah, durasi_detik, exp_didapat, nilai, rank_room, rank_kelas, rank_sekolah, tgl_main, created_at]
    sheetScore.appendRow([
      kdScore, kd_user, "TRAINING", wpm, accuracy, totalBenar, salah, durasi, expDidapat, 0, "-", "-", "-", now, now
    ]);
  }

  // 2. Update WPM Terbaik dan Akurasi Terbaik di Sheet 'users'
  const sheetUser = ss.getSheetByName("users");
  if(sheetUser) {
    const dataUser = sheetUser.getDataRange().getValues();
    
    for(let i = 1; i < dataUser.length; i++) {
      if(dataUser[i][0] === kd_user) {
        // Kolom L (Index 11) = total_wpm_terbaik
        // Kolom M (Index 12) = total_accuracy_terbaik
        let bestWpm = Number(dataUser[i][11]) || 0; 
        let bestAcc = Number(dataUser[i][12]) || 0; 
        
        let targetWpm = Number(wpm);
        let targetAcc = Number(accuracy);

        // Hanya update jika tembakan skor baru lebih tinggi dari rekor lama
        if(targetWpm > bestWpm) {
          sheetUser.getRange(i + 1, 12).setValue(targetWpm);
        }
        if(targetAcc > bestAcc) {
          sheetUser.getRange(i + 1, 13).setValue(targetAcc);
        }
        break; 
      }
    }
  }
  return true;
}

// Menarik data statistik terbaik untuk ditampilkan di Dashboard
function getUserStats(kd_user) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if(!sheetUser) return { wpm: 0, acc: 0 };

  const dataUser = sheetUser.getDataRange().getValues();
  for(let i = 1; i < dataUser.length; i++) {
    if(dataUser[i][0] === kd_user) {
      return {
        wpm: dataUser[i][11] || 0, // total_wpm_terbaik
        acc: dataUser[i][12] || 0  // total_accuracy_terbaik
      };
    }
  }
  return { wpm: 0, acc: 0 };
}