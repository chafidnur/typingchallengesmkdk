/**
 * Mengambil kalimat acak dari Bank Kalimat
 * @return {string} Teks tantangan
 */
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

/**
 * Fungsi Penyimpanan Skor & Update Statistik Terbaik Pemain
 */
function simpanSkorPemain(kd_user, wpm, accuracy, salah, total_karakter, durasi) {
  const ss = SpreadsheetApp.openById(DB_ID);
  
  // 1. Simpan Riwayat Permainan ke Sheet 'score'
  const sheetScore = ss.getSheetByName("score");
  if(sheetScore) {
    let now = new Date();
    let totalBenar = total_karakter - salah;
    let expDidapat = Math.round(wpm * 1.5); // Rumus sederhana V1.0
    let kdScore = "SCR-" + now.getTime(); // Generate ID unik
    
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

/**
 * Mengambil daftar Top 10 Pemain berdasarkan WPM Terbaik
 * @return {Array} Array of Object berisi data pemain
 */
function getLeaderboard() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return [];

  const data = sheetUser.getDataRange().getValues();
  let players = [];

  // Mulai dari baris 2 (index 1) untuk mengabaikan header
  for (let i = 1; i < data.length; i++) {
    // Hanya ambil role SISWA yang statusnya aktif (Asumsi kolom S (Index 18) = 'Y')
    if (data[i][6] === "SISWA") { 
      players.push({
        nama: data[i][5],             // Kolom F
        level: data[i][7] || 1,       // Kolom H
        wpm: data[i][11] || 0,        // Kolom L
        acc: data[i][12] || 0         // Kolom M
      });
    }
  }

  // Urutkan array secara menurun (Descending) berdasarkan WPM tertinggi
  players.sort((a, b) => b.wpm - a.wpm);

  // Potong array, kembalikan hanya 10 pemain teratas
  return players.slice(0, 10);
}

/**
 * Menarik statistik performa terbaik user untuk diolah pada komponen beranda
 */
function getUserStats(kd_user) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if(!sheetUser) return { wpm: 0, acc: 0 };

  const dataUser = sheetUser.getDataRange().getValues();
  for(let i = 1; i < dataUser.length; i++) {
    if(dataUser[i][0] === kd_user) {
      return {
        wpm: dataUser[i][11] || 0, // Mengambil total_wpm_terbaik dari Kolom L
        acc: dataUser[i][12] || 0  // Mengambil total_accuracy_terbaik dari Kolom M
      };
    }
  }
  return { wpm: 0, acc: 0 };
}

/**
 * Mengambil daftar Top 10 Pemain berdasarkan WPM Terbaik untuk Leaderboard
 * (Sudah dilengkapi dengan Gelar/Achievement)
 */
function getLeaderboard() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return [];

  const data = sheetUser.getDataRange().getValues();
  let players = [];

  for (let i = 1; i < data.length; i++) {
    // Pastikan hanya siswa yang dihitung
    if (data[i][6] === "SISWA") { 
      players.push({
        nama: data[i][5],             // Kolom F: Nama
        level: data[i][7] || 1,       // Kolom H: Level
        wpm: data[i][11] || 0,        // Kolom L: total_wpm_terbaik
        acc: data[i][12] || 0,        // Kolom M: total_accuracy_terbaik
        gelar: data[i][15] || '-'     // Kolom P: Gelar/Achievement
      });
    }
  }

  // Mengurutkan array dari yang tertinggi (Descending) berdasarkan WPM
  players.sort((a, b) => b.wpm - a.wpm);

  // Ambil 10 teratas
  return players.slice(0, 10);
}