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
        gelar: data[i][14] || '-'     // Kolom P: Gelar/Achievement
      });
    }
  }

  // Mengurutkan array dari yang tertinggi (Descending) berdasarkan WPM
  players.sort((a, b) => b.wpm - a.wpm);

  // Ambil 10 teratas
  return players.slice(0, 10);
}

/**
 * Memvalidasi dan Memberikan Achievement kepada User
 * Fungsi ini dipanggil secara otomatis oleh sistem setelah skor dan level user dihitung.
 */
function validasiAchievement(kd_user, levelSekarang) {
  try {
    let namaGelar = "";
    let deskripsiGelar = "";

    // 1. Tentukan Hierarki Gelar Berdasarkan Level Saat Ini
    if (levelSekarang >= 100) {
      namaGelar = "👁️ Eagle Eye";
      deskripsiGelar = "Mencapai Level 100. Mata dan jarimu telah menyatu! Kecepatan seperti elang yang sudah mengincar mangsa, membaca dan mengeksekusi teks tanpa satu pun typo.";
    } else if (levelSekarang >= 80) {
      namaGelar = "💨 Ghost Scripter";
      deskripsiGelar = "Mencapai Level 80. Kamu mengetik secepat kilat, menembus batas ruang dan waktu arena!";
    } else if (levelSekarang >= 60) {
      namaGelar = "⚡ Si Jari Kilat";
      deskripsiGelar = "Mencapai Level 60. Pergerakan tanganmu tak lagi terlihat, julukan Si Jari Kilat resmi menjadi milikmu.";
    } else if (levelSekarang >= 40) {
      namaGelar = "🌩️ Thunder Strike";
      deskripsiGelar = "Mencapai Level 40. Suara ketikan mechanical keyboard-mu mulai menggelegar bagaikan badai!";
    } else if (levelSekarang >= 20) {
      namaGelar = "🏃‍♂️ Type Speed Newbie";
      deskripsiGelar = "Mencapai Level 20. Pemanasan selesai! Manuver jarimu mulai lincah menyusuri papan ketik.";
    } else {
      return { status: "no_achievement" }; // Belum memenuhi syarat minimal Level 20
    }

    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetAch = ss.getSheetByName("achievement");
    if (!sheetAch) return { status: "error", message: "Sheet achievement tidak ditemukan." };
    
    const dataAch = sheetAch.getDataRange().getValues();

    // 2. VALIDASI: Mencegah Duplikasi Gelar
    // Cek apakah fkd_user sudah pernah mendapatkan nama_achievement ini
    for (let i = 1; i < dataAch.length; i++) {
      if (String(dataAch[i][1]).trim() === kd_user && String(dataAch[i][2]).trim() === namaGelar) {
        return { status: "already_unlocked" }; 
      }
    }

    // 3. REKAM KE SHEET ACHIEVEMENT (Sesuai urutan 9 kolom Yang Mulia)
    let now = new Date();
    let kdAch = "ACH-" + now.getTime();
    
    sheetAch.appendRow([
      kdAch,            // 0: kd_achievement
      kd_user,          // 1: fkd_user
      namaGelar,        // 2: nama_achievement
      deskripsiGelar,   // 3: deskripsi
      now,              // 4: tanggal_dapat
      now,              // 5: created_at
      "SYSTEM",         // 6: created_by
      now,              // 7: update_at
      "SYSTEM"          // 8: update_by
    ]);

    // 4. UPDATE GELAR KE PROFIL USER DI SHEET 'users' (Kolom O / Indeks 14)
    const sheetUser = ss.getSheetByName("users");
    const dataUser = sheetUser.getDataRange().getValues();
    for (let i = 1; i < dataUser.length; i++) {
      if (String(dataUser[i][0]).trim() === kd_user) {
        sheetUser.getRange(i + 1, 15).setValue(namaGelar); 
        break;
      }
    }

    // Kembalikan notifikasi sukses agar UI (Frontend) bisa memunculkan Pop-Up Selamat!
    return { 
      status: "new_achievement", 
      gelar: namaGelar, 
      pesan: "PENCAPAIAN TERBUKA: " + namaGelar 
    };

  } catch (e) {
    return { status: "error", message: e.message };
  }
}

/**
 * Memproses hasil akhir setelah siswa selesai mengetik 60 detik
 * Terintegrasi langsung dengan sheet 'score' dan 'exp_log'
 */
function submitGameResult(paketData) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetScore = ss.getSheetByName("score");
    const sheetExpLog = ss.getSheetByName("exp_log");
    const sheetUser = ss.getSheetByName("users");
    
    let wpm = parseInt(paketData.wpm);
    let acc = parseInt(paketData.akurasi);
    let totalBenar = parseInt(paketData.total_benar);
    let totalSalah = parseInt(paketData.total_salah);
    let kdUser = paketData.kd_user;
    
    // 1. CEK BATAS STAMINA DARI SHEET SCORE
    let todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    let dataScore = sheetScore.getDataRange().getValues();
    let playCountToday = 0;
    
    for (let i = 1; i < dataScore.length; i++) {
      if (String(dataScore[i][1]).trim() === kdUser) {
        let tglMain = new Date(dataScore[i][13]); // Kolom tanggal_main di index 13
        let tglStr = Utilities.formatDate(tglMain, "Asia/Jakarta", "yyyy-MM-dd");
        if (tglStr === todayStr) {
          playCountToday++;
        }
      }
    }

    // 2. HITUNG EXP
    let expDidapat = 0;
    let multiplierKesulitan = 15; // Asumsi rata-rata latihan
    
    if (playCountToday < 10) {
      expDidapat = (wpm * multiplierKesulitan);
      if (acc === 100) expDidapat += 500;
      else if (acc >= 95) expDidapat += 200;
    }

    let now = new Date();
    
    // 3. REKAM KE SHEET SCORE (Sesuai urutan 17 kolom Yang Mulia)
    let kdScore = "SCR-" + now.getTime();
    sheetScore.appendRow([
      kdScore,       // 0: kd_score
      kdUser,        // 1: fkd_user
      "TRAINING",    // 2: fkd_event (Kita isi TRAINING sementara belum ada event)
      wpm,           // 3: wpm
      acc,           // 4: accuracy
      totalBenar,    // 5: total_benar
      totalSalah,    // 6: total_salah
      60,            // 7: durasi_detik
      expDidapat,    // 8: exp_didapat
      wpm,           // 9: nilai (Kita ambil dari WPM sebagai nilai mentah)
      0,             // 10: ranking_room
      0,             // 11: ranking_kelas
      0,             // 12: ranking_sekolah
      now,           // 13: tanggal_main
      now,           // 14: created_at
      "SYSTEM",      // 15: created_by
      "SYSTEM"       // 16: update_by
    ]);

    // 4. REKAM KE SHEET EXP_LOG (Jika mendapat EXP)
    if (expDidapat > 0) {
      let kdExpLog = "EXP-" + now.getTime();
      sheetExpLog.appendRow([
        kdExpLog,                                   // 0: kd_exp_log
        kdUser,                                     // 1: fkd_user
        "Training Mode",                            // 2: aktivitas
        expDidapat,                                 // 3: exp_didapat
        `Selesai latihan dengan ${wpm} WPM`,        // 4: keterangan
        now,                                        // 5: tanggal
        now,                                        // 6: created_at
        "SYSTEM",                                   // 7: created_by
        now,                                        // 8: update_at
        "SYSTEM"                                    // 9: update_by
      ]);
    }

    // 5. UPDATE TOTAL EXP & LEVEL DI SHEET USER
    let dataUser = sheetUser.getDataRange().getValues();
    let totalExpBaru = 0;
    let levelBaru = 1;
    let gelarBaru = null;

    for (let i = 1; i < dataUser.length; i++) {
      if (String(dataUser[i][0]).trim() === kdUser) {
        let expLama = parseInt(dataUser[i][12]) || 0; // Sesuaikan indeks kolom total_exp
        totalExpBaru = expLama + expDidapat;
        levelBaru = Math.floor(totalExpBaru / 1000) + 1;
        
        sheetUser.getRange(i + 1, 13).setValue(totalExpBaru);
        sheetUser.getRange(i + 1, 14).setValue(levelBaru);
        
        // Panggil Validasi Achievement (Mangekyou Sharingan, dll)
        let cekGelar = validasiAchievement(kdUser, levelBaru);
        if(cekGelar.status === "new_achievement") {
          gelarBaru = cekGelar.gelar;
        }
        break;
      }
    }

    return { 
      status: "success", 
      exp: expDidapat,
      level: levelBaru,
      sisaStamina: 10 - (playCountToday + 1),
      achievement: gelarBaru
    };

  } catch (e) {
    return { status: "error", message: e.message };
  }
}