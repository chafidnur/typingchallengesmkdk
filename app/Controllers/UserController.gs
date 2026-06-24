/**
 * Mengambil profil lengkap siswa untuk ditampilkan di antarmuka Dashboard
 * @param {string} kd_user - Kode unik pengguna (Primary Key)
 * @return {Object|null} Data profil pengguna atau null jika tidak ditemukan
 */
function getProfilSiswa(kd_user) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  
  if(!sheetUser) return null;
  
  const dataUser = sheetUser.getDataRange().getValues();
  
  // Mencari data user berdasarkan kd_user (baris demi baris, melewati header)
  for(let i = 1; i < dataUser.length; i++) {
    if(dataUser[i][0] === kd_user) {
      return {
        kd_user: dataUser[i][0],      // Kolom A
        nama: dataUser[i][5],         // Kolom F
        role: dataUser[i][6],         // Kolom G
        level: dataUser[i][7] || 1,   // Kolom H: Level saat ini
        exp: dataUser[i][8] || 0,     // Kolom I: EXP berjalan
        total_exp: dataUser[i][9]|| 0,// Kolom J: Akumulasi EXP
        gelar: dataUser[i][15] || '-' // Kolom P: Gelar/Achievement yang dipakai
      };
    }
  }
  return null;
}

/**
 * Menambahkan EXP harian atau latihan, serta menghitung Level Up
 * @param {string} kd_user - Kode unik pengguna
 * @param {number} jumlah_exp - Total EXP yang didapatkan
 * @param {string} aktivitas - Deskripsi sumber EXP 
 * @return {Object} Status dan info level up
 */
function tambahExpSiswa(kd_user, jumlah_exp, aktivitas) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    
    // 1. Catat riwayat perolehan EXP
    const sheetExpLog = ss.getSheetByName("exp_log");
    if(sheetExpLog) {
      let now = new Date();
      let kdExpLog = "EXP-" + now.getTime();
      sheetExpLog.appendRow([
        kdExpLog, kd_user, aktivitas, jumlah_exp, "Sistem V1.1", now, now
      ]);
    }

    // 2. Kalkulasi EXP dan Level di profil Master User
    const sheetUser = ss.getSheetByName("users");
    if(sheetUser) {
      const dataUser = sheetUser.getDataRange().getValues();
      for(let i = 1; i < dataUser.length; i++) {
        if(dataUser[i][0] === kd_user) {
          
          let levelLama = Number(dataUser[i][7]) || 1; // Kolom H (Index 7)
          let expLama = Number(dataUser[i][8]) || 0;   // Kolom I (Index 8)
          let totalExpLama = Number(dataUser[i][9]) || 0; // Kolom J (Index 9)
          
          let expBaru = expLama + Number(jumlah_exp);
          let totalExpBaru = totalExpLama + Number(jumlah_exp);
          let levelBaru = levelLama;
          let isLevelUp = false;

          // Rumus sederhana: Butuh (Level x 100) EXP untuk naik ke level berikutnya
          let expDibutuhkan = levelLama * 100;

          if (expBaru >= expDibutuhkan) {
            levelBaru++; // Naik level!
            expBaru = expBaru - expDibutuhkan; // Sisa EXP dipindah ke level baru
            isLevelUp = true;
          }
          
          // Simpan pembaruan ke Spreadsheet
          sheetUser.getRange(i + 1, 8).setValue(levelBaru);
          sheetUser.getRange(i + 1, 9).setValue(expBaru);
          sheetUser.getRange(i + 1, 10).setValue(totalExpBaru);
          
          return { status: true, isLevelUp: isLevelUp, levelBaru: levelBaru };
        }
      }
    }
    return { status: false };
  } catch(e) {
    return { status: false, message: e.message };
  }
}