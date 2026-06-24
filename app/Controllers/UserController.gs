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
 * (Persiapan V1.1) Menambahkan EXP harian atau EXP latihan ke akun siswa
 * @param {string} kd_user - Kode unik pengguna
 * @param {number} jumlah_exp - Total EXP yang didapatkan
 * @param {string} aktivitas - Deskripsi sumber EXP (Contoh: LOGIN_HARIAN, LATIHAN)
 * @return {boolean} Status keberhasilan
 */
function tambahExpSiswa(kd_user, jumlah_exp, aktivitas) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    
    // 1. Catat riwayat perolehan EXP ke sheet 'exp_log'
    const sheetExpLog = ss.getSheetByName("exp_log");
    if(sheetExpLog) {
      let now = new Date();
      let kdExpLog = "EXP-" + now.getTime();
      // Format: [kd_exp_log, fkd_user, aktivitas, exp_didapat, keterangan, tanggal, created_at]
      sheetExpLog.appendRow([
        kdExpLog, kd_user, aktivitas, jumlah_exp, "Sistem V1.0", now, now
      ]);
    }

    // 2. Tambahkan EXP ke profil master user (sheet 'users')
    const sheetUser = ss.getSheetByName("users");
    if(sheetUser) {
      const dataUser = sheetUser.getDataRange().getValues();
      for(let i = 1; i < dataUser.length; i++) {
        if(dataUser[i][0] === kd_user) {
          
          let expLama = Number(dataUser[i][8]) || 0;
          let totalExpLama = Number(dataUser[i][9]) || 0;
          
          let expBaru = expLama + Number(jumlah_exp);
          let totalExpBaru = totalExpLama + Number(jumlah_exp);
          
          // Tulis ulang nilai EXP ke kolom I (index 9) dan J (index 10)
          sheetUser.getRange(i + 1, 9).setValue(expBaru);
          sheetUser.getRange(i + 1, 10).setValue(totalExpBaru);
          
          // Logika Kenaikan Level (Level Up) akan ditambahkan di sini pada V1.1
          break;
        }
      }
    }
    return true;
  } catch(e) {
    return false;
  }
}