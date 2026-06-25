function getProfilSiswa(kd_user) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  
  if(!sheetUser) return null;
  
  const dataUser = sheetUser.getDataRange().getValues();
  
  for(let i = 1; i < dataUser.length; i++) {
    if(dataUser[i][0] === kd_user) {
      return {
        kd_user: dataUser[i][0],      
        nama: dataUser[i][5],         
        role: dataUser[i][6],         
        level: dataUser[i][7] || 1,   
        exp: dataUser[i][8] || 0,     
        total_exp: dataUser[i][9]|| 0,
        gelar: dataUser[i][14] || '-' // PERBAIKAN: Indeks 14 untuk Gelar
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

/**
 * Memproses unggahan file foto dari frontend ke folder Google Drive Admin
 * Nama file dipaksa menggunakan username pengguna untuk sinkronisasi otomatis
 */
function simpanAvatarUser(base64String, mimeType, username) {
  try {
    const folder = DriveApp.getFolderById(AVATAR_FOLDER_ID);
    const namaFileTarget = String(username).trim();

    // Pembersihan: Cari dan hapus foto lama dengan nama yang sama agar tidak menumpuk di Drive
    const filesLama = folder.getFilesByName(namaFileTarget);
    while (filesLama.hasNext()) {
      filesLama.next().setTrashed(true);
    }

    // Ekstraksi data mentah Base64 dan konversi menjadi Blob
    const rawBase64 = base64String.split(",")[1];
    const decodedBytes = Utilities.base64Decode(rawBase64);
    const blobFile = Utilities.newBlob(decodedBytes, mimeType, namaFileTarget);

    // Rekam berkas baru ke dalam folder Drive
    folder.createFile(blobFile);

    return { status: "success", message: "Foto profil berhasil diperbarui!" };
  } catch (e) {
    return { status: "error", message: "Gagal mengunggah ke Drive: " + e.message };
  }
}

/**
 * Menarik file gambar dari Drive berdasarkan username dan mengubahnya menjadi string Base64
 */
function ambilAvatarBase64(username) {
  try {
    const folder = DriveApp.getFolderById(AVATAR_FOLDER_ID);
    const files = folder.getFilesByName(String(username).trim());
    
    if (files.hasNext()) {
      const fileGambar = files.next();
      const base64Data = Utilities.base64Encode(fileGambar.getBlob().getBytes());
      return "data:" + fileGambar.getMimeType() + ";base64," + base64Data;
    }
    return ""; // Kembalikan string kosong jika siswa belum mengunggah foto custom
  } catch (e) {
    return "";
  }
}