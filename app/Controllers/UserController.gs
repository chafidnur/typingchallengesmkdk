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

/**
 * Mengambil nama kelas siswa berdasarkan fkd_kelas yang terikat di akunnya
 */
function getKelasInfoUser(kd_user) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetUser = ss.getSheetByName("users");
    const dataUser = sheetUser.getDataRange().getValues();
    
    let fkdKelas = "";
    for(let i = 1; i < dataUser.length; i++) {
      if(String(dataUser[i][0]).trim() === kd_user) {
        fkdKelas = String(dataUser[i][1]).trim(); // Kolom B (fkd_kelas)
        break;
      }
    }
    
    if(!fkdKelas || fkdKelas === "-" || fkdKelas === "") return "-";
    
    const sheetKelas = ss.getSheetByName("kelas");
    const dataKelas = sheetKelas.getDataRange().getValues();
    for(let j = 1; j < dataKelas.length; j++) {
      if(String(dataKelas[j][0]).trim() === fkdKelas) {
        return dataKelas[j][3]; // Kolom D: nama_kelas (Contoh: XA-FARMASI)
      }
    }
    return "-";
  } catch(e) {
    return "-";
  }
}

/**
 * Mengambil data papan peringkat teratas dari sheet 'score' untuk Hall of Fame
 */
function getHallOfFameData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetScore = ss.getSheetByName("score");
    const sheetUser = ss.getSheetByName("users");
    const sheetKelas = ss.getSheetByName("kelas");
    
    const dataScore = sheetScore.getDataRange().getValues();
    const dataUser = sheetUser.getDataRange().getValues();
    const dataKelas = sheetKelas.getDataRange().getValues();
    
    let kelasMap = {};
    for(let k = 1; k < dataKelas.length; k++) {
      kelasMap[String(dataKelas[k][0]).trim()] = dataKelas[k][3]; // Nama Kelas
    }
    
    let rekorSiswa = {};
    
    for(let i = 1; i < dataScore.length; i++) {
      let kdUser = String(dataScore[i][1]).trim();
      let wpm = parseInt(dataScore[i][3]) || 0;
      let acc = parseInt(dataScore[i][4]) || 0;
      
      // Cari data user untuk melengkapi info
      let userData = dataUser.find(u => String(u[0]).trim() === kdUser);
      
      if(userData) {
        let nama = userData[2] || "Unknown";
        let kelasID = String(userData[1]).trim();
        let gelar = userData[14] || "Trainee"; // Pastikan kolom 14/O adalah Gelar
        let namaKelas = kelasMap[kelasID] || "Kelas Tidak Ditemukan";
        
        if(!rekorSiswa[kdUser] || wpm > rekorSiswa[kdUser].wpm) {
          rekorSiswa[kdUser] = {
            nama: nama,
            username: kdUser,
            kelas: namaKelas,
            gelar: gelar,
            avatar: ambilAvatarBase64(kdUser),
            wpm: wpm,
            acc: acc
          };
        }
      }
    }
    
    return Object.values(rekorSiswa).sort((a, b) => b.wpm - a.wpm).slice(0, 10);
  } catch(e) {
    return [];
  }
}

/**
 * Mengambil info relasional untuk sapaan Dashboard: Kelas, Prodi, dan Tahun Ajaran Aktif
 */
function getInfoSapaan(kd_user) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    
    // 1. Cari Tahun Ajaran Aktif
    const sheetTA = ss.getSheetByName("tahun_ajaran");
    let infoTA = "-";
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      for (let i = 1; i < dataTA.length; i++) {
        // Cek kolom status_aktif (Indeks 3 / Kolom D)
        if (String(dataTA[i][3]).trim() === "1") { 
          infoTA = dataTA[i][1] + " Semester " + dataTA[i][2]; 
          break;
        }
      }
    }

    // 2. Cari fkd_kelas pada data User
    const sheetUser = ss.getSheetByName("users");
    const dataUser = sheetUser.getDataRange().getValues();
    let fkdKelas = "";
    for(let i = 1; i < dataUser.length; i++) {
      if(String(dataUser[i][0]).trim() === kd_user) {
        fkdKelas = String(dataUser[i][1]).trim(); // Kolom B
        break;
      }
    }

    let namaKelas = "-";
    let fkdProdi = "";
    let namaProdi = "-";

    // 3. Tarik Nama Kelas dan fkd_prodi
    if (fkdKelas && fkdKelas !== "-") {
      const sheetKelas = ss.getSheetByName("kelas");
      if (sheetKelas) {
        const dataKelas = sheetKelas.getDataRange().getValues();
        for(let j = 1; j < dataKelas.length; j++) {
          if(String(dataKelas[j][0]).trim() === fkdKelas) {
            namaKelas = dataKelas[j][3]; // Kolom D
            fkdProdi = dataKelas[j][1];  // Kolom B
            break;
          }
        }
      }

      // 4. Tarik Nama Lengkap Program Studi
      if (fkdProdi) {
        const sheetProdi = ss.getSheetByName("programstudi");
        if (sheetProdi) {
          const dataProdi = sheetProdi.getDataRange().getValues();
          for (let k = 1; k < dataProdi.length; k++) {
            if (String(dataProdi[k][0]).trim() === fkdProdi) {
              namaProdi = dataProdi[k][1]; // Kolom B
              break;
            }
          }
        }
      }
    }

    return {
      status: "success",
      kelas: namaKelas,
      prodi: namaProdi,
      tahunAjaran: infoTA
    };
  } catch(e) {
    return { status: "error", message: e.message };
  }
}