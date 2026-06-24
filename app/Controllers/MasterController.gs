/**
 * Memvalidasi apakah Nomor Induk Siswa (NIS) aman untuk didaftarkan
 */
function cekNisAman(nisBaru) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return false;

  const dataUser = sheetUser.getDataRange().getValues();
  let nisTarget = String(nisBaru).trim();

  for(let i = 1; i < dataUser.length; i++) {
    let nisDb = String(dataUser[i][4]).trim(); 
    if(nisDb === nisTarget && nisTarget !== "" && nisTarget !== "-") {
      return false; // NIS duplikat
    }
  }
  return true; 
}

/**
 * Memvalidasi apakah Kode User (kd_user) manual sudah dipakai
 */
function cekKdUserAman(kdUserBaru) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return false;

  const dataUser = sheetUser.getDataRange().getValues();
  let kdTarget = String(kdUserBaru).trim();

  for(let i = 1; i < dataUser.length; i++) {
    if(String(dataUser[i][0]).trim() === kdTarget) {
      return false; // ID duplikat
    }
  }
  return true;
}

/**
 * Menyimpan data pengguna baru (Dengan Input kd_user Manual)
 */
function simpanUserBaru(dataSiswa) {
  // Validasi NIS untuk Siswa
  if (dataSiswa.role === "SISWA" && !cekNisAman(dataSiswa.nis)) {
    return { status: "error", message: "NIS " + dataSiswa.nis + " sudah terdaftar di database!" };
  }
  
  // Validasi Kode User Manual
  if (!cekKdUserAman(dataSiswa.kd_user)) {
    return { status: "error", message: "Kode User " + dataSiswa.kd_user + " sudah dipakai. Gunakan kode lain!" };
  }
  
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetUser = ss.getSheetByName("users");
    let now = new Date();
    
    // Simpan ke 22 Kolom
    sheetUser.appendRow([
      dataSiswa.kd_user,                 // 0: kd_user (MANUAL)
      dataSiswa.fkd_kelas || "-",        // 1: fkd_kelas
      dataSiswa.username,                // 2: username
      dataSiswa.password,                // 3: password
      dataSiswa.nis || "-",              // 4: nis
      dataSiswa.nama,                    // 5: nama
      dataSiswa.role || "SISWA",         // 6: role
      1,                                 // 7: level (Default 1)
      0,                                 // 8: exp (Default 0)
      0,                                 // 9: total_exp
      0,                                 // 10: total_menang
      0,                                 // 11: total_wpm_terbaik
      0,                                 // 12: total_accuracy_terbaik
      0,                                 // 13: streak_login
      "-",                               // 14: gelar
      "",                                // 15: login_terakhir 
      dataSiswa.tahun_lulus || "-",      // 16: tahun_lulus
      dataSiswa.status_aktif || "Y",     // 17: status_aktif
      now,                               // 18: created_at
      dataSiswa.created_by || "ADMIN",   // 19: created_by
      now,                               // 20: update_at
      dataSiswa.created_by || "ADMIN"    // 21: update_by
    ]);
    return { status: "success", message: "Data pengguna berhasil disimpan!" };
  } catch(e) {
    return { status: "error", message: "Gagal menyimpan: " + e.message };
  }
}

/**
 * Mengambil list pengguna untuk tabel
 */
function getListSiswa() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return [];

  const data = sheetUser.getDataRange().getValues();
  let listSiswa = [];

  for (let i = 1; i < data.length; i++) {
    listSiswa.push({
      kd_user: data[i][0],          
      kelas: data[i][1] || '-',     
      username: data[i][2],         
      password: data[i][3],         
      nis: data[i][4],              
      nama: data[i][5],             
      role: data[i][6],             
      level: data[i][7] || 1,       
      tahun_lulus: data[i][16] || '-',
      status: data[i][17] || 'Y'    
    });
  }
  
  listSiswa.sort((a, b) => a.role.localeCompare(b.role) || a.nama.localeCompare(b.nama));
  return listSiswa;
}

/**
 * Mengambil detail profil dan riwayat latihan siswa dari Sheet Score
 */
function getRiwayatSiswa(kd_user) {
  const ss = SpreadsheetApp.openById(DB_ID);
  
  // 1. Tarik Data Profil
  const sheetUser = ss.getSheetByName("users");
  let detailUser = null;
  if(sheetUser) {
    const dataUser = sheetUser.getDataRange().getValues();
    for(let i = 1; i < dataUser.length; i++) {
      if(dataUser[i][0] === kd_user) {
        detailUser = {
          nama: dataUser[i][5],
          nis: dataUser[i][4],
          kelas: dataUser[i][1],
          level: dataUser[i][7],
          wpm_terbaik: dataUser[i][11],
          acc_terbaik: dataUser[i][12],
          gelar: dataUser[i][14]
        };
        break;
      }
    }
  }

  // 2. Tarik Riwayat Main (Score)
  const sheetScore = ss.getSheetByName("score");
  let riwayat = [];
  if(sheetScore) {
    const dataScore = sheetScore.getDataRange().getValues();
    for(let i = 1; i < dataScore.length; i++) {
      if(dataScore[i][1] === kd_user) { // fkd_user ada di indeks 1
        riwayat.push({
          wpm: dataScore[i][3],       // wpm
          acc: dataScore[i][4],       // accuracy
          salah: dataScore[i][6],     // total_salah
          durasi: dataScore[i][7],    // durasi_detik
          tanggal: dataScore[i][13]   // tanggal_main
        });
      }
    }
  }

  // Urutkan riwayat dari yang paling baru
  riwayat.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

  return {
    detail: detailUser,
    history: riwayat,
    total_main: riwayat.length
  };
}