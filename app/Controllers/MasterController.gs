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
      dataSiswa.status_aktif !== undefined ? dataSiswa.status_aktif : 1, // 17: status_aktif (1 atau 0)
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
 * Mengambil detail profil dan riwayat latihan siswa dari Sheet 'score'
 * Dibentengi dengan Try-Catch dan pemaksaan tipe String agar tidak error.
 */
function getRiwayatSiswa(kd_user) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    // Paksa menjadi string dan bersihkan spasi agar perbandingan 100% akurat
    const targetKd = String(kd_user).trim(); 

    // ============================================
    // 1. TARIK DATA PROFIL DARI SHEET 'users'
    // ============================================
    const sheetUser = ss.getSheetByName("users");
    if (!sheetUser) return { status: "error", message: "Sheet 'users' belum ada di Spreadsheet." };

    let detailUser = null;
    const dataUser = sheetUser.getDataRange().getValues();
    
    // Loop cari profil siswa
    for(let i = 1; i < dataUser.length; i++) {
      if(String(dataUser[i][0]).trim() === targetKd) {
        detailUser = {
          nama: dataUser[i][5] || "Tanpa Nama", // Indeks 5: nama
          nis: dataUser[i][4] || "-",           // Indeks 4: nis
          kelas: dataUser[i][1] || "-",         // Indeks 1: fkd_kelas
          level: dataUser[i][7] || 1,           // Indeks 7: level
          wpm_terbaik: dataUser[i][11] || 0,    // Indeks 11: total_wpm_terbaik
          acc_terbaik: dataUser[i][12] || 0,    // Indeks 12: total_accuracy_terbaik
          gelar: dataUser[i][14] || "-"         // Indeks 14: gelar
        };
        break;
      }
    }

    if (!detailUser) return { status: "error", message: "Profil siswa tidak ditemukan di database." };

    // ============================================
    // 2. TARIK DATA LATIHAN DARI SHEET 'score'
    // ============================================
    const sheetScore = ss.getSheetByName("score");
    let riwayat = [];
    
    // Pastikan sheetScore ada dan datanya lebih dari 1 baris (header)
    if(sheetScore && sheetScore.getLastRow() > 1) {
      const dataScore = sheetScore.getDataRange().getValues();
      for(let i = 1; i < dataScore.length; i++) {
        // Cek fkd_user di Indeks 1
        if(String(dataScore[i][1]).trim() === targetKd) {
          riwayat.push({
            wpm: dataScore[i][3] || 0,          // Indeks 3: wpm
            acc: dataScore[i][4] || 0,          // Indeks 4: accuracy
            salah: dataScore[i][6] || 0,        // Indeks 6: total_salah
            durasi: dataScore[i][7] || 0,       // Indeks 7: durasi_detik
            tanggal: dataScore[i][13] || new Date() // Indeks 13: tgl_main
          });
        }
      }
    }

    // Urutkan riwayat dari yang paling baru (Descending)
    riwayat.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

    // Kirim balikan sukses
    return {
      status: "success",
      detail: detailUser,
      history: riwayat,
      total_main: riwayat.length
    };

  } catch(e) {
    // Jika ada error internal, kirim pesan error agar tidak stuck loading
    return { status: "error", message: "Sistem gagal membaca data: " + e.message };
  }
}

/**
 * Mengambil daftar kelas aktif untuk Dropdown di Form Registrasi User
 */
function getListKelasAktif() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetKelas = ss.getSheetByName("kelas");
    if(!sheetKelas) return [];
    
    const data = sheetKelas.getDataRange().getValues();
    let listKelas = [];
    
    // Looping dari baris 2 (mengabaikan header)
    for(let i = 1; i < data.length; i++) {
      // Cek apakah status_aktif (Kolom F / Indeks 5) bernilai 1
      if(String(data[i][5]).trim() === "1") { 
        listKelas.push({
          kd_kelas: data[i][0],    // Kolom A: kd_kelas
          nama_kelas: data[i][3]   // Kolom D: nama_kelas
        });
      }
    }
    
    // Urutkan sesuai abjad nama kelas
    listKelas.sort((a, b) => a.nama_kelas.localeCompare(b.nama_kelas));
    return listKelas;
  } catch(e) {
    return []; // Jika error, kembalikan array kosong agar UI tidak rusak
  }
}