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
    let nisDb = String(dataUser[i][4]).trim(); // Kolom E (Indeks 4) = NIS
    if(nisDb === nisTarget && nisTarget !== "" && nisTarget !== "-") {
      return false; // NIS duplikat
    }
  }
  return true; // Aman
}

/**
 * Menyimpan data pengguna baru ke dalam Sheet 'users' (Total 22 Kolom)
 */
function simpanUserBaru(dataSiswa) {
  // Hanya validasi NIS jika yang didaftarkan adalah SISWA (Guru mungkin NIS-nya kosong/-)
  if (dataSiswa.role === "SISWA" && !cekNisAman(dataSiswa.nis)) {
    return { status: "error", message: "NIS " + dataSiswa.nis + " sudah terdaftar di database!" };
  }
  
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetUser = ss.getSheetByName("users");
    let now = new Date();
    let kdUser = "USR-" + now.getTime();
    
    // Pemetaan berurutan sesuai kolom Sheet 'users' (Indeks 0 s.d 21)
    sheetUser.appendRow([
      kdUser,                            // 0: kd_user
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
      "",                                // 15: login_terakhir (Kosong karena belum login)
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
 * Mengambil seluruh data pengguna (Guru & Siswa) untuk ditampilkan di tabel Data Master
 */
function getListSiswa() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return [];

  const data = sheetUser.getDataRange().getValues();
  let listSiswa = [];

  // Looping dari baris 2 (mengabaikan header)
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
      status: data[i][17] || 'Y'    // Membaca status_aktif pada indeks 17
    });
  }
  
  // Mengurutkan berdasarkan Role lalu Nama
  listSiswa.sort((a, b) => a.role.localeCompare(b.role) || a.nama.localeCompare(b.nama));
  
  return listSiswa;
}