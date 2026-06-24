/**
 * Mengecek apakah NIS sudah terdaftar di sistem.
 * @param {string|number} nisBaru
 * @return {boolean} TRUE jika NIS belum ada (aman), FALSE jika dobel.
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

/**
 * Contoh fungsi saat Anda/Admin menambahkan user baru
 */
function simpanUserBaru(dataSiswa) {
  // Panggil validasi terlebih dahulu
  if (!cekNisAman(dataSiswa.nis)) {
    return { status: "error", message: "NIS " + dataSiswa.nis + " sudah terdaftar di database!" };
  }
  
  // Logika penyimpanan row baru bisa ditambahkan di bawah ini jika diperlukan ke depannya
}

/**
 * Mengambil seluruh data siswa untuk ditampilkan di tabel Data Master
 * @return {Array} Array of Object berisi daftar siswa
 */
function getListSiswa() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  if (!sheetUser) return [];

  const data = sheetUser.getDataRange().getValues();
  let listSiswa = [];

  // Looping dari baris 2 (mengabaikan header)
  for (let i = 1; i < data.length; i++) {
    // Hanya ambil yang role-nya "SISWA"
    if (data[i][6] === "SISWA") {
      listSiswa.push({
        kd_user: data[i][0],          // Kolom A
        kelas: data[i][1] || '-',     // Kolom B (fkd_kelas)
        username: data[i][2],         // Kolom C
        password: data[i][3],         // Kolom D
        nis: data[i][4],              // Kolom E
        nama: data[i][5],             // Kolom F
        level: data[i][7] || 1,       // Kolom H
        status: data[i][18] || 'Y'    // Kolom S (status_aktif)
      });
    }
  }
  
  // Mengurutkan berdasarkan Nama (Alphabetical)
  listSiswa.sort((a, b) => a.nama.localeCompare(b.nama));
  
  return listSiswa;
}