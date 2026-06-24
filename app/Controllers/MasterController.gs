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