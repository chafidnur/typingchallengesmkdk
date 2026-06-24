/**
 * Fungsi Autentikasi Login Dasar (Versi 1.0)
 * @param {string} username
 * @param {string} password
 * @return {Object} Status login dan profile singkat
 */
function prosesLogin(username, password) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetUser = ss.getSheetByName("users");
  
  // Validasi jika sheet 'users' belum dibuat di Spreadsheet
  if (!sheetUser) {
    return { status: "error", message: "Sistem Error: Sheet 'users' tidak ditemukan di database!" };
  }

  const dataUser = sheetUser.getDataRange().getValues();
  
  // Skip baris 1 (Header), mulai cari dari baris 2 (index 1)
  for(let i = 1; i < dataUser.length; i++) {
    
    // PERBAIKAN KRITIS: Paksa tipe data menjadi String dan hilangkan spasi kosong
    // Ini mencegah error jika di Excel password diset angka 1234, tapi JS membaca "1234"
    let dbUser = String(dataUser[i][2]).trim(); 
    let dbPass = String(dataUser[i][3]).trim(); 
    
    let inputUser = String(username).trim();
    let inputPass = String(password).trim();
    
    if(dbUser === inputUser && dbPass === inputPass) {
      // Jika cocok, kembalikan data user
      return {
        status: "success",
        kd_user: dataUser[i][0], // Kolom A: kd_user
        nama: dataUser[i][5],    // Kolom F: nama
        role: dataUser[i][6]     // Kolom G: role (ADMIN/GURU/SISWA)
      };
    }
  }
  
  // Jika loop selesai dan tidak ada yang cocok
  return { status: "error", message: "Username atau Password salah!" };
}