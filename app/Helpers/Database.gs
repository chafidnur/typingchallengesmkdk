/**
 * Fungsi Test Koneksi Database (Membaca Sheet 'setting')
 * @return {Object} Objek key-value pasang pengaturan sistem
 */
function getAppConfig() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheet = ss.getSheetByName("setting");
  
  // Pengaman jika sheet belum ada
  if (!sheet) return {}; 
  
  const data = sheet.getDataRange().getValues();
  let config = {};
  
  // Skip baris 1 (Header), mulai dari baris 2
  for (let i = 1; i < data.length; i++) {
    config[data[i][1]] = data[i][2]; // setting_key = setting_value
  }
  return config;
}