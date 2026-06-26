/**
 * Mengambil semua data dari Bank Kalimat
 */
function getListKalimat() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetKalimat = ss.getSheetByName("kalimat");
    if(!sheetKalimat) return { status: "error", message: "Sheet 'kalimat' belum dibuat." };
    
    const data = sheetKalimat.getDataRange().getValues();
    let listData = [];
    
    // Mulai dari baris 2 (mengabaikan header)
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] !== "") {
        listData.push({
          kd_kalimat: data[i][0],
          kategori: data[i][1],
          level: data[i][2],
          teks: data[i][3]
        });
      }
    }
    return { status: "success", data: listData };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

/**
 * Menyimpan Kalimat Baru atau Mengupdate Kalimat yang sudah ada
 */
function simpanKalimat(dataInput) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetKalimat = ss.getSheetByName("kalimat");
    const data = sheetKalimat.getDataRange().getValues();
    
    let isUpdate = false;
    let rowIndex = -1;

    // Cek apakah ini mode Edit (Update) berdasarkan kd_kalimat
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]).trim() === String(dataInput.kd_kalimat).trim()) {
        isUpdate = true;
        rowIndex = i + 1; // +1 karena array mulai dari 0, baris sheet mulai dari 1
        break;
      }
    }

    if(isUpdate) {
      // Eksekusi Update Teks
      sheetKalimat.getRange(rowIndex, 2).setValue(dataInput.kategori);
      sheetKalimat.getRange(rowIndex, 3).setValue(dataInput.level);
      sheetKalimat.getRange(rowIndex, 4).setValue(dataInput.teks);
      return { status: "success", message: "Data kalimat berhasil diperbarui!" };
    } else {
      // Eksekusi Simpan Baru
      // Buat ID Otomatis jika kosong
      let newId = dataInput.kd_kalimat;
      if(!newId || newId === "") {
        newId = "KLM-" + new Date().getTime();
      }
      sheetKalimat.appendRow([newId, dataInput.kategori, dataInput.level, dataInput.teks]);
      return { status: "success", message: "Kalimat baru berhasil ditambahkan ke Bank!" };
    }
  } catch (e) {
    return { status: "error", message: "Gagal menyimpan: " + e.message };
  }
}

/**
 * Menghapus Kalimat dari Database
 */
function hapusKalimat(kd_kalimat) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheetKalimat = ss.getSheetByName("kalimat");
    const data = sheetKalimat.getDataRange().getValues();
    
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]).trim() === String(kd_kalimat).trim()) {
        sheetKalimat.deleteRow(i + 1);
        return { status: "success", message: "Kalimat berhasil dihapus permanen!" };
      }
    }
    return { status: "error", message: "Kode kalimat tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: "Gagal menghapus: " + e.message };
  }
}

/**
 * Mengambil satu kalimat secara acak dari database yang cocok dengan jurusan siswa
 */
function ambilKalimatAcakSesuaiJurusan(kd_user) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    
    // 1. Identifikasi kelas dan jurusan siswa
    const sheetUser = ss.getSheetByName("users");
    const dataUser = sheetUser.getDataRange().getValues();
    let fkdKelas = "";
    
    for(let i = 1; i < dataUser.length; i++) {
      if(String(dataUser[i][0]).trim() === kd_user) {
        fkdKelas = String(dataUser[i][1]).trim(); 
        break;
      }
    }
    
    let singkatanProdi = "UMUM";
    if(fkdKelas !== "") {
      const sheetKelas = ss.getSheetByName("kelas");
      const dataKelas = sheetKelas.getDataRange().getValues();
      let fkdProdi = "";
      for(let i = 1; i < dataKelas.length; i++) {
        if(String(dataKelas[i][0]).trim() === fkdKelas) {
          fkdProdi = String(dataKelas[i][1]).trim(); 
          break;
        }
      }
      
      if(fkdProdi !== "") {
        const sheetProdi = ss.getSheetByName("programstudi");
        const dataProdi = sheetProdi.getDataRange().getValues();
        for(let i = 1; i < dataProdi.length; i++) {
          if(String(dataProdi[i][0]).trim() === fkdProdi) {
            singkatanProdi = String(dataProdi[i][2]).trim().toUpperCase(); 
            break;
          }
        }
      }
    }
    
    // 2. Filter kalimat yang sesuai dengan Jurusan Aktif ATAU kategori UMUM
    const sheetKalimat = ss.getSheetByName("kalimat");
    const dataKalimat = sheetKalimat.getDataRange().getValues();
    let poolKalimat = [];
    
    for(let i = 1; i < dataKalimat.length; i++) {
      let katKalimat = String(dataKalimat[i][1]).trim().toUpperCase(); // Kolom B: kategori_jurusan
      if(katKalimat === singkatanProdi || katKalimat === "UMUM") {
        poolKalimat.push({
          kd_kalimat: dataKalimat[i][0],
          teks: dataKalimat[i][3],
          level: dataKalimat[i][2]
        });
      }
    }
    
    // Bukaan Pengaman: Jika belum ada kalimat khusus prodi, ambil dari semua stok yang ada
    if(poolKalimat.length === 0 && dataKalimat.length > 1) {
      for(let i = 1; i < dataKalimat.length; i++) {
        poolKalimat.push({
          kd_kalimat: dataKalimat[i][0],
          teks: dataKalimat[i][3],
          level: dataKalimat[i][2]
        });
      }
    }
    
    if(poolKalimat.length === 0) {
      return { status: "error", message: "Bank Kalimat masih kosong. Silakan Admin isi terlebih dahulu." };
    }
    
    // Kocok data dan ambil satu kalimat secara acak
    let acakIndex = Math.floor(Math.random() * poolKalimat.length);
    return { status: "success", data: poolKalimat[acakIndex] };
    
  } catch(e) {
    return { status: "error", message: e.message };
  }
}