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