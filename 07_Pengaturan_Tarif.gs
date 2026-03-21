// ===================================================================================
// FILE: 07_Pengaturan_Tarif.gs
// DESKRIPSI: CRUD API untuk Pengaturan Tarif UPD oleh Admin / HRD
// ===================================================================================

/**
 * Mengambil seluruh data master tarif dari Google Sheets
 * @return {Array} Array objek berisi daftar tarif per jabatan
 */
function getAdminTarifData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Tarif_UPD");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  let res = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      res.push({
        id: i + 1,
        jabatan: String(data[i][0]).toUpperCase(),
        upd_full: parseFloat(data[i][1]) || 0,
        upd_half: parseFloat(data[i][2]) || 0,
        makan: parseFloat(data[i][3]) || 0,
        makan_libur: parseFloat(data[i][4]) || 0,
        lain_kerja: parseFloat(data[i][5]) || 0,
        lain_libur: parseFloat(data[i][6]) || 0,
      });
    }
  }
  return res;
}

/**
 * Menyimpan perubahan tarif dari UI Admin kembali ke Google Sheets
 * @param {number} row - Nomor baris di Google Sheets
 * @param {Object} newData - Data nominal tarif yang baru
 */
function simpanTarifData(row, newData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Tarif_UPD");
  if (!sheet) return { success: false, pesan: "Sheet Master_Tarif_UPD tidak ditemukan!" };

  sheet.getRange(row, 2).setValue(newData.upd_full);
  sheet.getRange(row, 3).setValue(newData.upd_half);
  sheet.getRange(row, 4).setValue(newData.makan);
  sheet.getRange(row, 5).setValue(newData.makan_libur);
  sheet.getRange(row, 6).setValue(newData.lain_kerja);
  sheet.getRange(row, 7).setValue(newData.lain_libur);

  return { success: true, pesan: "Tarif untuk jabatan " + newData.jabatan + " berhasil diupdate!" };
}

/**
 * Menambahkan jabatan dan tarif baru ke baris terbawah Google Sheets
 * @param {Object} newData - Data nominal tarif baru beserta jabatannya
 */
function tambahTarifData(newData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Tarif_UPD");
  if (!sheet) return { success: false, pesan: "Sheet Master_Tarif_UPD tidak ditemukan!" };

  const jabatanBaru = String(newData.jabatan).trim().toUpperCase();

  // 1. Validasi agar tidak ada jabatan ganda / duplikat
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === jabatanBaru) {
      return { success: false, pesan: "Jabatan '" + jabatanBaru + "' sudah ada! Silakan gunakan fitur Edit (ikon pensil hijau)." };
    }
  }

  // 2. Tambahkan ke baris baru
  sheet.appendRow([jabatanBaru, newData.upd_full, newData.upd_half, newData.makan, newData.makan_libur, newData.lain_kerja, newData.lain_libur]);

  return { success: true, pesan: "Jabatan " + jabatanBaru + " berhasil ditambahkan ke daftar!" };
}
