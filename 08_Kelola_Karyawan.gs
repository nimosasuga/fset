// ===================================================================================
// FILE: 08_Kelola_Karyawan.gs
// DESKRIPSI: Modul CRUD (Create, Read, Update, Delete) untuk Master Data Karyawan
// ===================================================================================

/**
 * Mengambil seluruh data Karyawan dari Master_Karyawan
 * @return {Array} Daftar Karyawan
 */
function getAdminKaryawanData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  let listKaryawan = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      listKaryawan.push({
        baris: i + 1, // Untuk patokan saat Update/Delete
        nrpp: String(data[i][0]),
        nama: String(data[i][1]),
        jabatan: String(data[i][2]),
        password: String(data[i][3]),
      });
    }
  }
  return listKaryawan;
}

/**
 * Menambahkan Karyawan Baru ke database
 * @param {Object} dataBaru - Objek berisi {nrpp, nama, jabatan, password}
 * @return {Object} Status success dan pesan
 */
function tambahKaryawan(dataBaru) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan");
  if (!sheet) return { success: false, pesan: "Sheet Master_Karyawan tidak ditemukan!" };

  const nrppBaru = String(dataBaru.nrpp).trim();
  const data = sheet.getDataRange().getValues();

  // Validasi: Cek apakah NRPP sudah terdaftar
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === nrppBaru) {
      return { success: false, pesan: "Gagal! NRPP " + nrppBaru + " sudah digunakan oleh " + data[i][1] };
    }
  }

  // Format NRPP sebagai Teks agar nol di depan tidak hilang
  const nrppAman = "'" + nrppBaru;

  sheet.appendRow([nrppAman, String(dataBaru.nama).toUpperCase(), String(dataBaru.jabatan).toUpperCase(), String(dataBaru.password)]);

  return { success: true, pesan: "Karyawan " + dataBaru.nama + " berhasil ditambahkan!" };
}

/**
 * Menyimpan perubahan data karyawan (Edit)
 * @param {string} nrppLama - NRPP sebelum diedit (sebagai kata kunci pencarian)
 * @param {Object} dataEdit - Data baru dari form
 * @return {Object} Status
 */
function editKaryawan(nrppLama, dataEdit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan");
  const data = sheet.getDataRange().getValues();

  let barisTarget = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(nrppLama).trim()) {
      barisTarget = i + 1;
      break;
    }
  }

  if (barisTarget === -1) return { success: false, pesan: "Data karyawan tidak ditemukan!" };

  const nrppAman = "'" + String(dataEdit.nrpp).trim();

  sheet.getRange(barisTarget, 1).setValue(nrppAman);
  sheet.getRange(barisTarget, 2).setValue(String(dataEdit.nama).toUpperCase());
  sheet.getRange(barisTarget, 3).setValue(String(dataEdit.jabatan).toUpperCase());
  sheet.getRange(barisTarget, 4).setValue(String(dataEdit.password));

  return { success: true, pesan: "Data karyawan berhasil diperbarui!" };
}

/**
 * Menghapus data karyawan berdasarkan NRPP
 * @param {string} nrpp - NRPP target
 * @return {Object} Status
 */
function hapusKaryawan(nrpp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(nrpp).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, pesan: "Karyawan dengan NRPP " + nrpp + " berhasil dihapus!" };
    }
  }

  return { success: false, pesan: "Data tidak ditemukan." };
}
