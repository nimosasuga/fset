// ===================================================================================
// FILE: 02_Auth.gs
// DESKRIPSI: Menangani Otorisasi (Login) dan Pengaturan Akun Karyawan
// ===================================================================================

function prosesLogin(nrpp, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKaryawan = ss.getSheetByName("Master_Karyawan");
  const dataKaryawan = sheetKaryawan.getDataRange().getValues();

  let userValid = false;
  let dataUser = {};

  for (let i = 1; i < dataKaryawan.length; i++) {
    if (dataKaryawan[i][0].toString() === nrpp.toString() && dataKaryawan[i][3].toString() === password.toString()) {
      userValid = true;
      dataUser = { nrpp: dataKaryawan[i][0], nama: dataKaryawan[i][1], jabatan: dataKaryawan[i][2] };
      break;
    }
  }

  if (!userValid) return { success: false, pesan: "NRPP atau Password salah!" };

  const sheetLog = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheetLog.getDataRange().getValues();
  dataUser.status_trip = "STANDBY";

  for (let i = dataLog.length - 1; i >= 1; i--) {
    if (dataLog[i][3].toString() === nrpp.toString() && dataLog[i][12] === "OUT") {
      dataUser.status_trip = "OUT";
      break;
    }
  }

  return { success: true, data: dataUser };
}

function prosesGantiPassword(nrpp, passLama, passBaru) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nrpp.toString()) {
      if (data[i][3].toString() === passLama.toString()) {
        sheet.getRange(i + 1, 4).setValue(passBaru.toString());
        return { success: true, pesan: "Password berhasil diperbarui!" };
      } else {
        return { success: false, pesan: "Password Lama tidak sesuai!" };
      }
    }
  }
  return { success: false, pesan: "User tidak ditemukan!" };
}
