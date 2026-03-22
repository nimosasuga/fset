// ===================================================================================
// FILE: 04_Admin_Rekap.gs
// DESKRIPSI: Menarik Data Riwayat, Arsip Karyawan, dan Fungsionalitas Panel Admin
// ===================================================================================

function getRiwayatUser(nrpp) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log_Perjalanan").getDataRange().getValues();
  let riwayat = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (!data[i][0]) continue;
    if (String(data[i][3]) === String(nrpp)) {
      let statusValidasi = data[i][14];
      if (!statusValidasi && data[i][12] === "IN") statusValidasi = "PENDING";

      let durasiAsli = parseDurasi(data[i][10]);
      let tglKeluarSafe = data[i][1] ? parseSafeDate(data[i][1]) : null;
      let tglMasukSafe = data[i][2] ? parseSafeDate(data[i][2]) : null;

      riwayat.push({
        id: data[i][0],
        iso_keluar: tglKeluarSafe ? Utilities.formatDate(tglKeluarSafe, "GMT+7", "yyyy-MM-dd'T'HH:mm") : "",
        iso_masuk: tglMasukSafe ? Utilities.formatDate(tglMasukSafe, "GMT+7", "yyyy-MM-dd'T'HH:mm") : "",
        tgl_keluar: tglKeluarSafe ? Utilities.formatDate(tglKeluarSafe, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
        tgl_masuk: tglMasukSafe ? Utilities.formatDate(tglMasukSafe, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
        st: formatST(data[i][5]),
        customer: data[i][8],
        lokasi: data[i][9],
        durasi: data[i][10] !== "" && data[i][12] === "IN" ? durasiAsli.toFixed(2) + " Jam" : "-",
        upd: data[i][11] ? "Rp " + parseFloat(data[i][11]).toLocaleString("id-ID") : "-",
        status: data[i][12],
        validasi: statusValidasi,
      });
    }
  }
  return riwayat;
}

function getRekapUPD(nrpp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();
  let jabatanUser = "";
  for (let k = 1; k < dataKaryawan.length; k++) {
    if (String(dataKaryawan[k][0]) === String(nrpp)) {
      jabatanUser = dataKaryawan[k][2];
      break;
    }
  }

  const dataLog = ss.getSheetByName("Log_Perjalanan").getDataRange().getValues();
  let rekapData = {};
  for (let i = 1; i < dataLog.length; i++) {
    if (String(dataLog[i][3]) === String(nrpp) && dataLog[i][12] === "IN") {
      let st = formatST(dataLog[i][5]);
      let waktuKeluar = parseSafeDate(dataLog[i][1]);

      let durasiJam = parseDurasi(dataLog[i][10]);
      let rincian = kalkulasiUPD(jabatanUser, waktuKeluar, durasiJam);

      if (!rekapData[st]) rekapData[st] = { upd_tlm: 0, uang_makan: 0, uang_makan_siang: 0, lain_lain: 0, total: 0, history: [] };

      rekapData[st].upd_tlm += rincian.upd_tlm;
      rekapData[st].uang_makan += rincian.uang_makan;
      rekapData[st].uang_makan_siang += rincian.uang_makan_siang;
      rekapData[st].lain_lain += rincian.lain_lain;
      rekapData[st].total += rincian.total;
      rekapData[st].history.push({ tanggal: Utilities.formatDate(waktuKeluar, "GMT+7", "dd/MM/yyyy"), durasi: durasiJam.toFixed(2), upd_harian: rincian.total });
    }
  }
  return rekapData;
}

function getRekapArsip(nrpp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();
  let jabatanUser = "";
  for (let k = 1; k < dataKaryawan.length; k++) {
    if (String(dataKaryawan[k][0]) === String(nrpp)) {
      jabatanUser = dataKaryawan[k][2];
      break;
    }
  }

  const allSheets = ss.getSheets();
  let rekapData = {};
  allSheets.forEach((sheet) => {
    if (sheet.getName().startsWith("Arsip_")) {
      const dataLog = sheet.getDataRange().getValues();
      for (let i = 1; i < dataLog.length; i++) {
        if (String(dataLog[i][3]) === String(nrpp) && dataLog[i][12] === "IN") {
          let st = formatST(dataLog[i][5]);
          let waktuKeluar = parseSafeDate(dataLog[i][1]);

          let durasiJam = parseDurasi(dataLog[i][10]);
          let rincian = kalkulasiUPD(jabatanUser, waktuKeluar, durasiJam);

          let month = String(waktuKeluar.getMonth() + 1).padStart(2, "0");
          let year = waktuKeluar.getFullYear();
          let monthYearKey = `${year}-${month}`;

          if (!rekapData[monthYearKey]) rekapData[monthYearKey] = {};
          if (!rekapData[monthYearKey][st]) rekapData[monthYearKey][st] = { upd_tlm: 0, uang_makan: 0, uang_makan_siang: 0, lain_lain: 0, total: 0, history: [] };

          rekapData[monthYearKey][st].upd_tlm += rincian.upd_tlm;
          rekapData[monthYearKey][st].uang_makan += rincian.uang_makan;
          rekapData[monthYearKey][st].uang_makan_siang += rincian.uang_makan_siang;
          rekapData[monthYearKey][st].lain_lain += rincian.lain_lain;
          rekapData[monthYearKey][st].total += rincian.total;
          rekapData[monthYearKey][st].history.push({ tanggal: Utilities.formatDate(waktuKeluar, "GMT+7", "dd/MM/yyyy"), durasi: durasiJam.toFixed(2), upd_harian: rincian.total });
        }
      }
    }
  });

  const sortedKeys = Object.keys(rekapData).sort().reverse();
  let sortedRekap = {};
  sortedKeys.forEach((key) => (sortedRekap[key] = rekapData[key]));
  return sortedRekap;
}

function getAdminDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheetLog.getDataRange().getValues();

  let aktifOut = [];
  let totalUPD_BulanIni = 0;
  let tglSekarang = new Date();
  let mapPersonil = {};

  for (let i = dataLog.length - 1; i >= 1; i--) {
    if (!dataLog[i][0]) continue;
    let nrpp = String(dataLog[i][3]);
    let tglKeluar = parseSafeDate(dataLog[i][1]);
    let statusTrip = String(dataLog[i][12]).toUpperCase();
    let updHarian = parseFloat(dataLog[i][11]) || 0;

    if (!mapPersonil[nrpp]) {
      mapPersonil[nrpp] = true;
      if (statusTrip === "OUT") {
        aktifOut.push({
          nrpp: nrpp,
          nama: dataLog[i][4],
          st: formatST(dataLog[i][5]),
          customer: dataLog[i][8],
          lokasi: dataLog[i][9],
          waktu_keluar: Utilities.formatDate(tglKeluar, "GMT+7", "dd/MM/yyyy HH:mm"),
        });
      }
    }
    if (statusTrip === "IN" && tglKeluar.getMonth() === tglSekarang.getMonth() && tglKeluar.getFullYear() === tglSekarang.getFullYear()) {
      totalUPD_BulanIni += updHarian;
    }
  }
  return { aktifOut: aktifOut, totalOut: aktifOut.length, totalUPDBulanIni: totalUPD_BulanIni };
}

// PERBAIKAN: Menambahkan kolom Jabatan ke dalam array agar bisa ditarik di Frontend
function getListKaryawan() {
  const dataKaryawan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Karyawan").getDataRange().getValues();
  let list = [];
  for (let i = 1; i < dataKaryawan.length; i++) {
    if (dataKaryawan[i][0] !== "") {
      list.push({
        nrpp: dataKaryawan[i][0],
        nama: dataKaryawan[i][1],
        jabatan: dataKaryawan[i][2],
      });
    }
  }
  return list;
}

function getAllRekapForAdmin(nrpp) {
  return { aktif: getRekapUPD(nrpp), arsip: getRekapArsip(nrpp), log: getRiwayatUser(nrpp) };
}

function simpanKoreksiLog(idTransaksi, dataBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheet.getDataRange().getValues();

  let barisUpdate = -1;
  let targetNrpp = "";

  for (let i = 1; i < dataLog.length; i++) {
    if (dataLog[i][0] === idTransaksi) {
      barisUpdate = i + 1;
      targetNrpp = dataLog[i][3];
      break;
    }
  }

  if (barisUpdate === -1) return { success: false, pesan: "Data log tidak ditemukan di database." };

  const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();
  let jabatanUser = "";
  for (let k = 1; k < dataKaryawan.length; k++) {
    if (String(dataKaryawan[k][0]) === String(targetNrpp)) {
      jabatanUser = dataKaryawan[k][2];
      break;
    }
  }

  let tglKeluarObj = dataBaru.waktuKeluar ? parseSafeDate(dataBaru.waktuKeluar) : "";
  let tglMasukObj = dataBaru.waktuMasuk ? parseSafeDate(dataBaru.waktuMasuk) : "";
  let durasiJam = 0;
  let updTotal = 0;

  if (dataBaru.status === "IN" && tglKeluarObj && tglMasukObj) {
    durasiJam = (tglMasukObj.getTime() - tglKeluarObj.getTime()) / (1000 * 60 * 60);
    if (durasiJam < 0) return { success: false, pesan: "Waktu masuk tidak boleh lebih awal dari waktu keluar." };
    let rincian = kalkulasiUPD(jabatanUser, tglKeluarObj, durasiJam);
    updTotal = rincian.total;
  }

  const stAman = "'" + dataBaru.st;
  sheet.getRange(barisUpdate, 2).setValue(tglKeluarObj || "");
  sheet.getRange(barisUpdate, 3).setValue(tglMasukObj || "");
  sheet.getRange(barisUpdate, 6).setValue(stAman);
  sheet.getRange(barisUpdate, 9).setValue(dataBaru.customer);
  sheet.getRange(barisUpdate, 10).setValue(dataBaru.lokasi);

  if (dataBaru.status === "IN") {
    sheet.getRange(barisUpdate, 11).setValue(durasiJam);
    sheet.getRange(barisUpdate, 11).setNumberFormat("0.00");
  } else {
    sheet.getRange(barisUpdate, 11).setValue("");
  }

  sheet.getRange(barisUpdate, 12).setValue(dataBaru.status === "IN" ? updTotal : "");
  sheet.getRange(barisUpdate, 13).setValue(dataBaru.status);
  sheet.getRange(barisUpdate, 15).setValue(dataBaru.status === "IN" ? "ADJUSTED" : "");

  sheet.getRange(barisUpdate, 2, 1, 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sheet.getRange(barisUpdate, 6).setNumberFormat("@");

  return { success: true, pesan: "Data log berhasil dikoreksi dan status berubah menjadi Penyesuaian!" };
}

function approveLogAdmin(idTransaksi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheet.getDataRange().getValues();

  for (let i = 1; i < dataLog.length; i++) {
    if (dataLog[i][0] === idTransaksi) {
      sheet.getRange(i + 1, 15).setValue("APPROVED");
      return { success: true, pesan: "Log perjalanan berhasil di-Approve!" };
    }
  }
  return { success: false, pesan: "Data log tidak ditemukan." };
}

function getRiwayatArsipGrouped(nrpp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  let riwayatGrouped = {};

  allSheets.forEach((sheet) => {
    if (sheet.getName().startsWith("Arsip_")) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;

        if (String(data[i][3]) === String(nrpp)) {
          let tglKeluar = parseSafeDate(data[i][1]);
          let month = String(tglKeluar.getMonth() + 1).padStart(2, "0");
          let year = tglKeluar.getFullYear();
          let monthYearKey = `${year}-${month}`;

          if (!riwayatGrouped[monthYearKey]) riwayatGrouped[monthYearKey] = [];

          let statusValidasi = data[i][14];
          if (!statusValidasi && data[i][12] === "IN") statusValidasi = "PENDING";

          let durasiAsli = parseDurasi(data[i][10]);
          let tglMasukSafe = data[i][2] ? parseSafeDate(data[i][2]) : null;

          riwayatGrouped[monthYearKey].push({
            id: data[i][0],
            tgl_keluar: tglKeluar ? Utilities.formatDate(tglKeluar, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
            tgl_masuk: tglMasukSafe ? Utilities.formatDate(tglMasukSafe, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
            st: formatST(data[i][5]),
            customer: data[i][8],
            lokasi: data[i][9],
            durasi: data[i][10] !== "" && data[i][12] === "IN" ? durasiAsli.toFixed(2) + " Jam" : "-",
            upd: data[i][11] ? "Rp " + parseFloat(data[i][11]).toLocaleString("id-ID") : "-",
            status: data[i][12],
            validasi: statusValidasi,
          });
        }
      }
    }
  });

  const sortedKeys = Object.keys(riwayatGrouped).sort().reverse();
  let sortedRiwayat = {};
  sortedKeys.forEach((key) => (sortedRiwayat[key] = riwayatGrouped[key]));
  return sortedRiwayat;
}

function getArsipKaryawanLengkap(nrpp) {
  return {
    arsipUPD: getRekapArsip(nrpp),
    arsipLog: getRiwayatArsipGrouped(nrpp),
  };
}

// ===================================================================================
// FUNGSI BARU: RADAR NOTIFIKASI REAL-TIME ABSEN POS (HRD)
// ===================================================================================

// ===================================================================================
// FUNGSI UPDATE: RADAR NOTIFIKASI REAL-TIME ABSEN POS (HRD)
// ===================================================================================

function cekNotifikasiAbsenBaruHRD(stateLokal) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log_Perjalanan");
    if (!sheet) return { hasChanged: false, newState: stateLokal };

    const totalBaris = sheet.getLastRow();

    // Tarik data saat ini dari fungsi dashboard yang sudah ada
    const currentData = getAdminDashboardData();
    const totalOut = currentData.totalOut;

    // KUNCI: Buat "State Signature" (Gabungan jumlah baris & jumlah orang di luar)
    const currentState = totalBaris + "_" + totalOut;

    // Jika State berubah (ada yang keluar ATAU ada yang kembali)
    if (currentState !== stateLokal) {
      // Kita kirimkan data terbarunya sekalian ke Frontend (Menghemat 1x pemanggilan server!)
      return { hasChanged: true, newState: currentState, data: currentData };
    }

    return { hasChanged: false, newState: currentState };
  } catch (e) {
    return { hasChanged: false, newState: stateLokal };
  }
}

/// ===================================================================================
// PERBAIKAN FUNGSI: MENARIK DATA YANG BUTUH APPROVAL (DENGAN DATA LENGKAP UNTUK MODAL)
// ===================================================================================
function getPendingApprovalHRD() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log_Perjalanan");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let pendingList = [];

    // Looping dari bawah ke atas agar data terbaru muncul duluan
    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];

      let status = String(row[12] || "")
        .trim()
        .toUpperCase();
      let validasi = String(row[14] || "")
        .trim()
        .toUpperCase();

      if (status === "IN" && validasi !== "APPROVED") {
        let tglKeluarObj = row[1] ? new Date(row[1]) : null;
        let tglMasukObj = row[2] ? new Date(row[2]) : null;

        pendingList.push({
          id: row[0],
          nama: row[4],
          st: row[5],
          customer: row[8], // Tambahan data untuk modal koreksi
          lokasi: row[9],
          status: status, // Tambahan data untuk modal koreksi
          iso_keluar: tglKeluarObj ? Utilities.formatDate(tglKeluarObj, "GMT+7", "yyyy-MM-dd'T'HH:mm") : "", // Tambahan untuk modal
          iso_masuk: tglMasukObj ? Utilities.formatDate(tglMasukObj, "GMT+7", "yyyy-MM-dd'T'HH:mm") : "", // Tambahan untuk modal
          waktuKeluar: tglKeluarObj ? Utilities.formatDate(tglKeluarObj, "GMT+7", "dd/MM HH:mm") : "-",
          waktuMasuk: tglMasukObj ? Utilities.formatDate(tglMasukObj, "GMT+7", "dd/MM HH:mm") : "-",
        });
      }
    }
    return pendingList;
  } catch (e) {
    return [];
  }
}
