// ===================================================================================
// FILE: 06_Automasi.gs
// DESKRIPSI: Cron Jobs & Trigger untuk Sistem Auto-Archive dan Auto-Check In
// ===================================================================================

function autoArchiveBulanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheetLog.getDataRange().getValues();

  if (dataLog.length <= 1) return;

  const headers = dataLog[0];
  const tglSekarang = new Date();

  const batasWaktuBulanIni = new Date(tglSekarang.getFullYear(), tglSekarang.getMonth(), 1);

  let rowsToKeep = [headers];
  let rowsToArchive = [];

  for (let i = 1; i < dataLog.length; i++) {
    let row = dataLog[i];
    let tglKeluar = parseSafeDate(row[1]);
    let statusTrip = row[12];

    if (statusTrip === "IN" && tglKeluar.getTime() < batasWaktuBulanIni.getTime()) {
      let stValue = String(row[5]).replace(/^'/, "");
      row[5] = "'" + stValue;
      rowsToArchive.push(row);
    } else {
      let stValue = String(row[5]).replace(/^'/, "");
      row[5] = "'" + stValue;
      rowsToKeep.push(row);
    }
  }

  if (rowsToArchive.length > 0) {
    const tahunArsip = tglSekarang.getFullYear();
    const namaSheetArsip = "Arsip_" + tahunArsip;
    let sheetArsip = ss.getSheetByName(namaSheetArsip);

    if (!sheetArsip) {
      sheetArsip = ss.insertSheet(namaSheetArsip);
      sheetArsip.appendRow(headers);
    }

    const barisAwalArsip = sheetArsip.getLastRow() + 1;

    sheetArsip.getRange(barisAwalArsip, 6, rowsToArchive.length, 1).setNumberFormat("@");
    sheetArsip.getRange(barisAwalArsip, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);

    sheetLog.clearContents();
    sheetLog.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

    sheetLog.getRange(2, 2, sheetLog.getLastRow(), 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheetLog.getRange(2, 6, sheetLog.getLastRow(), 1).setNumberFormat("@");
  }
}

function autoCheckInHarian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheetLog.getDataRange().getValues();

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");

  const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();
  let mapJabatan = {};
  for (let k = 1; k < dataKaryawan.length; k++) {
    mapJabatan[String(dataKaryawan[k][0])] = dataKaryawan[k][2];
  }

  let isUpdated = false;

  for (let i = 1; i < dataLog.length; i++) {
    if (!dataLog[i][0]) continue;

    let status = dataLog[i][12];
    let tglKeluar = parseSafeDate(dataLog[i][1]);
    let logDateStr = Utilities.formatDate(tglKeluar, "GMT+7", "yyyy-MM-dd");

    if (status === "OUT" && logDateStr !== todayStr) {
      let wktMasuk = new Date(tglKeluar.getFullYear(), tglKeluar.getMonth(), tglKeluar.getDate(), 23, 59, 0);

      let nrpp = String(dataLog[i][3]);
      let jabatanUser = mapJabatan[nrpp] || "";

      let durasiJam = (wktMasuk.getTime() - tglKeluar.getTime()) / (1000 * 60 * 60);
      let rincian = kalkulasiUPD(jabatanUser, tglKeluar, durasiJam);
      let updTotal = rincian.total;

      let baris = i + 1;
      sheetLog.getRange(baris, 3).setValue(wktMasuk);

      // PERBAIKAN: Kunci Mutlak format desimal agar tidak menjadi Waktu (Time)
      sheetLog.getRange(baris, 11).setValue(durasiJam);
      sheetLog.getRange(baris, 11).setNumberFormat("0.00");

      sheetLog.getRange(baris, 12).setValue(updTotal);
      sheetLog.getRange(baris, 13).setValue("IN");
      sheetLog.getRange(baris, 15).setValue("PENDING");

      isUpdated = true;
    }
  }

  if (isUpdated) {
    sheetLog.getRange(2, 2, sheetLog.getLastRow(), 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  }
}

function setupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]); // Bersihkan jadwal lama
  }

  // 1. Robot Arsip Perjalanan Dinas HRD (Jalan Jam 01:00 Pagi, Tgl 1)
  ScriptApp.newTrigger("autoArchiveBulanan").timeBased().onMonthDay(1).atHour(1).create();

  // 2. Robot Arsip Kendaraan GA (Jalan Jam 02:00 Pagi, Tgl 1)
  ScriptApp.newTrigger("autoArchiveKendaraanBulanan").timeBased().onMonthDay(1).atHour(2).create();

  // 3. Robot Auto Check-In Harian (Jalan Jam 00:01 Setiap Hari)
  ScriptApp.newTrigger("autoCheckInHarian").timeBased().everyDays(1).atHour(0).nearMinute(1).create();
}

// ===================================================================================
// FUNGSI BARU: ROBOT ARSIP OTOMATIS IZIN KENDARAAN (GA)
// ===================================================================================

function autoArchiveKendaraanBulanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetIzin = ss.getSheetByName("Izin_Kendaraan");
  if (!sheetIzin) return;

  const dataIzin = sheetIzin.getDataRange().getValues();
  if (dataIzin.length <= 1) return; // Jika kosong/hanya header, batalkan.

  const headers = dataIzin[0];
  const tglSekarang = new Date();

  // Menetapkan batas waktu: Tanggal 1 bulan ini, jam 00:00
  const batasWaktuBulanIni = new Date(tglSekarang.getFullYear(), tglSekarang.getMonth(), 1);

  let rowsToKeep = [headers];
  let rowsToArchive = [];

  for (let i = 1; i < dataIzin.length; i++) {
    let row = dataIzin[i];
    let tglSubmit = new Date(row[1]); // Kolom Waktu Submit (Kolom B)

    // Jika tanggal submit SEBELUM bulan ini (Bulan lalu), pindahkan ke Arsip
    if (tglSubmit.getTime() < batasWaktuBulanIni.getTime()) {
      // Amankan format Plat & Fleet agar tidak rusak saat dipindah
      let platAman = String(row[2]).replace(/^'/, "");
      row[2] = "'" + platAman;
      let fleetAman = String(row[10]).replace(/^'/, "");
      row[10] = "'" + fleetAman;

      rowsToArchive.push(row);
    } else {
      // Biarkan data bulan ini tetap di Sheet utama
      let platAman = String(row[2]).replace(/^'/, "");
      row[2] = "'" + platAman;
      let fleetAman = String(row[10]).replace(/^'/, "");
      row[10] = "'" + fleetAman;

      rowsToKeep.push(row);
    }
  }

  // Jika ada data yang harus diarsipkan
  if (rowsToArchive.length > 0) {
    const tahunArsip = tglSekarang.getFullYear();
    const namaSheetArsip = "Arsip_Kendaraan_" + tahunArsip;
    let sheetArsip = ss.getSheetByName(namaSheetArsip);

    // Buat Sheet Arsip jika belum ada
    if (!sheetArsip) {
      sheetArsip = ss.insertSheet(namaSheetArsip);
      sheetArsip.appendRow(headers);
      sheetArsip.getRange("A1:S1").setFontWeight("bold").setBackground("#f8fafc");
    }

    const barisAwalArsip = sheetArsip.getLastRow() + 1;

    // Kunci format sebelum paste (Plat, Fleet)
    sheetArsip.getRange(barisAwalArsip, 3, rowsToArchive.length, 1).setNumberFormat("@");
    sheetArsip.getRange(barisAwalArsip, 11, rowsToArchive.length, 1).setNumberFormat("@");

    // Paste Data ke Brankas Arsip
    sheetArsip.getRange(barisAwalArsip, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    sheetArsip.getRange(barisAwalArsip, 2, rowsToArchive.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");

    // Timpa sheet utama hanya dengan data bulan ini
    sheetIzin.clearContents();
    sheetIzin.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

    // Kembalikan format sheet utama
    sheetIzin.getRange(2, 2, sheetIzin.getLastRow(), 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheetIzin.getRange(2, 3, sheetIzin.getLastRow(), 1).setNumberFormat("@");
    sheetIzin.getRange(2, 11, sheetIzin.getLastRow(), 1).setNumberFormat("@");
  }
}
