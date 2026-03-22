// ===================================================================================
// FILE: 09_Izin_Kendaraan.gs
// DESKRIPSI: Modul untuk menangani Izin Keluar Kendaraan Dinas (General Affair)
// ===================================================================================

function simpanIzinKendaraan(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Izin_Kendaraan");

    if (!sheet) {
      sheet = ss.insertSheet("Izin_Kendaraan");
    }

    // Header diperbarui sampai Kolom S (SISA_SALDO)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "ID_IZIN",
        "WAKTU_SUBMIT",
        "PLAT_KENDARAAN",
        "NRPP_PEMOHON",
        "NAMA_PEMOHON",
        "PARTNER_DIPILIH",
        "CUSTOMER",
        "ALAMAT",
        "DEPARTEMEN",
        "KM_AWAL",
        "KARTU_FLEET",
        "STATUS_GA",
        "BBM",
        "TOL",
        "PARKIR",
        "LAIN_LAIN",
        "TOTAL_BIAYA",
        "TOP_UP_FLEET",
        "SISA_SALDO",
      ]);
      sheet.getRange("A1:S1").setFontWeight("bold").setBackground("#f8fafc");
    }

    const timestamp = new Date();
    const idIzin = "GA-" + data.nrpp + "-" + Utilities.formatDate(timestamp, "GMT+7", "yyMMddHHmm");
    const partnerTeks = data.partner && data.partner.length > 0 ? data.partner.join(", ") : "Tidak Ada / Sendiri";

    // Paksa Kartu Fleet menjadi Teks agar tidak jadi 8.9E+15
    const fleetAman = data.kartu_fleet ? "'" + data.kartu_fleet : "";

    sheet.appendRow([idIzin, timestamp, data.plat, data.nrpp, data.nama, partnerTeks, data.customer, data.alamat, data.departemen, data.km_awal, fleetAman, "PENDING"]);

    const barisBaru = sheet.getLastRow();
    sheet.getRange(barisBaru, 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");

    return { success: true, pesan: "Formulir Izin Kendaraan berhasil dikirim ke GA!" };
  } catch (error) {
    return { success: false, pesan: error.toString() };
  }
}

// ===================================================================================
// FUNGSI UPDATE: TARIK DATA DENGAN ALGORITMA RUNNING BALANCE (MUTASI REAL-TIME)
// ===================================================================================

// 1. FUNGSI MENARIK DATA KE DASHBOARD GA (FIX BUG RINCIAN BIAYA & TAMBAH KM AKHIR)
function getIzinKendaraanData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Izin_Kendaraan");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let fleetBalances = {};

    let arsipSheets = ss.getSheets().filter((s) => s.getName().startsWith("Arsip_Kendaraan_"));
    for (let s of arsipSheets) {
      let dataArsip = s.getDataRange().getValues();
      for (let i = 1; i < dataArsip.length; i++) {
        if (!dataArsip[i][0]) continue;
        let f = String(dataArsip[i][10]).replace(/^'/, "").trim();
        if (f && f !== "-") {
          let biaya = parseFloat(dataArsip[i][16]) || 0;
          let topup = parseFloat(dataArsip[i][17]) || 0;
          fleetBalances[f] = (fleetBalances[f] || 0) + topup - biaya;
        }
      }
    }

    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      let f = String(data[i][10]).replace(/^'/, "").trim();
      let biaya = parseFloat(data[i][16]) || 0;
      let topup = parseFloat(data[i][17]) || 0;
      let currentSaldo = 0;

      if (f && f !== "-") {
        fleetBalances[f] = (fleetBalances[f] || 0) + topup - biaya;
        currentSaldo = fleetBalances[f];
      } else {
        currentSaldo = topup - biaya;
      }

      let tglObj = new Date(data[i][1]);
      let strWaktu = isNaN(tglObj.getTime()) ? "-" : Utilities.formatDate(tglObj, "GMT+7", "dd/MM/yyyy HH:mm");

      result.push({
        id: data[i][0],
        waktu: strWaktu,
        plat: data[i][2],
        nama: data[i][4],
        partner: data[i][5],
        customer: data[i][6],
        alamat: data[i][7],
        departemen: data[i][8],
        km: data[i][9] || "-", // KM Awal
        kmAkhir: data[i][19] || "", // TAMBAHAN: KM Akhir (Kolom T / Index 19)
        fleet: f || "-",
        status: data[i][11],
        bbm: data[i][12] || 0, // FIX BUG: Rincian biaya dikirim ke GA
        tol: data[i][13] || 0,
        parkir: data[i][14] || 0,
        lain: data[i][15] || 0,
        totalBiaya: biaya,
        topup: topup,
        saldo: currentSaldo,
      });
    }
    return result.reverse();
  } catch (e) {
    return [];
  }
}

// 2. FUNGSI SIMPAN PENGELUARAN DARI USER (TAMBAH KM AKHIR)
function simpanPengeluaranKendaraan(idIzin, biaya) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();

    // Pastikan Header dibuat sampai Kolom T (KM_AKHIR)
    if (data[0].length < 20) {
      sheet.getRange("M1:T1").setValues([["BBM", "TOL", "PARKIR", "LAIN_LAIN", "TOTAL_BIAYA", "TOP_UP_FLEET", "SISA_SALDO", "KM_AKHIR"]]);
      sheet.getRange("M1:T1").setFontWeight("bold").setBackground("#f8fafc");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idIzin) {
        let valBbm = Number(biaya.bbm) || 0;
        let valTol = Number(biaya.tol) || 0;
        let valParkir = Number(biaya.parkir) || 0;
        let valLain = Number(biaya.lain) || 0;
        let kmAkhir = biaya.kmAkhir || ""; // Data KM Akhir dari User

        let total = valBbm + valTol + valParkir + valLain;
        let valTopup = Number(data[i][17]) || 0;
        let sisaSaldo = valTopup - total;

        sheet.getRange(i + 1, 13).setValue(valBbm);
        sheet.getRange(i + 1, 14).setValue(valTol);
        sheet.getRange(i + 1, 15).setValue(valParkir);
        sheet.getRange(i + 1, 16).setValue(valLain);
        sheet.getRange(i + 1, 17).setValue(total);
        sheet.getRange(i + 1, 19).setValue(sisaSaldo);
        sheet.getRange(i + 1, 20).setValue(kmAkhir); // Simpan KM Akhir ke Kolom T

        return { success: true, pesan: "Pengeluaran & KM Akhir berhasil disimpan!" };
      }
    }
    return { success: false, pesan: "Data Izin tidak ditemukan." };
  } catch (error) {
    return { success: false, pesan: error.toString() };
  }
}

// 3. FUNGSI KOREKSI GA (EDIT FORM - TAMBAH KM AKHIR)
function updateKoreksiIzinGA(idIzin, dataBaru) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();
    let fleetTarget = "";

    // Pastikan Header dibuat sampai Kolom T (KM_AKHIR)
    if (data[0].length < 20) {
      sheet.getRange("M1:T1").setValues([["BBM", "TOL", "PARKIR", "LAIN_LAIN", "TOTAL_BIAYA", "TOP_UP_FLEET", "SISA_SALDO", "KM_AKHIR"]]);
      sheet.getRange("M1:T1").setFontWeight("bold").setBackground("#f8fafc");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idIzin) {
        const row = i + 1;
        let fleetAman = dataBaru.fleet ? "'" + dataBaru.fleet : "";
        fleetTarget = dataBaru.fleet ? String(dataBaru.fleet).replace(/^'/, "").trim() : "";

        sheet.getRange(row, 3).setValue(dataBaru.plat.toUpperCase());
        sheet.getRange(row, 6).setValue(dataBaru.partner);
        sheet.getRange(row, 7).setValue(dataBaru.customer.toUpperCase());
        sheet.getRange(row, 8).setValue(dataBaru.alamat);
        sheet.getRange(row, 9).setValue(dataBaru.departemen.toUpperCase());
        sheet.getRange(row, 10).setValue(dataBaru.km); // KM Awal
        sheet.getRange(row, 11).setValue(fleetAman);

        let valBbm = parseFloat(dataBaru.bbm) || 0;
        let valTol = parseFloat(dataBaru.tol) || 0;
        let valParkir = parseFloat(dataBaru.parkir) || 0;
        let valLain = parseFloat(dataBaru.lain) || 0;
        let valTopup = parseFloat(dataBaru.topup) || 0;
        let total = valBbm + valTol + valParkir + valLain;

        sheet.getRange(row, 13).setValue(valBbm);
        sheet.getRange(row, 14).setValue(valTol);
        sheet.getRange(row, 15).setValue(valParkir);
        sheet.getRange(row, 16).setValue(valLain);
        sheet.getRange(row, 17).setValue(total);
        sheet.getRange(row, 18).setValue(valTopup);
        sheet.getRange(row, 20).setValue(dataBaru.kmAkhir || ""); // Koreksi KM Akhir
        break;
      }
    }

    if (fleetTarget && fleetTarget !== "-") {
      recalculateLedgerFleet(fleetTarget);
    }
    return { success: true, pesan: "Data terkoreksi & Mutasi diperbarui!" };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}

// 4. FUNGSI MENARIK RIWAYAT USER (AGAR SAAT DI-EDIT, KM AKHIR MUNCUL)
function getRiwayatIzinKendaraanUser(nrpp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const limit = 150;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 20).getValues(); // Ambil sampai kolom 20

  let result = [];
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][3]) === String(nrpp)) {
      let tglPengajuan = new Date(data[i][1]);

      if (tglPengajuan.getMonth() === currentMonth && tglPengajuan.getFullYear() === currentYear) {
        result.push({
          id: data[i][0],
          waktu: Utilities.formatDate(tglPengajuan, "GMT+7", "dd/MM/yyyy HH:mm"),
          plat: data[i][2],
          partner: data[i][5],
          customer: data[i][6],
          alamat: data[i][7],
          departemen: data[i][8],
          km: data[i][9] || "-",
          fleet: data[i][10] || "-",
          status: data[i][11],
          bbm: data[i][12] || 0,
          tol: data[i][13] || 0,
          parkir: data[i][14] || 0,
          lain: data[i][15] || 0,
          totalBiaya: data[i][16] || 0,
          kmAkhir: data[i][19] || "", // Tarik KM Akhir
        });
      }
    }
  }
  return result;
}

function updateStatusIzinGA(idIzin, statusBaru) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idIzin) {
      sheet.getRange(i + 1, 12).setValue(statusBaru);
      return { success: true, pesan: "Status izin kendaraan berhasil diperbarui!" };
    }
  }
  return { success: false, pesan: "ID Izin tidak ditemukan di database." };
}

// --- FUNGSI UPDATE: GA MENGKOREKSI DATA, BIAYA, TOP UP & SALDO ---
function updateKoreksiIzinGA(idIzin, dataBaru) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();

    // Pastikan Header lengkap sampai Kolom S
    if (data[0].length < 19) {
      sheet.getRange("M1:S1").setValues([["BBM", "TOL", "PARKIR", "LAIN_LAIN", "TOTAL_BIAYA", "TOP_UP_FLEET", "SISA_SALDO"]]);
      sheet.getRange("M1:S1").setFontWeight("bold").setBackground("#f8fafc");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idIzin) {
        const row = i + 1;

        let fleetAman = dataBaru.fleet ? "'" + dataBaru.fleet : "";

        sheet.getRange(row, 3).setValue(dataBaru.plat.toUpperCase());
        sheet.getRange(row, 6).setValue(dataBaru.partner);
        sheet.getRange(row, 7).setValue(dataBaru.customer.toUpperCase());
        sheet.getRange(row, 8).setValue(dataBaru.alamat);
        sheet.getRange(row, 9).setValue(dataBaru.departemen.toUpperCase());
        sheet.getRange(row, 10).setValue(dataBaru.km);
        sheet.getRange(row, 11).setValue(fleetAman);

        let valBbm = parseFloat(dataBaru.bbm) || 0;
        let valTol = parseFloat(dataBaru.tol) || 0;
        let valParkir = parseFloat(dataBaru.parkir) || 0;
        let valLain = parseFloat(dataBaru.lain) || 0;
        let valTopup = parseFloat(dataBaru.topup) || 0;

        // MATEMATIKA SALDO
        let total = valBbm + valTol + valParkir + valLain;
        let sisaSaldo = valTopup - total;

        sheet.getRange(row, 13).setValue(valBbm);
        sheet.getRange(row, 14).setValue(valTol);
        sheet.getRange(row, 15).setValue(valParkir);
        sheet.getRange(row, 16).setValue(valLain);
        sheet.getRange(row, 17).setValue(total);
        sheet.getRange(row, 18).setValue(valTopup);
        sheet.getRange(row, 19).setValue(sisaSaldo); // Menulis Saldo ke Kolom S

        return { success: true, pesan: "Data pengajuan & pengeluaran berhasil dikoreksi!" };
      }
    }
    return { success: false, pesan: "Data Izin tidak ditemukan." };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}

// --- UPDATE FUNGSI QUICK SAVE (Pemicu Mutasi) ---
function quickSaveTopUpGA(idIzin, nominalTopUp) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();
    let fleetTarget = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idIzin) {
        fleetTarget = String(data[i][10]).replace(/^'/, "").trim();
        let valTopup = parseFloat(nominalTopUp) || 0;
        sheet.getRange(i + 1, 18).setValue(valTopup); // Simpan Top Up
        break;
      }
    }

    // Picu perhitungan mutasi berantai ke bawah
    if (fleetTarget && fleetTarget !== "-") {
      recalculateLedgerFleet(fleetTarget);
    }
    return { success: true, pesan: "Top Up berhasil & Mutasi dikalkulasi ulang!" };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}

// ===================================================================================
// FUNGSI UPDATE: ANALITIK KARTU FLEET (HANYA BULAN BERJALAN & ANTI-CRASH)
// ===================================================================================

function getDashboardAnalyticsGA() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let allData = [];

    // Ambil data dari sheet aktif
    let sheetAktif = ss.getSheetByName("Izin_Kendaraan");
    if (sheetAktif) {
      let data = sheetAktif.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) allData.push(data[i]);
      }
    }

    let currentMonth = new Date().getMonth();
    let currentYear = new Date().getFullYear();

    let totalBulanIni = 0;
    let bbmBulanIni = 0;
    let tolBulanIni = 0;
    let parkirBulanIni = 0;
    let lainBulanIni = 0;

    let breakdownFleet = {};

    allData.forEach((row) => {
      let tgl = new Date(row[1]);
      if (isNaN(tgl.getTime())) tgl = new Date();

      // LOGIKA KUNCI: HANYA proses data jika bulannya adalah bulan ini
      if (tgl.getMonth() === currentMonth && tgl.getFullYear() === currentYear) {
        let fleetRaw = row[10] ? String(row[10]).replace(/^'/, "").trim() : "";
        let fleet = fleetRaw;

        // Kelompokkan yang tidak pakai kartu menjadi CASH
        if (fleet === "-" || fleet === "" || fleet.toLowerCase() === "undefined") {
          fleet = "TANPA KARTU (CASH)";
        }

        let bbm = parseFloat(row[12]) || 0;
        let tol = parseFloat(row[13]) || 0;
        let parkir = parseFloat(row[14]) || 0;
        let lain = parseFloat(row[15]) || 0;
        let biayaTotal = parseFloat(row[16]) || 0;

        // Tambahkan ke Total Utama
        totalBulanIni += biayaTotal;
        bbmBulanIni += bbm;
        tolBulanIni += tol;
        parkirBulanIni += parkir;
        lainBulanIni += lain;

        // Tambahkan ke Rincian Per Kartu Fleet
        if (!breakdownFleet[fleet]) {
          breakdownFleet[fleet] = { total: 0, bbm: 0, tol: 0, parkir: 0, lain: 0 };
        }
        breakdownFleet[fleet].total += biayaTotal;
        breakdownFleet[fleet].bbm += bbm;
        breakdownFleet[fleet].tol += tol;
        breakdownFleet[fleet].parkir += parkir;
        breakdownFleet[fleet].lain += lain;
      }
    });

    return {
      success: true,
      totalBulanIni: totalBulanIni,
      bbmBulanIni: bbmBulanIni,
      tolBulanIni: tolBulanIni,
      parkirBulanIni: parkirBulanIni,
      lainBulanIni: lainBulanIni,
      breakdownFleet: breakdownFleet,
      listFleet: Object.keys(breakdownFleet).sort(), // Mengurutkan nomor dari kecil ke besar
    };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}

function getListArsipKendaraan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let arsipList = [];
  ss.getSheets().forEach((s) => {
    if (s.getName().startsWith("Arsip_Kendaraan_")) {
      arsipList.push(s.getName());
    }
  });
  return arsipList.sort().reverse();
}

function getArsipKendaraanDetail(namaSheet) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let result = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (!data[i][0]) continue;
    result.push({
      waktu: Utilities.formatDate(new Date(data[i][1]), "GMT+7", "dd/MM/yyyy HH:mm"),
      plat: data[i][2],
      nama: data[i][4],
      customer: data[i][6],
      fleet: data[i][10] || "-",
      totalBiaya: data[i][16] || 0,
      topup: data[i][17] || 0,
    });
  }
  return result;
}

// ===================================================================================
// FUNGSI BARU: ALGORITMA BUKU MUTASI KARTU FLEET (RUNNING BALANCE)
// ===================================================================================
function recalculateLedgerFleet(nomorFleet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Cari Saldo Terakhir di Brankas Arsip (Sebagai Saldo Awal)
  let saldoAwal = 0;
  let sheets = ss.getSheets();
  let arsipSheets = sheets.filter((s) => s.getName().startsWith("Arsip_Kendaraan_")).sort((a, b) => b.getName().localeCompare(a.getName()));

  for (let s of arsipSheets) {
    let dataArsip = s.getDataRange().getValues();
    let found = false;
    // Cari dari baris paling bawah (transaksi paling terakhir di arsip)
    for (let i = dataArsip.length - 1; i >= 1; i--) {
      let fleetArsip = String(dataArsip[i][10]).replace(/^'/, "").trim();
      if (fleetArsip === nomorFleet) {
        saldoAwal = parseFloat(dataArsip[i][18]) || 0; // Ambil nilai Kolom S terakhir
        found = true;
        break;
      }
    }
    if (found) break; // Jika ketemu, hentikan pencarian
  }

  // 2. Tarik Data Sheet Aktif Bulan Ini
  const sheetAktif = ss.getSheetByName("Izin_Kendaraan");
  const dataAktif = sheetAktif.getDataRange().getValues();

  // Siapkan array baru khusus untuk update Kolom S (Sangat Cepat & Anti Lag)
  let kolomS_Baru = [];
  for (let i = 0; i < dataAktif.length; i++) {
    kolomS_Baru.push([dataAktif[i][18]]); // Kopi data Kolom S yang ada
  }

  let currentSaldo = saldoAwal;

  // 3. Kalkulasi Ulang Berantai dari Atas ke Bawah
  for (let i = 1; i < dataAktif.length; i++) {
    let rowFleet = String(dataAktif[i][10]).replace(/^'/, "").trim();
    if (rowFleet === nomorFleet) {
      let biaya = parseFloat(dataAktif[i][16]) || 0;
      let topup = parseFloat(dataAktif[i][17]) || 0;

      currentSaldo = currentSaldo + topup - biaya;
      kolomS_Baru[i][0] = currentSaldo; // Timpa dengan saldo mutasi baru
    }
  }

  // 4. Tulis balik SATU KOLOM FULL ke spreadsheet dalam 1 kedipan mata
  sheetAktif.getRange(1, 19, kolomS_Baru.length, 1).setValues(kolomS_Baru);
}

// --- FUNGSI UPDATE: TARIK DATA MASTER FLEET ---
function getListFleet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Fleet");
    if (!sheet) return []; // Pastikan nama sheet adalah Master_Fleet

    const data = sheet.getDataRange().getValues();
    let list = [];

    // Asumsi data kartu fleet ada di Kolom A, mulai baris ke-2
    for (let i = 1; i < data.length; i++) {
      let val = data[i][0];
      if (val && String(val).trim() !== "") {
        // Bersihkan tanda kutip (') di awal angka jika ada
        list.push(String(val).replace(/^'/, "").trim());
      }
    }
    return list;
  } catch (e) {
    return [];
  }
}

// ===================================================================================
// PERBAIKAN: RADAR GA REAL-TIME (MENDETEKSI TAMBAH BIAYA & STATUS)
// ===================================================================================
function cekNotifikasiIzinBaruGA(stateLokal) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    if (!sheet) return { hasChanged: false, newState: stateLokal };

    const data = sheet.getDataRange().getValues();
    const totalRows = data.length;

    // Bikin "State Signature" gabungan dari Total Baris, Total Biaya, dan Jumlah Pending
    let totalBiayaSeluruhnya = 0;
    let totalPending = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      totalBiayaSeluruhnya += parseFloat(data[i][16]) || 0; // Kolom Q (TOTAL_BIAYA)
      if (data[i][11] === "PENDING") totalPending++; // Kolom L (STATUS_GA)
    }

    const currentState = totalRows + "_" + totalBiayaSeluruhnya + "_" + totalPending;

    // Jika ada perubahan (User isi biaya, GA Approve, atau ada Izin baru masuk)
    if (currentState !== stateLokal) {
      // Sekalian tarik data terbarunya dari server agar Frontend tidak perlu memanggil 2 kali
      const freshData = getIzinKendaraanData();
      return { hasChanged: true, newState: currentState, data: freshData, totalBaris: totalRows };
    }

    return { hasChanged: false, newState: currentState };
  } catch (e) {
    return { hasChanged: false, newState: stateLokal };
  }
}

// ===================================================================================
// MODUL BARU: FUEL & COST ANOMALY DETECTOR (BACKEND)
// ===================================================================================

function simpanDataMasterFleet(kartuFleet, merk, tipe, bbm, batasBiaya) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Master_Fleet");

    // Safety check: Buat sheet otomatis jika belum ada atau terhapus
    if (!sheet) {
      sheet = ss.insertSheet("Master_Fleet");
      sheet.appendRow(["KARTU_FLEET", "MERK", "TIPE", "JENIS_BBM", "BATAS_BIAYA_PER_KM"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f8fafc");
    }

    const data = sheet.getDataRange().getValues();
    let isFound = false;
    let platAman = "'" + String(kartuFleet).trim(); // Cegah error format (E+) di Sheets

    // Cari apakah plat sudah terdaftar, jika ya, Update!
    for (let i = 1; i < data.length; i++) {
      let existingPlat = String(data[i][0]).replace(/^'/, "").trim();
      if (existingPlat === String(kartuFleet).trim()) {
        sheet.getRange(i + 1, 2).setValue(merk);
        sheet.getRange(i + 1, 3).setValue(tipe);
        sheet.getRange(i + 1, 4).setValue(bbm);
        sheet.getRange(i + 1, 5).setValue(batasBiaya);
        isFound = true;
        break;
      }
    }

    // Jika belum ada, tambahkan sebagai armada baru
    if (!isFound) {
      sheet.appendRow([platAman, merk, tipe, bbm, batasBiaya]);
    }
    return { success: true, pesan: "Spesifikasi kendaraan " + kartuFleet + " berhasil disimpan!" };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}

// Algoritma Cerdas Deteksi Anomali (Tahan Banting terhadap Sheet Kosong)
function getDataAnomaliRealtime() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Tarik dan Amankan Sheet Master_Fleet
    const sheetFleet = ss.getSheetByName("Master_Fleet");
    if (!sheetFleet) return []; // Jika sheet tidak ada, kembalikan kosong

    // Cerdas: Jika header (baris 1) cuma ada 1 kolom, kita buatkan sisa kolomnya otomatis
    const batasKolomFleet = sheetFleet.getLastColumn();
    if (batasKolomFleet < 5) {
      sheetFleet.getRange("A1:E1").setValues([["KARTU_FLEET", "MERK", "TIPE", "JENIS_BBM", "BATAS_BIAYA_PER_KM"]]);
      sheetFleet.getRange("A1:E1").setFontWeight("bold").setBackground("#f8fafc");
    }

    const dataFleet = sheetFleet.getDataRange().getValues();
    let mapBatasBiaya = {};

    // Looping data fleet (mulai dari baris 2)
    for (let i = 1; i < dataFleet.length; i++) {
      if (!dataFleet[i][0]) continue; // Lewati jika plat kosong
      let plat = String(dataFleet[i][0]).replace(/^'/, "").trim();
      // Pastikan kolom ke-5 (index 4) ada datanya
      let batas = dataFleet[i][4] ? parseFloat(dataFleet[i][4]) : 0;

      if (plat && batas > 0) {
        mapBatasBiaya[plat] = batas; // Menyimpan batas Rp/KM per plat
      }
    }

    // 2. Tarik Data Izin GA
    const sheetIzin = ss.getSheetByName("Izin_Kendaraan");
    if (!sheetIzin) return [];
    const dataIzin = sheetIzin.getDataRange().getValues();

    let hasilAnomali = [];

    // Jika data izin masih kosong (cuma ada header), langsung laporkan "Aman!"
    if (dataIzin.length <= 1) return [];

    let kmTracker = {};
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

    // 3. Looping untuk menghitung selisih KM
    for (let i = 1; i < dataIzin.length; i++) {
      let row = dataIzin[i];
      if (!row[0]) continue;

      let tglObj = new Date(row[1]);
      let plat = String(row[2]).replace(/^'/, "").trim();
      let user = row[4];
      let kmSekarang = parseFloat(row[9]) || 0;
      let biayaBbm = parseFloat(row[12]) || 0; // HANYA MENGAMBIL BIAYA BBM (Kolom M / Index 12)

      if (kmTracker[plat] !== undefined) {
        let jarakTempuh = kmSekarang - kmTracker[plat].lastKm;

        if (jarakTempuh > 0 && mapBatasBiaya[plat]) {
          let prevBiaya = kmTracker[plat].lastBiaya; // Ini sekarang hanya berisi BBM
          if (prevBiaya > 0) {
            let rasio = prevBiaya / jarakTempuh;
            let batasWajar = mapBatasBiaya[plat];

            if (rasio > batasWajar) {
              let prevTglObj = kmTracker[plat].lastTglObj;
              if (prevTglObj.getMonth() === currentMonth && prevTglObj.getFullYear() === currentYear) {
                hasilAnomali.push({
                  tgl: Utilities.formatDate(prevTglObj, "GMT+7", "dd/MM/yyyy"),
                  user: kmTracker[plat].lastUser,
                  fleet: plat,
                  rasioAktual: Math.round(rasio).toLocaleString("id-ID"),
                  batasWajar: Math.round(batasWajar).toLocaleString("id-ID"),
                });
              }
            }
          }
        }
      }

      kmTracker[plat] = {
        lastKm: kmSekarang,
        lastBiaya: biayaBbm, // <-- SIMPAN HANYA BIAYA BBM SAJA
        lastUser: user,
        lastTglObj: tglObj,
      };
    }

    return hasilAnomali.reverse();
  } catch (e) {
    return [];
  }
}
