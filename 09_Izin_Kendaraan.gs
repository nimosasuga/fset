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

function getIzinKendaraanData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Izin_Kendaraan");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let fleetBalances = {};

    // 1. Tarik Saldo Terakhir dari Brankas Arsip (Agar mutasi All-Time akurat)
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

    // 2. Kalkulasi Mutasi Bulan Ini (Dari Baris Atas / Terlama ke Bawah / Terbaru)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Lewati jika ID kosong

      let f = String(data[i][10]).replace(/^'/, "").trim();
      let biaya = parseFloat(data[i][16]) || 0;
      let topup = parseFloat(data[i][17]) || 0;
      let currentSaldo = 0;

      // KUNCI MUTASI: Tambahkan Top Up dan Kurangi Biaya ke dalam memori Saldo Kartu
      if (f && f !== "-") {
        fleetBalances[f] = (fleetBalances[f] || 0) + topup - biaya;
        currentSaldo = fleetBalances[f];
      } else {
        currentSaldo = topup - biaya;
      }

      // Format Tanggal yang Aman
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
        km: data[i][9] || "-",
        fleet: f || "-",
        status: data[i][11],
        totalBiaya: biaya,
        topup: topup,
        saldo: currentSaldo, // <- SEKARANG DATA INI DIKIRIM KE FRONTEND
      });
    }

    // 3. Balik urutan Array agar transaksi TERBARU muncul di PALING ATAS tabel GA
    return result.reverse();
  } catch (e) {
    return [];
  }
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

// --- FUNGSI UPDATE: RIWAYAT USER (HANYA BULAN BERJALAN & SUPER CEPAT) ---
function getRiwayatIzinKendaraanUser(nrpp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Kita tarik batas 150 data terakhir agar aman mencakup 1 bulan penuh
  const limit = 150;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, 19).getValues();
  let result = [];

  // LOGIKA BARU: Dapatkan bulan dan tahun saat ini (Real-time)
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][3]) === String(nrpp)) {
      // Mengubah teks tanggal dari database menjadi Objek Tanggal
      let tglPengajuan = new Date(data[i][1]);

      // VERIFIKASI: Hanya kirim data ke frontend JIKA bulannya = bulan ini & tahunnya = tahun ini
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
        });
      }
    }
  }
  return result;
}

// --- FUNGSI UPDATE: SIMPAN PENGELUARAN KENDARAAN (MENGHITUNG SALDO OTOMATIS) ---
function simpanPengeluaranKendaraan(idIzin, biaya) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();

    // Pastikan Header SISA_SALDO (S) dibuat
    if (data[0].length < 19) {
      sheet.getRange("M1:S1").setValues([["BBM", "TOL", "PARKIR", "LAIN_LAIN", "TOTAL_BIAYA", "TOP_UP_FLEET", "SISA_SALDO"]]);
      sheet.getRange("M1:S1").setFontWeight("bold").setBackground("#f8fafc");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idIzin) {
        let valBbm = Number(biaya.bbm) || 0;
        let valTol = Number(biaya.tol) || 0;
        let valParkir = Number(biaya.parkir) || 0;
        let valLain = Number(biaya.lain) || 0;

        let total = valBbm + valTol + valParkir + valLain;
        let valTopup = Number(data[i][17]) || 0; // Mengambil nominal Top Up saat ini (jika ada)
        let sisaSaldo = valTopup - total; // Otomatis mengkalkulasi saldo

        sheet.getRange(i + 1, 13).setValue(valBbm);
        sheet.getRange(i + 1, 14).setValue(valTol);
        sheet.getRange(i + 1, 15).setValue(valParkir);
        sheet.getRange(i + 1, 16).setValue(valLain);
        sheet.getRange(i + 1, 17).setValue(total);
        sheet.getRange(i + 1, 19).setValue(sisaSaldo); // Menulis ke Kolom S (19)

        return { success: true, pesan: "Pengeluaran berhasil disimpan!" };
      }
    }
    return { success: false, pesan: "Data Izin tidak ditemukan." };
  } catch (error) {
    return { success: false, pesan: error.toString() };
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

// --- UPDATE FUNGSI KOREKSI GA LENGKAP (Pemicu Mutasi) ---
function updateKoreksiIzinGA(idIzin, dataBaru) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    const data = sheet.getDataRange().getValues();
    let fleetTarget = "";

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
        sheet.getRange(row, 10).setValue(dataBaru.km);
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
        break;
      }
    }

    // Picu perhitungan mutasi berantai ke bawah
    if (fleetTarget && fleetTarget !== "-") {
      recalculateLedgerFleet(fleetTarget);
    }
    return { success: true, pesan: "Data terkoreksi & Mutasi Saldo berantai diperbarui!" };
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
// FUNGSI BARU: RADAR NOTIFIKASI REAL-TIME (BACKGROUND POLLING)
// ===================================================================================

function cekNotifikasiIzinBaruGA(jumlahDataLokal) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izin_Kendaraan");
    if (!sheet) return { hasNew: false, totalBaru: jumlahDataLokal };

    // Hitung total baris saat ini (dikurangi 1 baris header)
    const totalDataServer = sheet.getLastRow() - 1;

    // Jika data di server lebih banyak dari data di layar GA, berarti ada yang baru!
    if (totalDataServer > jumlahDataLokal) {
      return { hasNew: true, totalBaru: totalDataServer };
    }

    return { hasNew: false, totalBaru: totalDataServer };
  } catch (e) {
    return { hasNew: false, totalBaru: jumlahDataLokal };
  }
}
