// ===================================================================================
// FILE: 03_Absen_Payroll.gs
// DESKRIPSI: Logika Geofencing GPS, Scanner Barcode, dan Kalkulasi Uang Saku (UPD)
// ===================================================================================

function hitungJarakServer(lat, lng) {
  var R = 6371e3;
  var radLat1 = (POS_LAT * Math.PI) / 180;
  var radLat2 = (lat * Math.PI) / 180;
  var deltaLat = ((lat - POS_LAT) * Math.PI) / 180;
  var deltaLng = ((lng - POS_LNG) * Math.PI) / 180;
  var a = Math.sin(deltaLat / 2) * Math.sin(deltaLat / 2) + Math.cos(radLat1) * Math.cos(radLat2) * Math.sin(deltaLng / 2) * Math.sin(deltaLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function validasiLokasi(lat, lng, nrpp) {
  if (String(nrpp) === GOD_MODE_NRPP) return { valid: true, jarak: 0 };
  var jarak = hitungJarakServer(lat, lng);
  if (jarak > MAX_RADIUS) return { valid: false, pesan: "Jarak Anda " + Math.round(jarak) + "m dari Pos. Maks 20m." };
  return { valid: true, jarak: jarak };
}

function getListKendaraan() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Kendaraan").getDataRange().getValues();
  let list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== "") list.push({ plat: data[i][0], jenis: data[i][1] });
  }
  return list;
}

let GLOBAL_TARIF_MAP = null;

function kalkulasiUPD(jabatan, tglKeluar, durasiJam) {
  if (!GLOBAL_TARIF_MAP) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Tarif_UPD");
    GLOBAL_TARIF_MAP = {};
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          GLOBAL_TARIF_MAP[String(data[i][0]).toUpperCase()] = {
            upd_full: parseFloat(data[i][1]) || 0,
            upd_half: parseFloat(data[i][2]) || 0,
            makan: parseFloat(data[i][3]) || 0,
            makan_libur: parseFloat(data[i][4]) || 0,
            lain_kerja: parseFloat(data[i][5]) || 0,
            lain_libur: parseFloat(data[i][6]) || 0,
          };
        }
      }
    }
  }

  const libur = isHariLibur(tglKeluar);
  let jbtn = String(jabatan).toUpperCase();
  let t = GLOBAL_TARIF_MAP[jbtn];

  if (!t) {
    let key = Object.keys(GLOBAL_TARIF_MAP).find((k) => jbtn.includes(k));
    if (key) t = GLOBAL_TARIF_MAP[key];
  }

  if (!t) return { upd_tlm: 0, uang_makan: 0, uang_makan_siang: 0, lain_lain: 0, total: 0 };

  let upd_tlm = durasiJam >= 8 ? t.upd_full : t.upd_half;
  let uang_makan = t.makan;
  let uang_makan_siang = libur ? t.makan_libur : 0;
  let lain_lain = libur ? t.lain_libur : t.lain_kerja;
  let total = upd_tlm + uang_makan + uang_makan_siang + lain_lain;

  return { upd_tlm: upd_tlm, uang_makan: uang_makan, uang_makan_siang: uang_makan_siang, lain_lain: lain_lain, total: total };
}

function prosesDataScan(data) {
  const nrppString = String(data.nrpp);
  const teksQR = String(data.qr_text).trim();

  if (nrppString !== GOD_MODE_NRPP && teksQR !== KODE_QR_VALID) {
    return { success: false, pesan: "Barcode Palsu/Tidak Valid!<br><small class='text-muted'>Terbaca: '" + teksQR + "'</small>" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Log_Perjalanan");
  const dataLog = sheet.getDataRange().getValues();
  const waktuSekarang = new Date();

  if (data.tipe_scan === "KELUAR") {
    // =================================================================
    // BLOK VALIDASI MUTLAK: 1 HARI = 1 TRIP (TERMASUK GOD MODE)
    // =================================================================
    let apakahSudahDinasHariIni = cekBatasHarianDinas(nrppString);

    if (apakahSudahDinasHariIni) {
      // Tolak mentah-mentah jika sudah dinas hari ini
      return {
        success: false,
        pesan: "AKSES DITOLAK!<br>Anda sudah melakukan perjalanan dinas hari ini. Batas maksimal adalah 1 kali sehari.",
      };
    }
    // =================================================================

    const idTransaksi = "TRP-" + data.nrpp + "-" + Utilities.formatDate(waktuSekarang, "GMT+7", "yyMMddHHmm");
    const arrKendaraan = data.kendaraan.split("|");
    const stAman = "'" + data.st;

    sheet.appendRow([idTransaksi, waktuSekarang, "", data.nrpp, data.nama, stAman, arrKendaraan[0], arrKendaraan[1], data.customer, data.lokasi, "", "", "OUT", "", ""]);

    const barisBaru = sheet.getLastRow();
    sheet.getRange(barisBaru, 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(barisBaru, 6).setNumberFormat("@");

    return { success: true, pesan: nrppString === GOD_MODE_NRPP ? "Sukses Keluar! (Mode God)" : "Sukses Keluar! Hati-hati di jalan." };
  } else if (data.tipe_scan === "MASUK") {
    const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();
    let jabatanUser = "";

    for (let k = 1; k < dataKaryawan.length; k++) {
      if (String(dataKaryawan[k][0]) === nrppString) {
        jabatanUser = dataKaryawan[k][2];
        break;
      }
    }

    for (let i = dataLog.length - 1; i >= 1; i--) {
      if (String(dataLog[i][3]) === nrppString && dataLog[i][12] === "OUT") {
        const waktuKeluar = parseSafeDate(dataLog[i][1]);
        const durasiJam = (waktuSekarang.getTime() - waktuKeluar.getTime()) / (1000 * 60 * 60);
        const rincian = kalkulasiUPD(jabatanUser, waktuKeluar, durasiJam);
        const updTotal = rincian.total;
        const baris = i + 1;

        sheet.getRange(baris, 3).setValue(waktuSekarang);
        sheet.getRange(baris, 3).setNumberFormat("dd/MM/yyyy HH:mm:ss");

        // ====================================================================
        // PERBAIKAN: Memaksa format titik (.) dengan menjadikannya Plain Text
        // ====================================================================
        let durasiTeks = durasiJam.toFixed(2); // Menghasilkan format string dengan titik, misal "15.29"

        sheet.getRange(baris, 11).setNumberFormat("@"); // Set sel menjadi Plain Text murni terlebih dahulu
        sheet.getRange(baris, 11).setValue(durasiTeks); // Masukkan nilai teks ber-titik
        // ====================================================================

        sheet.getRange(baris, 12).setValue(updTotal);
        sheet.getRange(baris, 13).setValue("IN");
        sheet.getRange(baris, 15).setValue("PENDING");

        return { success: true, pesan: `Sukses Lapor Masuk!<br>Durasi: ${durasiJam.toFixed(2)} Jam<br>UPD: Rp ${updTotal.toLocaleString("id-ID")}` };
      }
    }
  }
}

// ===================================================================================
// FUNGSI SECURITY: VALIDASI BATAS MAKSIMAL 1X PERJALANAN PER HARI
// ===================================================================================
function cekBatasHarianDinas(nrpp) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log_Perjalanan");
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return false; // Belum ada data sama sekali, izinkan!

    const headers = data[0];
    const idxNRPP = headers.indexOf("NRPP");
    const idxKeluar = headers.indexOf("Waktu_Keluar");

    if (idxNRPP === -1 || idxKeluar === -1) return false;

    // Ambil tanggal hari ini (Format: YYYY-MM-DD)
    const hariIni = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

    // Looping dari bawah (data paling baru) ke atas
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idxNRPP]).trim() === String(nrpp).trim()) {
        let wKeluar = data[i][idxKeluar];

        // Pastikan format wKeluar valid dan merupakan tanggal
        if (wKeluar && wKeluar instanceof Date) {
          let tglKeluarLog = Utilities.formatDate(wKeluar, "GMT+7", "yyyy-MM-dd");

          // JIKA DIA SUDAH PERNAH SCAN KELUAR HARI INI...
          if (tglKeluarLog === hariIni) {
            return true; // TRUE = GEMBOK AKTIF (Tolak proses scan OUT selanjutnya!)
          }
        }

        // Karena kita hanya peduli pada histori perjalanan TERAKHIRNYA,
        // kita bisa langsung hentikan pencarian agar server tidak lelah.
        break;
      }
    }

    return false; // FALSE = GEMBOK TERBUKA (Belum pernah dinas hari ini)
  } catch (e) {
    return false; // Jika error, anggap lolos agar operasional tidak mati total
  }
}
