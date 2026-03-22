// ===================================================================================
// FILE: 05_Cetak_PDF.gs
// DESKRIPSI: Generator Dokumen HTML to PDF Base64
// ===================================================================================

function generatePDFBase64(nrpp, st) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();

    let namaUser = "",
      jabatanUser = "";
    for (let k = 1; k < dataKaryawan.length; k++) {
      if (String(dataKaryawan[k][0]) === String(nrpp)) {
        namaUser = dataKaryawan[k][1];
        jabatanUser = dataKaryawan[k][2];
        break;
      }
    }

    let trips = [];
    let customer = "",
      lokasi = "",
      plat = "";
    let minTglKeluar = null,
      maxTglMasuk = null;
    let grandTotal = 0;

    const allSheets = ss.getSheets();
    let dataKumpulan = [];
    allSheets.forEach((sheet) => {
      let sheetName = sheet.getName();
      if (sheetName === "Log_Perjalanan" || sheetName.startsWith("Arsip_")) {
        let sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) dataKumpulan.push(sheetData[i]);
      }
    });

    for (let i = 0; i < dataKumpulan.length; i++) {
      if (!dataKumpulan[i][0]) continue;

      if (String(dataKumpulan[i][3]) === String(nrpp) && formatST(dataKumpulan[i][5]) === formatST(st) && dataKumpulan[i][12] === "IN") {
        if (!customer) customer = dataKumpulan[i][8];
        if (!lokasi) lokasi = dataKumpulan[i][9];
        if (!plat) plat = dataKumpulan[i][6] + " - " + dataKumpulan[i][7];

        let tglKeluar = new Date(dataKumpulan[i][1]);
        let tglMasuk = new Date(dataKumpulan[i][2]);

        // PERBAIKAN: Gunakan parseDurasi
        let durasiJam = parseDurasi(dataKumpulan[i][10]);

        if (!minTglKeluar || tglKeluar < minTglKeluar) minTglKeluar = tglKeluar;
        if (!maxTglMasuk || tglMasuk > maxTglMasuk) maxTglMasuk = tglMasuk;

        let rincian = kalkulasiUPD(jabatanUser, tglKeluar, durasiJam);
        grandTotal += rincian.total;

        trips.push({
          tanggalStr: formatTanggalIndo(tglKeluar),
          upd_tlm: rincian.upd_tlm,
          makan_siang: rincian.uang_makan_siang,
          makan_total: rincian.uang_makan,
          lain_lain: rincian.lain_lain,
          total: rincian.total,
        });
      }
    }

    // ====================================================================
    // PERBAIKAN: Format ST Mengikuti Bulan & Tahun Keberangkatan
    // ====================================================================
    const bulanRomawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];

    // Gunakan tanggal keberangkatan (minTglKeluar). Jika kosong, baru gunakan tanggal cetak
    const tglReferensi = minTglKeluar ? minTglKeluar : new Date();

    const blnRomawi_Str = bulanRomawi[tglReferensi.getMonth()]; // Mengambil bulan dari tanggal berangkat
    const thn_Str = tglReferensi.getFullYear(); // Mengambil tahun dari tanggal berangkat

    const formatSTLengkap = `${formatST(st)}/JKT/ST-PS/${blnRomawi_Str}/${thn_Str}`;
    // ====================================================================

    const template = HtmlService.createTemplateFromFile("Template_PDF");
    template.data = {
      st_asli: st,
      st_format: formatSTLengkap,
      nama: namaUser,
      jabatan: "",
      customer: customer,
      lokasi: lokasi,
      plat: "Kendaraan Dinas",
      berangkat: minTglKeluar ? Utilities.formatDate(minTglKeluar, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
      kembali: maxTglMasuk ? Utilities.formatDate(maxTglMasuk, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
      trips: trips,
      grandTotal: grandTotal,
    };

    const htmlOutput = template.evaluate();
    const blob = Utilities.newBlob(htmlOutput.getContent(), "text/html", "Laporan_" + st + ".html").getAs("application/pdf");

    return { success: true, base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (error) {
    return { success: false, pesan: error.toString() };
  }
}

function generateLogPDFBase64(nrpp, st, periodKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataKaryawan = ss.getSheetByName("Master_Karyawan").getDataRange().getValues();

    let namaUser = "Karyawan Tidak Ditemukan";
    for (let k = 1; k < dataKaryawan.length; k++) {
      if (String(dataKaryawan[k][0]) === String(nrpp)) {
        namaUser = dataKaryawan[k][1];
        break;
      }
    }

    let targetData = [];
    let teksPeriode = "";

    if (!periodKey || periodKey === "CURRENT") {
      targetData = getRiwayatUser(nrpp);
      teksPeriode = "BULAN INI";
    } else {
      const arsipGrouped = getRiwayatArsipGrouped(nrpp);
      targetData = arsipGrouped[periodKey] || [];
      teksPeriode = "ARSIP PERIODE " + periodKey;
    }

    const dataRiwayatST = st === "ALL" ? targetData : targetData.filter((item) => item.st === st);
    const displayST = st === "ALL" ? "SEMUA SURAT TUGAS (" + teksPeriode + ")" : st;

    const template = HtmlService.createTemplateFromFile("Template_Log_PDF");
    template.data = {
      nrpp: nrpp,
      nama: namaUser,
      st: displayST,
      tglCetak: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"),
      log: dataRiwayatST,
    };

    const htmlOutput = template.evaluate();
    const fileName = st === "ALL" ? "Lampiran_Semua_Log_" + nrpp + ".html" : "Lampiran_Log_" + st + "_" + nrpp + ".html";
    const blob = Utilities.newBlob(htmlOutput.getContent(), "text/html", fileName).getAs("application/pdf");

    return { success: true, base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (error) {
    return { success: false, pesan: error.toString() };
  }
}
