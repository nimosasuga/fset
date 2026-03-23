// ===================================================================================
// MESIN CETAK PDF ANALYTICS & ANOMALI GA
// ===================================================================================
function cetakLaporanGAPDFServer(payload) {
  try {
    // Memanggil file template HTML khusus PDF
    let htmlTemplate = HtmlService.createTemplateFromFile('Template_Laporan_GA');
    htmlTemplate.data = payload; // Masukkan semua data (Total, Gambar Grafik, Anomali) ke template

    // Evaluasi HTML menjadi string
    let htmlEvaluated = htmlTemplate.evaluate().getContent();
    
    // Konversi HTML menjadi file PDF
    let blob = Utilities.newBlob(htmlEvaluated, MimeType.HTML, "Laporan_Sementara.html");
    let pdfBlob = blob.getAs(MimeType.PDF).setName("Laporan_Analytics_GA_" + payload.waktuCetakKotor + ".pdf");

    // Ubah PDF menjadi Base64 agar bisa langsung di-download otomatis oleh Frontend
    let base64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    return { success: true, fileName: pdfBlob.getName(), base64: base64 };
  } catch (e) {
    return { success: false, pesan: e.toString() };
  }
}