// ===================================================================================
// FILE: 01_Config.gs
// DESKRIPSI: Variabel Global, Konfigurasi Sistem, Router, dan Fungsi Helper
// ===================================================================================

const POS_LAT = -6.2706266;
const POS_LNG = 107.1962988;
const MAX_RADIUS = 20;
const KODE_QR_VALID = "POS-SECURITY-FSET-01";
const GOD_MODE_NRPP = "150623";

const LIBUR_NASIONAL = ["2026-01-01", "2026-02-17", "2026-03-03", "2026-03-20", "2026-04-10", "2026-05-01", "2026-05-14", "2026-06-01", "2026-08-17", "2026-12-25"];

function doGet() {
  return HtmlService.createTemplateFromFile("Index").evaluate().setTitle("FSET - Dashboard Karyawan").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function isHariLibur(dateObj) {
  const day = dateObj.getDay();
  if (day === 0 || day === 6) return true;
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, "0");
  const dateNum = String(dateObj.getDate()).padStart(2, "0");
  const formattedDate = `${year}-${month}-${dateNum}`;
  return LIBUR_NASIONAL.includes(formattedDate);
}

function formatST(stValue) {
  let st = String(stValue).trim();
  if (!st || st === "undefined" || st === "null" || st === "") return "TANPA-ST";
  if (/^\d+$/.test(st) && st.length < 4) {
    return st.padStart(4, "0");
  }
  return st;
}

function formatTanggalIndo(dateObj) {
  const hari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return `${hari[dateObj.getDay()]}, ${String(dateObj.getDate()).padStart(2, "0")} ${bulan[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
}

/** PERBAIKAN: Penyelamat data dari Bug Waktu Google Sheets */
function parseDurasi(val) {
  if (val === "" || val === null || val === undefined) return 0;

  // Jika Google Sheets diam-diam merubahnya jadi Waktu (Date)
  if (val instanceof Date) {
    let jam = Utilities.formatDate(val, "GMT+7", "HH");
    let menit = Utilities.formatDate(val, "GMT+7", "mm");
    return parseFloat(parseInt(jam, 10) + "." + menit);
  }

  if (typeof val === "string") {
    if (val.includes(":")) {
      let parts = val.split(":");
      return parseFloat(parseInt(parts[0], 10) + "." + parts[1]);
    }
    val = val.replace(",", ".");
  }
  return parseFloat(val) || 0;
}

/** PERBAIKAN: Penyelamat jika format tanggal di Sheets berubah jadi Text */
function parseSafeDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === "string") {
    let regex = /^(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})$/;
    let match = val.match(regex);
    if (match) {
      return new Date(match[3], parseInt(match[2]) - 1, match[1], match[4], match[5], match[6]);
    }
    return new Date(val);
  }
  return new Date();
}
