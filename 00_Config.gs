const CONFIG = Object.freeze({
  DATABASE: {
    // Tempelkan ID Spreadsheet kamu di sini:
    SPREADSHEET_ID: "1kVQjMrwEKfvfqGsSZrJbCQ3pWuI3FkLIEyHeyzHRQ5Y",
    MASTER_SHEET: "Master_Events",
    DATA_SHEET: "Presensi_Master"
  },
  CACHE_TTL: 120, // 2 menit
  LOCK_TIMEOUT: 4000, // 4 detik
  TIMEZONE: "Asia/Jakarta"
});

// Helper agar tidak null baik di Standalone Script maupun Container-bound
function getDatabaseSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && CONFIG.DATABASE.SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(CONFIG.DATABASE.SPREADSHEET_ID);
  }
  if (!ss) {
    throw new Error("Spreadsheet database tidak ditemukan. Pastikan SPREADSHEET_ID sudah diisi di 00_Config.gs.");
  }
  return ss;
}

// Alias jika di modul lain memanggil getDbSpreadsheet()
function getDbSpreadsheet() {
  return getDatabaseSpreadsheet();
}
