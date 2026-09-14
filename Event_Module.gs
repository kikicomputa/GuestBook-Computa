// File: Event_Module.gs

const EventModule = {
  /**
   * Mengambil metadata event yang berstatus AKTIF atau TUTUP_FORM
   */
  getActiveEvent() {
    try {
      if (typeof CoreCache !== "undefined" && CoreCache.get) {
        const cached = CoreCache.get("ACTIVE_EVENT");
        if (cached) {
          if (cached.status === "SELESAI") return null;
          return cached;
        }
      }

      const ss = (typeof getDbSpreadsheet === "function") ? getDbSpreadsheet() : ((typeof getDatabaseSpreadsheet === "function") ? getDatabaseSpreadsheet() : SpreadsheetApp.getActiveSpreadsheet());
      const sheetName = (typeof CONFIG !== "undefined" && CONFIG.DATABASE && CONFIG.DATABASE.MASTER_SHEET) ? CONFIG.DATABASE.MASTER_SHEET : "Master_Events";
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return null;

      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        const st = String(rows[i][7] || "").toUpperCase();
        if (st === "AKTIF" || st === "TUTUP_FORM") {
          const meta = {
            eventId: String(rows[i][0] || ""),
            title: String(rows[i][1] || "Acara Tanpa Judul"),
            date: String(rows[i][2] || "-"),
            time: String(rows[i][3] || "-"),
            location: String(rows[i][4] || "-"),
            formId: String(rows[i][5] || ""),
            formUrl: String(rows[i][6] || ""),
            status: st
          };

          if (typeof CoreCache !== "undefined" && CoreCache.put) {
            CoreCache.put("ACTIVE_EVENT", meta);
          }
          return meta;
        }
      }
      return null;
    } catch (e) {
      Logger.log("getActiveEvent Error: " + e.message);
      return null;
    }
  },

  /**
   * Membuat Google Form, mendaftarkan ke Master_Events, dan mengembalikan URL
   */
  createEvent(payload) {
    const runner = () => {
      if (!payload) {
        return { status: "error", message: "Data formulir event tidak diterima." };
      }

      const ss = (typeof getDbSpreadsheet === "function") ? getDbSpreadsheet() : ((typeof getDatabaseSpreadsheet === "function") ? getDatabaseSpreadsheet() : SpreadsheetApp.getActiveSpreadsheet());
      const sheetName = (typeof CONFIG !== "undefined" && CONFIG.DATABASE && CONFIG.DATABASE.MASTER_SHEET) ? CONFIG.DATABASE.MASTER_SHEET : "Master_Events";
      let master = ss.getSheetByName(sheetName);

      if (!master) {
        master = ss.insertSheet(sheetName);
        master.appendRow(["Event ID", "Nama Acara", "Tanggal", "Waktu", "Lokasi", "Form ID", "Form URL", "Status"]);
      }

      const tz = (typeof CONFIG !== "undefined" && CONFIG.TIMEZONE) ? CONFIG.TIMEZONE : "Asia/Jakarta";
      const newId = "EVT-" + Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
      
      const title = String(payload.title || "Acara Baru").trim();
      const date = String(payload.date || "-").trim();
      const time = String(payload.time || "-").trim();
      const location = String(payload.location || "-").trim();

      const formDesc = "Silakan isi formulir registrasi.\n\nTanggal: " + date + "\nTempat: " + location + "\nWaktu: " + time;

      // 1. Buat Google Form
      const form = FormApp.create(title);
      form.setDescription(formDesc)
          .setAllowResponseEdits(false)
          .setCollectEmail(false);

      form.addTextItem().setTitle("Email").setRequired(true);
      form.addTextItem().setTitle("Nama Lengkap").setRequired(true);
      form.addTextItem().setTitle("Instansi").setRequired(true);
      form.addTextItem().setTitle("Nomor WhatsApp").setRequired(true);

      const formUrl = form.getPublishedUrl();
      const formId = form.getId();

      // 2. Nonaktifkan status event aktif yang lama
      const lastRow = master.getLastRow();
      if (lastRow > 1) {
        const statuses = master.getRange(2, 8, lastRow - 1, 1).getValues();
        for (let i = 0; i < statuses.length; i++) {
          if (String(statuses[i][0]).toUpperCase() === "AKTIF" || String(statuses[i][0]).toUpperCase() === "TUTUP_FORM") {
            statuses[i][0] = "SELESAI";
          }
        }
        master.getRange(2, 8, lastRow - 1, 1).setValues(statuses);
      }

      // 3. Rekam event baru ke Master_Events (8 kolom)
      master.appendRow([
        newId,
        title,
        date,
        time,
        location,
        formId,
        formUrl,
        "AKTIF"
      ]);

      // Hubungkan formulir ke spreadsheet
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

      if (typeof CoreCache !== "undefined" && CoreCache.remove) {
        CoreCache.remove("ACTIVE_EVENT");
      }

      return {
        status: "success",
        eventId: newId,
        formUrl: formUrl
      };
    };

    if (typeof CoreLock !== "undefined" && CoreLock.execute) {
      return CoreLock.execute(runner);
    }
    return runner();
  },

  /**
   * Buka / Tutup penerimaan respon formulir
   */
  // File: Event_Module.gs

  toggleFormAccepting(isOpen) {
    const runner = () => {
      const ss = (typeof getDatabaseSpreadsheet === "function") 
        ? getDatabaseSpreadsheet() 
        : SpreadsheetApp.getActiveSpreadsheet();

      const sheetName = (typeof CONFIG !== "undefined" && CONFIG.DATABASE && CONFIG.DATABASE.MASTER_SHEET) 
        ? CONFIG.DATABASE.MASTER_SHEET 
        : "Master_Events";

      const master = ss.getSheetByName(sheetName);
      if (!master || master.getLastRow() < 2) {
        return { status: "error", message: "Sheet Master_Events tidak ditemukan." };
      }

      const data = master.getDataRange().getValues();
      let targetRowIndex = -1;

      // Cari event aktif atau yang sedang di-off-kan sementara (bukan SELESAI)
      for (let i = data.length - 1; i >= 1; i--) {
        const st = String(data[i][7] || data[i][8] || "").trim().toUpperCase();
        if (st === "AKTIF" || st === "TUTUP_FORM" || st === "NONAKTIF") {
          targetRowIndex = i + 1;
          break;
        }
      }

      if (targetRowIndex === -1) {
        return { status: "error", message: "Tidak ada event aktif yang bisa diubah statusnya." };
      }

      const formId = String(data[targetRowIndex - 1][5] || data[targetRowIndex - 1][6] || "").trim();
      const newStatus = isOpen ? "AKTIF" : "TUTUP_FORM";

      // 1. Ubah setelan Google Form langsung
      if (formId) {
        try {
          const form = FormApp.openById(formId);
          form.setAcceptingResponses(Boolean(isOpen));
          if (!isOpen) {
            try {
              form.setCustomClosedFormMessage("Pendaftaran untuk acara ini ditutup sementara waktu.");
            } catch (e) {}
          }
        } catch (err) {
          Logger.log("Gagal sinkron form: " + err.message);
        }
      }

      // 2. Perbarui status di kolom Master_Events
      const colStatus = data[0].length; // Kolom Status terakhir
      master.getRange(targetRowIndex, colStatus).setValue(newStatus);

      // 3. Bersihkan cache
      if (typeof CoreCache !== "undefined" && CoreCache.remove) {
        CoreCache.remove("ACTIVE_EVENT");
      }

      return {
        status: "success",
        isOpen: Boolean(isOpen),
        newStatus: newStatus
      };
    };

    if (typeof CoreLock !== "undefined" && CoreLock.execute) {
      return CoreLock.execute(runner);
    }
    return runner();
  },

  /**
   * Arsipkan event yang sedang berjalan
   */
  archiveCurrentEvent() {
    const runner = () => {
      const ss = (typeof getDbSpreadsheet === "function") ? getDbSpreadsheet() : ((typeof getDatabaseSpreadsheet === "function") ? getDatabaseSpreadsheet() : SpreadsheetApp.getActiveSpreadsheet());
      const sheetName = (typeof CONFIG !== "undefined" && CONFIG.DATABASE && CONFIG.DATABASE.MASTER_SHEET) ? CONFIG.DATABASE.MASTER_SHEET : "Master_Events";
      const master = ss.getSheetByName(sheetName);
      if (!master || master.getLastRow() < 2) {
        return { status: "error", message: "Tidak ada event di master." };
      }

      const rows = master.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const st = String(rows[i][7] || "").toUpperCase();
        if (st === "AKTIF" || st === "TUTUP_FORM") {
          master.getRange(i + 1, 8).setValue("SELESAI");
          try {
            if (rows[i][5]) {
              FormApp.openById(rows[i][5]).setAcceptingResponses(false);
            }
          } catch (e) {}
        }
      }

      const dataSheetName = (typeof CONFIG !== "undefined" && CONFIG.DATABASE && CONFIG.DATABASE.DATA_SHEET) ? CONFIG.DATABASE.DATA_SHEET : "Presensi_Master";
      const presensiSheet = ss.getSheetByName(dataSheetName);
      if (presensiSheet && presensiSheet.getLastRow() > 1) {
        presensiSheet.deleteRows(2, presensiSheet.getLastRow() - 1);
      }

      if (typeof CoreCache !== "undefined" && CoreCache.remove) {
        CoreCache.remove("ACTIVE_EVENT");
      }

      return { status: "success", message: "Event berhasil diarsipkan." };
    };

    if (typeof CoreLock !== "undefined" && CoreLock.execute) {
      return CoreLock.execute(runner);
    }
    return runner();
  }
};

// Fungsi Top-Level Wrapper agar fungsi dapat dipanggil langsung dari luar
function createEvent(payload) {
  return EventModule.createEvent(payload);
}
