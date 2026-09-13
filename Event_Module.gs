const EventModule = {
  getActiveEvent() {
    const cached = CoreCache.get("ACTIVE_EVENT");
    if (cached) {
      if (cached.status === "SELESAI") return null;
      return cached;
    }

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.DATABASE.MASTER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;

    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      const st = String(rows[i][7] || "").toUpperCase();
      if (st === "AKTIF" || st === "TUTUP_FORM") {
        const meta = {
          eventId: rows[i][0],
          title: rows[i][1],
          date: rows[i][2],
          time: rows[i][3],
          location: rows[i][4],
          formId: rows[i][5],
          formUrl: rows[i][6],
          status: rows[i][7]
        };
        CoreCache.put("ACTIVE_EVENT", meta);
        return meta;
      }
    }
    return null;
  },

  createEvent(payload) {
    return CoreLock.execute(() => {
      const ss = getDbSpreadsheet();
      let master = ss.getSheetByName(CONFIG.DATABASE.MASTER_SHEET);
      if (!master) {
        master = ss.insertSheet(CONFIG.DATABASE.MASTER_SHEET);
        master.appendRow(["Event ID", "Nama Acara", "Tanggal", "Waktu", "Lokasi", "Form ID", "Form URL", "Status"]);
      }

      const newId = "EVT-" + Utilities.formatDate(new Date(), CONFIG.TIMEZONE || "Asia/Jakarta", "yyyyMMdd-HHmmss");
      const formDesc = `Silakan isi formulir registrasi.\n\nTanggal: ${payload.date}\nTempat: ${payload.location}\nWaktu: ${payload.time}`;

      const form = FormApp.create(payload.title);
      form.setDescription(formDesc)
          .setAllowResponseEdits(false)
          .setCollectEmail(false);

      form.addTextItem().setTitle("Email").setRequired(true);
      form.addTextItem().setTitle("Nama Lengkap").setRequired(true);
      form.addTextItem().setTitle("Instansi").setRequired(true);
      form.addTextItem().setTitle("Nomor WhatsApp").setRequired(true);

      const formUrl = form.getPublishedUrl();

      // Nonaktifkan event lama
      const lastRow = master.getLastRow();
      if (lastRow > 1) {
        const statuses = master.getRange(2, 8, lastRow - 1, 1).getValues();
        for (let i = 0; i < statuses.length; i++) {
          if (String(statuses[i][0]).toUpperCase() === "AKTIF") {
            statuses[i][0] = "SELESAI";
          }
        }
        master.getRange(2, 8, lastRow - 1, 1).setValues(statuses);
      }

      master.appendRow([
        newId,
        payload.title,
        payload.date,
        payload.time,
        payload.location,
        form.getId(),
        formUrl,
        "AKTIF"
      ]);

      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      CoreCache.remove("ACTIVE_EVENT");

      return {
        status: "success",
        eventId: newId,
        formUrl: formUrl
      };
    });
  },

 toggleFormAccepting(isOpen) {
    return CoreLock.execute(() => {
      const active = this.getActiveEvent();
      if (!active || !active.eventId) {
        return { status: "error", message: "Tidak ada event yang sedang aktif." };
      }

      const ss = getDbSpreadsheet();
      const master = ss.getSheetByName(CONFIG.DATABASE.MASTER_SHEET);
      if (!master || master.getLastRow() < 2) {
        return { status: "error", message: "Sheet Master_Events tidak ditemukan." };
      }

      const data = master.getDataRange().getValues();
      const headers = data[0];

      // Deteksi dinamis indeks kolom
      let formIdIdx = headers.findIndex(h => String(h).toLowerCase().includes("form id"));
      let statusIdx = headers.findIndex(h => String(h).toLowerCase().includes("status"));

      if (formIdIdx === -1) formIdIdx = 5; 
      if (statusIdx === -1) statusIdx = headers.length - 1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(active.eventId).trim()) {
          const formId = String(data[i][formIdIdx] || "").trim();
          const targetStatus = isOpen ? "AKTIF" : "TUTUP_FORM";

          if (!formId) {
            return { status: "error", message: "Form ID tidak ditemukan pada baris event ini." };
          }

          // 1. Eksekusi pembaruan status Google Form secara bertahap & aman
          try {
            const form = FormApp.openById(formId);
            
            if (isOpen) {
              // Jika DIBUKA: Cukup nyalakan penerimaan tanggapan
              form.setAcceptingResponses(true);
            } else {
              // Jika DITUTUP: Pasang pesan penutup terlebih dahulu, baru matikan tanggapan
              const closeMsg = "Mohon maaf, pendaftaran untuk kegiatan ini telah ditutup karena kuota telah terpenuhi.";
              try {
                form.setCustomClosedFormMessage(closeMsg);
              } catch (msgErr) {
                Logger.log("Peringatan setCustomClosedFormMessage diabaikan: " + msgErr.message);
              }
              form.setAcceptingResponses(false);
            }
          } catch (formErr) {
            Logger.log("Gagal memodifikasi status form: " + formErr.message);
            return { status: "error", message: "Gagal menghubungkan ke Google Form: " + formErr.message };
          }

          // 2. Perbarui status di sheet Master_Events
          master.getRange(i + 1, statusIdx + 1).setValue(targetStatus);

          // 3. Bersihkan cache aktif agar dashboard membaca status baru
          CoreCache.remove("ACTIVE_EVENT");

          return {
            status: "success",
            isOpen: Boolean(isOpen),
            newStatus: targetStatus
          };
        }
      }

      return { status: "error", message: "Event aktif tidak cocok dengan database master." };
    });
  },

  archiveCurrentEvent() {
    return CoreLock.execute(() => {
      const ss = getDbSpreadsheet();
      const master = ss.getSheetByName(CONFIG.DATABASE.MASTER_SHEET);
      if (!master || master.getLastRow() < 2) {
        return { status: "error", message: "Tidak ada event di master." };
      }

      // 1. Ubah status event berjalan menjadi SELESAI
      const rows = master.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const st = String(rows[i][7] || "").toUpperCase();
        if (st === "AKTIF" || st === "TUTUP_FORM") {
          master.getRange(i + 1, 8).setValue("SELESAI");
          try {
            if (rows[i][5]) {
              FormApp.openById(rows[i][5]).setAcceptingResponses(false);
            }
          } catch (e) {
            Logger.log("Gagal mengunci form: " + e.message);
          }
        }
      }

      // 2. Kosongkan isi Presensi_Master otomatis (sisakan baris header)
      const presensiSheet = ss.getSheetByName(CONFIG.DATABASE.DATA_SHEET);
      if (presensiSheet && presensiSheet.getLastRow() > 1) {
        presensiSheet.deleteRows(2, presensiSheet.getLastRow() - 1);
      }

      // 3. Bersihkan cache
      CoreCache.remove("ACTIVE_EVENT");

      return { status: "success", message: "Event berhasil diarsipkan." };
    });
  }
};
