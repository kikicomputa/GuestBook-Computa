const AttendanceModule = {
  handleFormSubmit(e) {
    const active = EventModule.getActiveEvent();
    if (!active || active.status !== "AKTIF") {
      Logger.log("Peringatan: Tidak ada event berstatus AKTIF.");
      return;
    }

    const ss = getDbSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.DATABASE.DATA_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.DATABASE.DATA_SHEET);
      sheet.appendRow([
        "Timestamp", "Event ID", "Ticket ID", "Nama Lengkap", 
        "Instansi", "Email", "WhatsApp", "Status", "Waktu Masuk", "Gate"
      ]);
    }

    CoreLock.execute(() => {
      let timestamp, email, name, agency, wa;

      // 1. Prioritas: Ambil dari event e.values bawaan Form Submit
      if (e && e.values && e.values.length > 0) {
        timestamp = e.values[0] || new Date();
        email     = String(e.values[1] || "-").trim();
        name      = String(e.values[2] || "-").trim();
        agency    = String(e.values[3] || "-").trim();
        wa        = String(e.values[4] || "-").trim();
      } 
      // 2. Fallback: Ambil dari e.namedValues jika struktur object bernama
      else if (e && e.namedValues) {
        timestamp = e.namedValues["Timestamp"] ? e.namedValues["Timestamp"][0] : new Date();
        email     = e.namedValues["Email"] ? e.namedValues["Email"][0] : "-";
        name      = e.namedValues["Nama Lengkap"] ? e.namedValues["Nama Lengkap"][0] : "-";
        agency    = e.namedValues["Instansi"] ? e.namedValues["Instansi"][0] : "-";
        wa        = e.namedValues["Nomor WhatsApp"] ? e.namedValues["Nomor WhatsApp"][0] : "-";
      } 
      // 3. Fail-Safe: Jika dijalankan lewat tombol Run editor untuk pengetesan
      else {
        Logger.log("Menjalankan mode simulasi/manual (e undefined). Membaca baris terakhir sheet respon...");
        // Cek tab respon form bawaan
        const formSheet = ss.getSheets().find(s => s.getName().toLowerCase().includes("form") || s.getName().toLowerCase().includes("jawaban"));
        if (formSheet && formSheet.getLastRow() > 1) {
          const lastRowVals = formSheet.getRange(formSheet.getLastRow(), 1, 1, formSheet.getLastColumn()).getValues()[0];
          timestamp = lastRowVals[0];
          email     = lastRowVals[1];
          name      = lastRowVals[2];
          agency    = lastRowVals[3];
          wa        = lastRowVals[4];
        } else {
          // Data dummy pengujian jika sheet form belum ada baris sama sekali
          timestamp = new Date();
          email     = Session.getActiveUser().getEmail() || "test@domain.com";
          name      = "Tester Simulasi";
          agency    = "Computa Center";
          wa        = "08123456789";
        }
      }

      // Generate Kode Tiket Acak
      const ticketId = "TIK-" + Math.floor(1000 + Math.random() * 9000);

      // Masukkan ke tab Presensi_Master
      sheet.appendRow([
        timestamp,
        active.eventId,
        ticketId,
        name,
        agency,
        email,
        wa,
        "Belum",
        "-",
        "-"
      ]);

      // Kirim Tiket Digital QR ke Email Peserta
      if (email && email.indexOf("@") > -1) {
        try {
          this.sendEmail(email, name, agency, ticketId, active);
        } catch (mailErr) {
          Logger.log("Peringatan pengiriman email: " + mailErr.toString());
        }
      }
    });
  },

  sendEmail(to, name, agency, ticketId, meta) {
    const qr = "https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=" + encodeURIComponent(ticketId);
    const html = `
      <div style="font-family: Arial, sans-serif; max-width: 440px; margin: auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
        <div style="background: #0f172a; padding: 20px; text-align: center; color: #fff;">
          <h2 style="margin: 0; font-size: 18px;">E-Ticket Presensi</h2>
          <p style="margin: 4px 0 0; font-size: 12px; color: #94a3b8;">${meta.title}</p>
        </div>
        <div style="padding: 24px; text-align: center;">
          <img src="${qr}" alt="QR Ticket" style="width: 170px; height: 170px; border-radius: 8px; border: 1px solid #cbd5e1; padding: 6px;" />
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; margin: 16px 0;">
            <span style="font-size: 10px; color: #64748b; text-transform: uppercase; font-weight: bold; display: block;">Kode Cadangan Manual</span>
            <span style="font-family: monospace; font-size: 22px; font-weight: bold; color: #0f172a; letter-spacing: 2px;">${ticketId}</span>
          </div>
          <table style="width: 100%; text-align: left; font-size: 12px; color: #334155; border-top: 1px solid #f1f5f9; padding-top: 12px;">
            <tr><td style="color: #94a3b8; padding: 3px 0;">Nama Tamu</td><td style="font-weight: bold;">${name}</td></tr>
            <tr><td style="color: #94a3b8; padding: 3px 0;">Instansi</td><td style="font-weight: bold;">${agency}</td></tr>
            <tr><td style="color: #94a3b8; padding: 3px 0;">Jadwal</td><td style="font-weight: bold;">${meta.date} (${meta.time})</td></tr>
            <tr><td style="color: #94a3b8; padding: 3px 0;">Tempat</td><td style="font-weight: bold;">${meta.location}</td></tr>
          </table>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: to,
      subject: `[E-Ticket] ${meta.title} - ${name}`,
      htmlBody: html
    });
  },

  checkIn(ticketCode, gateName) {
    return CoreLock.execute(() => {
      const active = EventModule.getActiveEvent();
      if (!active) return { status: "error", message: "Tidak ada event yang aktif." };
      if (active.status !== "AKTIF") return { status: "error", message: "Event telah selesai / terkunci." };

      const ss = getDbSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.DATABASE.DATA_SHEET);
      if (!sheet) return { status: "error", message: "Database presensi belum ada." };

      const rows = sheet.getDataRange().getValues();
      const cleanCode = String(ticketCode).trim().toUpperCase();

      for (let i = 1; i < rows.length; i++) {
        if (rows[i][1] === active.eventId && String(rows[i][2]).trim().toUpperCase() === cleanCode) {
          if (rows[i][7] === "Hadir") {
            return {
              status: "already",
              message: "Tiket sudah pernah check-in.",
              name: rows[i][3],
              time: rows[i][8],
              gate: rows[i][9]
            };
          }

          const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss");
          sheet.getRange(i + 1, 8, 1, 3).setValues([["Hadir", now, gateName]]);

          return {
            status: "success",
            name: rows[i][3],
            agency: rows[i][4],
            time: now,
            gate: gateName
          };
        }
      }
      return { status: "not_found", message: "Kode tiket tidak ditemukan." };
    });
  }
};
