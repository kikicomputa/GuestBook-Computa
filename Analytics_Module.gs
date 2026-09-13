const AnalyticsModule = {
  // Helper: Mengekstrak waktu presisi kapan event dibuat dari Event ID
  parseEventCreationTime(eventId) {
    if (!eventId || !eventId.startsWith("EVT-")) return new Date(0);
    try {
      const parts = eventId.replace("EVT-", "").split("-");
      if (parts.length === 2 && parts[0].length === 8 && parts[1].length === 6) {
        const y = parseInt(parts[0].substring(0, 4), 10);
        const m = parseInt(parts[0].substring(4, 6), 10) - 1;
        const d = parseInt(parts[0].substring(6, 8), 10);
        const hh = parseInt(parts[1].substring(0, 2), 10);
        const mm = parseInt(parts[1].substring(2, 4), 10);
        const ss = parseInt(parts[1].substring(4, 6), 10);
        return new Date(y, m, d, hh, mm, ss);
      }
    } catch (e) {}
    return new Date(0);
  },

  // Sinkronisasi hanya data form yang masuk SETELAH event dibuat
  syncFormToMaster(ss, activeEvent) {
    if (!activeEvent || activeEvent.status === "SELESAI" || !activeEvent.eventId) return;

    let master = ss.getSheetByName(CONFIG.DATABASE.DATA_SHEET);
    if (!master) {
      master = ss.insertSheet(CONFIG.DATABASE.DATA_SHEET);
      master.appendRow(["Timestamp", "Event ID", "Ticket ID", "Nama Lengkap", "Instansi", "Email", "WhatsApp", "Status", "Waktu Masuk", "Gate"]);
    }

    // Waktu pembuatan event aktif saat ini
    const eventCreatedAt = this.parseEventCreationTime(activeEvent.eventId);

    // Kumpulkan entri yang sudah tercatat di master untuk event aktif ini
    const existingData = master.getLastRow() > 1 ? master.getDataRange().getValues() : [];
    const registeredKeys = new Set();
    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][1]).trim() === activeEvent.eventId) {
        const key = String(existingData[i][3]).trim().toLowerCase() + "_" + String(existingData[i][5]).trim().toLowerCase();
        registeredKeys.add(key);
      }
    }

    const sheets = ss.getSheets();
    sheets.forEach(sh => {
      const sName = sh.getName().toLowerCase();
      // Periksa tab respon Google Form
      if ((sName.includes("form") || sName.includes("jawaban") || sName.includes("response")) && sh.getName() !== CONFIG.DATABASE.DATA_SHEET) {
        const lastRow = sh.getLastRow();
        if (lastRow > 1) {
          const formRows = sh.getDataRange().getValues();
          for (let r = 1; r < formRows.length; r++) {
            const row = formRows[r];
            const rawTime = row[0];
            const submitTime = rawTime instanceof Date ? rawTime : new Date(rawTime);

            // PENGAMAN KRUSIAL: Lewati jika pendaftaran dilakukan SEBELUM event ini dibuat
            if (submitTime < eventCreatedAt) {
              continue;
            }

            const email = String(row[1] || "-").trim();
            const name = String(row[2] || "-").trim();
            const agency = String(row[3] || "-").trim();
            const wa = String(row[4] || "-").trim();

            if (!name || name === "-") continue;

            const checkKey = name.toLowerCase() + "_" + email.toLowerCase();
            if (registeredKeys.has(checkKey)) continue;

            // Buat Ticket ID baru
            const ticketId = "TIK-" + Math.floor(1000 + Math.random() * 9000);

            master.appendRow([
              rawTime,
              activeEvent.eventId,
              ticketId,
              name,
              agency,
              email,
              wa,
              "Belum",
              "-",
              "-"
            ]);

            registeredKeys.add(checkKey);

            // Kirim tiket QR ke email
            if (email.includes("@") && typeof AttendanceModule !== 'undefined') {
              try {
                AttendanceModule.sendEmail(email, name, agency, ticketId, activeEvent);
              } catch (mailErr) {
                Logger.log("Gagal kirim email: " + mailErr.toString());
              }
            }
          }
        }
      }
    });
  },

  getDashboardData() {
    try {
      const active = EventModule.getActiveEvent();

      // KONDISI 1: TIDAK ADA EVENT AKTIF / EVENT DIARSIPKAN
      if (!active || active.status === "SELESAI" || !active.eventId) {
        return {
          status: "empty",
          meta: { eventId: "", title: "Tidak Ada Event Aktif", status: "SELESAI", formUrl: "" },
          stats: { total: 0, hadir: 0, belum: 0, percent: 0, gates: { "Lajur 01": 0, "Lajur 02": 0, "Lajur 03": 0, "Lajur 04": 0, "Lajur 05": 0 } },
          guests: []
        };
      }

      const ss = getDbSpreadsheet();

      // Tarik data baru dengan penyaring waktu
      this.syncFormToMaster(ss, active);

      const master = ss.getSheetByName(CONFIG.DATABASE.DATA_SHEET);
      if (!master || master.getLastRow() < 2) {
        return {
          status: "success",
          meta: active,
          stats: { total: 0, hadir: 0, belum: 0, percent: 0, gates: { "Lajur 01": 0, "Lajur 02": 0, "Lajur 03": 0, "Lajur 04": 0, "Lajur 05": 0 } },
          guests: []
        };
      }

      const rows = master.getDataRange().getValues();
      let total = 0;
      let hadir = 0;
      const gates = { "Lajur 01": 0, "Lajur 02": 0, "Lajur 03": 0, "Lajur 04": 0, "Lajur 05": 0 };
      const guests = [];
      const currentEventId = String(active.eventId).trim();

      // Hanya proses baris yang cocok persis dengan Event ID aktif saat ini
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (String(r[1]).trim() !== currentEventId) continue;
        if (!r[2] && !r[3]) continue;

        total++;
        const isHadir = (String(r[7] || "").trim().toLowerCase() === "hadir");

        if (isHadir) {
          hadir++;
          const gateRaw = String(r[9] || "").trim();
          if (gates[gateRaw] !== undefined) {
            gates[gateRaw]++;
          } else if (gateRaw.includes("1")) gates["Lajur 01"]++;
          else if (gateRaw.includes("2")) gates["Lajur 02"]++;
          else if (gateRaw.includes("3")) gates["Lajur 03"]++;
          else if (gateRaw.includes("4")) gates["Lajur 04"]++;
          else if (gateRaw.includes("5")) gates["Lajur 05"]++;
        }

        guests.push({
          ticket: String(r[2] || "-"),
          name: String(r[3] || "-"),
          agency: String(r[4] || "-"),
          wa: String(r[6] || "-"),
          status: isHadir ? "Hadir" : "Belum",
          time: String(r[8] || "-"),
          gate: String(r[9] || "-")
        });
      }

      const percent = total > 0 ? Math.round((hadir / total) * 100) : 0;

      return {
        status: "success",
        meta: active,
        stats: {
          total: total,
          hadir: hadir,
          belum: (total - hadir),
          percent: percent,
          gates: gates
        },
        guests: guests
      };

    } catch (err) {
      Logger.log("Error getDashboardData: " + err.toString());
      return {
        status: "empty",
        meta: { title: "Tidak Ada Event Aktif", status: "SELESAI" },
        stats: { total: 0, hadir: 0, belum: 0, percent: 0, gates: {} },
        guests: []
      };
    }
  },

  getAllArchives() {
    const ss = getDbSpreadsheet();
    const master = ss.getSheetByName(CONFIG.DATABASE.MASTER_SHEET);
    if (!master || master.getLastRow() < 2) return [];

    const rows = master.getDataRange().getValues();
    const archives = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      archives.push({
        id: String(r[0]),
        title: String(r[1]),
        schedule: `${r[2]} (${r[3]})`,
        loc: String(r[4]),
        status: String(r[7] || "SELESAI")
      });
    }
    return archives.reverse();
  }
};
