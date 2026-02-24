(function () {
  const REPORT_COLUMNS = [
    "Country",
    "Client Name",
    "Stage",
    "Current Progress",
    "Problems",
    "Next Action",
    "Follow-up Score",
    "Is Due",
  ];
  const TIMELINE_COLUMNS = [
    "Client ID",
    "Client Name",
    "Country",
    "Date",
    "Type",
    "Detail",
    "Stage After Change",
    "Days Since Previous Interaction",
  ];
  const SUMMARY_COLUMNS = [
    "Client ID",
    "Client Name",
    "Country",
    "First Contact Date",
    "Last Contact Date",
    "Total Interactions",
    "Days Active",
    "Current Stage",
    "Average Gap Days",
  ];
  const DAY_MS = 24 * 60 * 60 * 1000;

  function safeText(value) {
    return (value || "").toString().trim();
  }

  function sanitizeFilenameToken(value) {
    const cleaned = safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_");
    return cleaned || "Week";
  }

  function getClientId(client) {
    return safeText(client?.id || client?.clientId);
  }

  function getClientName(client) {
    return safeText(client?.contactName || client?.name || client?.clientName || client?.company);
  }

  function getClientCountry(client) {
    return safeText(client?.country) || "Unknown";
  }

  function getClientStage(client) {
    return safeText(client?.stage || client?.status);
  }

  function toMs(value) {
    const t = new Date(value).getTime();
    return Number.isFinite(t) ? t : null;
  }

  function toISODate(value) {
    const t = toMs(value);
    if (t === null) return safeText(value);
    return new Date(t).toISOString().slice(0, 10);
  }

  function normalizeHistoryRecord(rawRecord) {
    const r = rawRecord || {};
    const rawDate = safeText(r.date || r.timestamp || r.createdAt || r.updatedAt || r.time);
    return {
      dateRaw: rawDate,
      dateMs: toMs(rawDate),
      type: safeText(r.type || r.action || r.event || r.category),
      detail: safeText(r.detail || r.note || r.notes || r.description),
      stageAfter: safeText(r.stageAfterChange || r.stageAfter || r.stage || r.status),
    };
  }

  function sortedHistoryForClient(client) {
    const history = Array.isArray(client?.history) ? client.history : [];
    return history
      .map(normalizeHistoryRecord)
      .sort((a, b) => {
        if (a.dateMs !== null && b.dateMs !== null) return a.dateMs - b.dateMs;
        if (a.dateMs !== null) return -1;
        if (b.dateMs !== null) return 1;
        return safeText(a.dateRaw).localeCompare(safeText(b.dateRaw));
      });
  }

  function normalizeReportClient(client, updatesById) {
    const c = client || {};
    const id = safeText(c.id || c.clientId);
    const progressFromMap = updatesById instanceof Map ? safeText(updatesById.get(id)) : "";
    const currentProgress = progressFromMap || safeText(
      c.currentProgress || c.latestProgress || c.progress || c.weeklyNotes || c.notes
    );

    return {
      country: safeText(c.country) || "Unknown",
      clientName: safeText(c.contactName || c.name || c.clientName || c.company),
      stage: safeText(c.stage || c.status),
      currentProgress,
      problems: safeText(c.problems || c.problem || c.blockers || c.risks),
      nextAction: safeText(c.nextAction || c.next_step || c.plan || c.nextStep),
      followUpScore: safeText(c.followUpScore || c.score),
      isDue: c.isDue ? "Yes" : "No",
    };
  }

  function autoColumnWidths(rows) {
    const maxCols = rows.reduce((acc, row) => Math.max(acc, row.length), 0);
    const maxByCol = Array.from({ length: maxCols }, () => 0);
    rows.forEach((row) => {
      row.forEach((cell, i) => {
        const width = safeText(cell).length;
        if (width > maxByCol[i]) maxByCol[i] = width;
      });
    });
    return maxByCol.map((w) => ({ wch: Math.min(Math.max(w + 2, 12), 48) }));
  }

  function applyHeaderStyle(ws) {
    if (!ws || !ws["!ref"]) return;
    const range = window.XLSX.utils.decode_range(ws["!ref"]);

    for (let c = 0; c <= range.e.c; c += 1) {
      const addr = window.XLSX.utils.encode_cell({ r: 0, c });
      if (!ws[addr]) continue;
      ws[addr].s = {
        font: { bold: true },
        alignment: { vertical: "center", horizontal: "center", wrapText: true },
      };
    }
  }

  function applyCountryHeaderStyle(ws, countryHeaderRows) {
    if (!ws) return;
    countryHeaderRows.forEach((rowIndex) => {
      const addr = window.XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
      if (!ws[addr]) return;
      ws[addr].s = {
        font: { bold: true },
        alignment: { vertical: "center", horizontal: "left" },
      };
    });
  }

  function buildGroupedReportRows(sourceClients, updatesById) {
    const normalized = sourceClients
      .map((client) => normalizeReportClient(client, updatesById))
      .filter((c) => c.clientName);

    normalized.sort((a, b) => {
      const byCountry = a.country.localeCompare(b.country);
      if (byCountry !== 0) return byCountry;
      return a.clientName.localeCompare(b.clientName);
    });

    const rows = [REPORT_COLUMNS];
    const merges = [];
    const countryHeaderRows = [];

    let currentCountry = "";
    normalized.forEach((c) => {
      if (c.country !== currentCountry) {
        currentCountry = c.country;
        const headerRowIndex = rows.length;
        rows.push([currentCountry, "", "", "", "", "", "", ""]);
        merges.push({ s: { r: headerRowIndex, c: 0 }, e: { r: headerRowIndex, c: 7 } });
        countryHeaderRows.push(headerRowIndex);
      }

      rows.push([
        c.country,
        c.clientName,
        c.stage,
        c.currentProgress,
        c.problems,
        c.nextAction,
        c.followUpScore,
        c.isDue,
      ]);
    });

    return { rows, merges, countryHeaderRows };
  }

  function buildTimelineRows(sourceClients) {
    const rows = [TIMELINE_COLUMNS];
    let totalHistoryRecords = 0;

    sourceClients.forEach((client) => {
      const clientId = getClientId(client);
      const clientName = getClientName(client);
      const country = getClientCountry(client);
      const history = sortedHistoryForClient(client);

      let prevMs = null;
      history.forEach((record) => {
        const daysSincePrev = (prevMs !== null && record.dateMs !== null)
          ? String(Math.max(0, Math.round((record.dateMs - prevMs) / DAY_MS)))
          : "";

        rows.push([
          clientId,
          clientName,
          country,
          toISODate(record.dateRaw),
          record.type,
          record.detail,
          record.stageAfter,
          daysSincePrev,
        ]);

        if (record.dateMs !== null) prevMs = record.dateMs;
        totalHistoryRecords += 1;
      });
    });

    return { rows, totalHistoryRecords };
  }

  function buildLifecycleSummaryRows(sourceClients) {
    const rows = [SUMMARY_COLUMNS];
    const todayMs = Date.now();

    sourceClients.forEach((client) => {
      const clientId = getClientId(client);
      const clientName = getClientName(client);
      const country = getClientCountry(client);
      const currentStage = getClientStage(client);
      const history = sortedHistoryForClient(client);
      const datePoints = history.map((h) => h.dateMs).filter((ms) => ms !== null);

      const firstMs = datePoints.length ? datePoints[0] : null;
      const lastMs = datePoints.length ? datePoints[datePoints.length - 1] : null;
      const daysActive = firstMs !== null ? String(Math.max(0, Math.round((todayMs - firstMs) / DAY_MS))) : "";

      const gaps = [];
      for (let i = 1; i < datePoints.length; i += 1) {
        gaps.push(Math.max(0, Math.round((datePoints[i] - datePoints[i - 1]) / DAY_MS)));
      }
      const avgGap = gaps.length
        ? (gaps.reduce((sum, d) => sum + d, 0) / gaps.length).toFixed(1)
        : "";

      rows.push([
        clientId,
        clientName,
        country,
        firstMs !== null ? toISODate(firstMs) : "",
        lastMs !== null ? toISODate(lastMs) : "",
        String(history.length),
        daysActive,
        currentStage,
        avgGap,
      ]);
    });

    rows.splice(
      1,
      rows.length - 1,
      ...rows.slice(1).sort((a, b) => {
        const byCountry = safeText(a[2]).localeCompare(safeText(b[2]));
        if (byCountry !== 0) return byCountry;
        return safeText(a[1]).localeCompare(safeText(b[1]));
      })
    );

    return rows;
  }

  async function exportWeeklyWorkbook(options) {
    if (!window.XLSX || !window.XLSX.utils || !window.XLSX.writeFile) {
      throw new Error("SheetJS is required for Excel export.");
    }

    const weekLabel = safeText(options?.weekLabel || `Week ${new Date().toISOString().slice(0, 10)}`);
    const sourceClients = Array.isArray(options?.clients) ? options.clients : [];
    const updatesById = options?.updatesById instanceof Map ? options.updatesById : new Map();

    const grouped = buildGroupedReportRows(sourceClients, updatesById);
    const timeline = buildTimelineRows(sourceClients);
    const lifecycleRows = buildLifecycleSummaryRows(sourceClients);

    const wsWeekly = window.XLSX.utils.aoa_to_sheet(grouped.rows);
    wsWeekly["!merges"] = grouped.merges;
    wsWeekly["!cols"] = autoColumnWidths(grouped.rows);
    wsWeekly["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
    wsWeekly["!autofilter"] = { ref: "A1:H1" };
    applyHeaderStyle(wsWeekly);
    applyCountryHeaderStyle(wsWeekly, grouped.countryHeaderRows);

    const wsTimeline = window.XLSX.utils.aoa_to_sheet(timeline.rows);
    wsTimeline["!cols"] = autoColumnWidths(timeline.rows);
    wsTimeline["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
    wsTimeline["!autofilter"] = { ref: "A1:H1" };
    applyHeaderStyle(wsTimeline);

    const wsLifecycle = window.XLSX.utils.aoa_to_sheet(lifecycleRows);
    wsLifecycle["!cols"] = autoColumnWidths(lifecycleRows);
    wsLifecycle["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
    wsLifecycle["!autofilter"] = { ref: "A1:I1" };
    applyHeaderStyle(wsLifecycle);

    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, wsWeekly, "Weekly_Report");
    window.XLSX.utils.book_append_sheet(wb, wsTimeline, "Full_Interaction_Timeline");
    window.XLSX.utils.book_append_sheet(wb, wsLifecycle, "Client_Lifecycle_Summary");
    window.XLSX.writeFile(wb, `Weekly_Client_Progress_${sanitizeFilenameToken(weekLabel)}.xlsx`);

    return {
      totalClients: Math.max(grouped.rows.length - grouped.countryHeaderRows.length - 1, 0),
      countrySections: grouped.countryHeaderRows.length,
      totalHistoryRecords: timeline.totalHistoryRecords,
      weekLabel,
    };
  }

  window.excelExport = {
    BASE_COLUMNS: REPORT_COLUMNS,
    REPORT_COLUMNS,
    exportWeeklyWorkbook,
  };
})();
