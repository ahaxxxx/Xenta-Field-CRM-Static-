(function () {
  const REPORT_COLUMNS = [
    "\u516c\u53f8\u540d\u79f0",
    "\u56fd\u5bb6",
    "\u57ce\u5e02",
    "\u516c\u53f8\u7b49\u7ea7",
    "\u8054\u7cfb\u4eba\u59d3\u540d",
    "\u804c\u4f4d",
    "\u7535\u8bdd",
    "\u90ae\u7bb1",
    "WhatsApp",
    "LinkedIn",
    "\u4ea7\u54c1\u65b9\u5411",
    "\u81ea\u5b9a\u4e49\u4ea7\u54c1\u65b9\u5411",
    "\u7ade\u4e89\u54c1\u724c",
    "\u5f53\u524d\u9636\u6bb5",
    "\u4f18\u5148\u7ea7",
    "\u6700\u540e\u8054\u7cfb\u65e5\u671f",
    "\u4e0b\u6b21\u8ddf\u8fdb\u65e5\u671f",
    "\u603b\u8bc4\u5206",
    "\u6c9f\u901a\u5386\u53f2",
  ];
  const TIMELINE_COLUMNS = [
    "\u5ba2\u6237ID",
    "\u5ba2\u6237\u540d\u79f0",
    "\u56fd\u5bb6",
    "\u65e5\u671f",
    "\u6c9f\u901a\u65b9\u5f0f",
    "\u6c9f\u901a\u5185\u5bb9",
    "\u9636\u6bb5\uff08\u53d8\u66f4\u540e\uff09",
    "\u8ddd\u4e0a\u6b21\u6c9f\u901a\u5929\u6570",
  ];
  const SUMMARY_COLUMNS = [
    "\u5ba2\u6237ID",
    "\u5ba2\u6237\u540d\u79f0",
    "\u56fd\u5bb6",
    "\u9996\u6b21\u6c9f\u901a\u65e5\u671f",
    "\u6700\u540e\u6c9f\u901a\u65e5\u671f",
    "\u603b\u6c9f\u901a\u6b21\u6570",
    "\u8ddf\u8fdb\u5468\u671f\uff08\u5929\uff09",
    "\u5f53\u524d\u9636\u6bb5",
    "\u5e73\u5747\u95f4\u9694\uff08\u5929\uff09",
  ];
  const DAY_MS = 24 * 60 * 60 * 1000;

  function safeText(value) {
    return (value || "").toString().trim();
  }

  function sanitizeFilenameToken(value) {
    const cleaned = safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_");
    return cleaned || "Week";
  }

  function getDateToken() {
    const today = new Date();
    return (
      today.getFullYear().toString() +
      String(today.getMonth() + 1).padStart(2, "0") +
      String(today.getDate()).padStart(2, "0")
    );
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
    const history = sortedHistoryForClient(c);
    const historyText = history.map((h) => {
      const d = toISODate(h.dateRaw);
      const t = safeText(h.type) || "Note";
      const detail = safeText(h.detail);
      return detail ? `[${d}][${t}] ${detail}` : "";
    }).filter(Boolean).join("\n");
    const focusList = Array.isArray(c.productFocus)
      ? c.productFocus.map((v) => safeText(v)).filter(Boolean)
      : safeText(c.interest).split(",").map((v) => safeText(v)).filter(Boolean);
    const knownFocus = new Set(["CLIA POCT", "FIA", "LFA", "PCR", "Rapid test"]);
    const customFocus = focusList.filter((v) => !knownFocus.has(v));
    const position = safeText(c.positionCustom || c.positionCategory || c.role);

    return {
      companyName: safeText(c.company),
      country: safeText(c.country) || "Unknown",
      city: safeText(c.city || c.region),
      companyLevel: safeText(c.companyLevel),
      contactName: safeText(c.contactName || c.name || c.clientName),
      position,
      phone: safeText(c.phone),
      email: safeText(c.email),
      whatsapp: safeText(c.phone),
      linkedIn: safeText(c.website),
      productFocus: focusList.filter((v) => knownFocus.has(v)).join(", "),
      productFocusCustom: customFocus.join(", "),
      brands: safeText(c.brands),
      stage: safeText(c.stage || c.status),
      priority: safeText(c.priority || c.level),
      lastContactDate: toISODate(c.lastContactDate || c.lastContact || c.updatedAt),
      nextFollowUpDate: safeText(c.nextFollowUpDate || c.next_contact_date || c.nextStepDate),
      totalScore: safeText(c.followUpScore || c.score),
      communicationHistory: historyText || safeText(c.communicationLog),
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
      .filter((c) => c.companyName || c.contactName);

    normalized.sort((a, b) => {
      const byCountry = a.country.localeCompare(b.country);
      if (byCountry !== 0) return byCountry;
      const byCompany = safeText(a.companyName).localeCompare(safeText(b.companyName));
      if (byCompany !== 0) return byCompany;
      return safeText(a.contactName).localeCompare(safeText(b.contactName));
    });

    const rows = [REPORT_COLUMNS];
    normalized.forEach((c) => {
      rows.push([
        c.companyName,
        c.country,
        c.city,
        c.companyLevel,
        c.contactName,
        c.position,
        c.phone,
        c.email,
        c.whatsapp,
        c.linkedIn,
        c.productFocus,
        c.productFocusCustom,
        c.brands,
        c.stage,
        c.priority,
        c.lastContactDate,
        c.nextFollowUpDate,
        c.totalScore,
        c.communicationHistory,
      ]);
    });

    return { rows, merges: [], countryHeaderRows: [] };
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
    const weeklyLastCol = window.XLSX.utils.encode_col(REPORT_COLUMNS.length - 1);
    wsWeekly["!autofilter"] = { ref: `A1:${weeklyLastCol}1` };
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
    window.XLSX.utils.book_append_sheet(wb, wsWeekly, "\u5468\u8fdb\u5c55\u62a5\u544a");
    window.XLSX.utils.book_append_sheet(wb, wsTimeline, "\u5b8c\u6574\u6c9f\u901a\u65f6\u95f4\u7ebf");
    window.XLSX.utils.book_append_sheet(wb, wsLifecycle, "\u5ba2\u6237\u751f\u547d\u5468\u671f\u6458\u8981");
    const excelBuffer = window.XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const fileName = `Xenta_Weekly_Report_${getDateToken()}.xlsx`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);

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
