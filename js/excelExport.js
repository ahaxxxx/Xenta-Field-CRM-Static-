(function () {
  const BASE_COLUMNS = [
    "Client ID",
    "Client Name",
    "Country",
    "City",
    "Role/Title",
    "Organization",
    "Email",
    "WhatsApp",
    "LinkedIn",
    "Source",
    "Product Interest",
    "Stage",
    "Priority",
    "Owner",
    "Last Contact Date",
  ];

  function safeText(value) {
    return (value || "").toString().trim();
  }

  function xmlEscape(value) {
    return safeText(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function sanitizeFilenameToken(value) {
    const cleaned = safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_");
    return cleaned || "Week";
  }

  function isWeekColumn(name) {
    return /^Week\s+\d{4}-\d{2}-\d{2}$/i.test(safeText(name));
  }

  function weekSortValue(label) {
    const m = safeText(label).match(/(\d{4}-\d{2}-\d{2})/);
    if (!m) return Number.MAX_SAFE_INTEGER;
    const t = new Date(`${m[1]}T00:00:00Z`).getTime();
    return Number.isFinite(t) ? t : Number.MAX_SAFE_INTEGER;
  }

  function normalizeClient(client) {
    const c = client || {};
    return {
      "Client ID": safeText(c.id || c.clientId),
      "Client Name": safeText(c.contactName || c.name || c.clientName || c.company),
      "Country": safeText(c.country),
      "City": safeText(c.city),
      "Role/Title": safeText(c.role || c.title),
      "Organization": safeText(c.company || c.organization || c.org),
      "Email": safeText(c.email),
      "WhatsApp": safeText(c.whatsapp || c.whatsApp || c.phone),
      "LinkedIn": safeText(c.linkedin || c.linkedIn),
      "Source": safeText(c.source),
      "Product Interest": safeText(c.interest || c.productInterest || c.focus),
      "Stage": safeText(c.stage || c.status),
      "Priority": safeText(c.priority || c.level),
      "Owner": safeText(c.owner || c.salesOwner || c.accountOwner),
      "Last Contact Date": safeText(c.lastContactDate || c.lastContact || c.updatedAt || c.updated_at),
    };
  }

  function parseCellTexts(rowEl) {
    const out = [];
    let cursor = 1;
    const cells = Array.from(rowEl.children).filter((el) => el.localName === "Cell");

    cells.forEach((cellEl) => {
      const idxRaw = cellEl.getAttribute("ss:Index") || cellEl.getAttribute("Index");
      const idx = parseInt(idxRaw, 10);
      if (Number.isFinite(idx) && idx > 0) {
        cursor = idx;
      }

      const dataEl = Array.from(cellEl.children).find((el) => el.localName === "Data");
      out[cursor - 1] = dataEl ? safeText(dataEl.textContent) : "";
      cursor += 1;
    });

    return out.map((v) => safeText(v));
  }

  function findWorksheetByName(doc, targetName) {
    const sheets = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "Worksheet");
    return sheets.find((el) => safeText(el.getAttribute("ss:Name") || el.getAttribute("Name")) === targetName) || null;
  }

  async function parseExistingWorkbook(file) {
    if (!file) {
      return { rowsById: new Map(), weeklyColumns: [] };
    }

    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
      console.warn("[excelExport] Uploaded .xlsx appears to be ZIP-based. Merge parser supports SpreadsheetML XML exports.");
      return { rowsById: new Map(), weeklyColumns: [] };
    }

    const text = new TextDecoder("utf-8").decode(buffer);
    if (!text.includes("<Workbook")) {
      console.warn("[excelExport] Unsupported workbook format. Expected SpreadsheetML XML content.");
      return { rowsById: new Map(), weeklyColumns: [] };
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(text, "application/xml");
    const clientsSheet = findWorksheetByName(doc, "Clients");
    if (!clientsSheet) {
      console.warn("[excelExport] No 'Clients' sheet found in uploaded workbook.");
      return { rowsById: new Map(), weeklyColumns: [] };
    }

    const table = Array.from(clientsSheet.children).find((el) => el.localName === "Table");
    if (!table) return { rowsById: new Map(), weeklyColumns: [] };

    const rows = Array.from(table.children).filter((el) => el.localName === "Row");
    if (!rows.length) return { rowsById: new Map(), weeklyColumns: [] };

    const headers = parseCellTexts(rows[0]);
    const idIdx = headers.indexOf("Client ID");
    if (idIdx === -1) {
      console.warn("[excelExport] 'Client ID' column not found in uploaded workbook.");
      return { rowsById: new Map(), weeklyColumns: [] };
    }

    const weeklyColumns = headers
      .filter((h) => isWeekColumn(h))
      .sort((a, b) => weekSortValue(a) - weekSortValue(b) || a.localeCompare(b));

    const rowsById = new Map();
    rows.slice(1).forEach((rowEl) => {
      const cells = parseCellTexts(rowEl);
      const id = safeText(cells[idIdx]);
      if (!id) return;

      const rowObj = {};
      headers.forEach((h, i) => {
        rowObj[h] = safeText(cells[i] || "");
      });
      rowsById.set(id, rowObj);
    });

    return { rowsById, weeklyColumns };
  }

  function buildSummaryRows(rows, weekLabel) {
    const stageCount = new Map();
    let highPriorityCount = 0;
    let updatedThisWeek = 0;

    rows.forEach((row) => {
      const stage = safeText(row["Stage"]) || "Unknown";
      stageCount.set(stage, (stageCount.get(stage) || 0) + 1);

      const priority = safeText(row["Priority"]).toLowerCase();
      if (priority.includes("high") || priority === "a") highPriorityCount += 1;
      if (safeText(row[weekLabel])) updatedThisWeek += 1;
    });

    const summaryRows = [
      ["Metric", "Value"],
      ["Total clients", String(rows.length)],
      ["Clients updated this week", String(updatedThisWeek)],
      ["High priority count", String(highPriorityCount)],
      ["", ""],
      ["Count by Stage", ""],
    ];

    Array.from(stageCount.keys()).sort((a, b) => a.localeCompare(b)).forEach((stage) => {
      summaryRows.push([stage, String(stageCount.get(stage))]);
    });

    return summaryRows;
  }

  function makeSheetXml(name, rows) {
    const rowXml = rows
      .map((row) => {
        const cells = row
          .map((value) => `<Cell><Data ss:Type="String">${xmlEscape(value)}</Data></Cell>`)
          .join("");
        return `<Row>${cells}</Row>`;
      })
      .join("");

    return `
  <Worksheet ss:Name="${xmlEscape(name).slice(0, 31)}">
    <Table>
      ${rowXml}
    </Table>
  </Worksheet>`;
  }

  function buildWorkbookXml(clientsColumns, clientRows, includeSummary, weekLabel) {
    const clientsRows = [clientsColumns, ...clientRows.map((row) => clientsColumns.map((c) => safeText(row[c])))];
    const clientsSheet = makeSheetXml("Clients", clientsRows);
    const summarySheet = includeSummary ? makeSheetXml("Weekly Summary", buildSummaryRows(clientRows, weekLabel)) : "";

    return `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
${clientsSheet}
${summarySheet}
</Workbook>`;
  }

  function downloadWorkbook(xml, weekLabel) {
    const blob = new Blob([xml], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Weekly_Client_Progress_${sanitizeFilenameToken(weekLabel)}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportWeeklyWorkbook(options) {
    const weekLabel = safeText(options?.weekLabel || `Week ${new Date().toISOString().slice(0, 10)}`);
    const existingFile = options?.existingFile || null;
    const updatesById = options?.updatesById instanceof Map ? options.updatesById : new Map();
    const includeSummary = options?.includeSummary !== false;
    const sourceClients = Array.isArray(options?.clients) ? options.clients : [];

    const existing = await parseExistingWorkbook(existingFile);
    const normalizedClients = sourceClients.map(normalizeClient).filter((c) => safeText(c["Client ID"]));
    const crmById = new Map(normalizedClients.map((c) => [c["Client ID"], c]));

    const weeklyColumns = [...existing.weeklyColumns];
    if (!weeklyColumns.includes(weekLabel)) weeklyColumns.push(weekLabel);
    weeklyColumns.sort((a, b) => weekSortValue(a) - weekSortValue(b) || a.localeCompare(b));

    const allIds = Array.from(new Set([
      ...Array.from(existing.rowsById.keys()),
      ...Array.from(crmById.keys()),
    ])).sort((a, b) => a.localeCompare(b));

    const mergedRows = allIds.map((id) => {
      const existingRow = existing.rowsById.get(id) || {};
      const crmRow = crmById.get(id) || { "Client ID": id };
      const row = {};

      BASE_COLUMNS.forEach((col) => {
        if (col === "Client ID") {
          row[col] = id;
          return;
        }
        row[col] = safeText(crmRow[col] || existingRow[col] || "");
      });

      weeklyColumns.forEach((wk) => {
        row[wk] = safeText(existingRow[wk] || "");
      });

      if (updatesById.has(id)) {
        row[weekLabel] = safeText(updatesById.get(id));
      } else if (!safeText(row[weekLabel])) {
        row[weekLabel] = "";
      }

      return row;
    });

    const columns = [...BASE_COLUMNS, ...weeklyColumns];
    const xml = buildWorkbookXml(columns, mergedRows, includeSummary, weekLabel);
    downloadWorkbook(xml, weekLabel);

    return {
      totalClients: mergedRows.length,
      updatedThisWeek: mergedRows.filter((r) => safeText(r[weekLabel])).length,
      weekLabel,
      usedExistingWorkbook: !!existingFile,
    };
  }

  window.excelExport = {
    BASE_COLUMNS,
    exportWeeklyWorkbook,
  };
})();
