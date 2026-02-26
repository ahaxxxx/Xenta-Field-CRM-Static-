/**************************************/

const DB_KEY = "xenta_crm_distributors_v1";
const LANG_KEY = "xenta_crm_lang";

let distributors = [];
let currentEditId = null;
let currentLang = "en";

let sortState = { key: "score", direction: "desc" };
let sortMode = "score";

const followUpDays = {
  A: 2,
  B: 5,
  C: 10,
  D: 30,
};

const gradeWeightMap = {
  A: 3,
  B: 2,
  C: 1,
  D: 0.5,
};

const stageWeightMap = {
  new: 1,
  contacted: 2,
  negotiation: 5,
  quotation: 6,
  quoted: 6,
  contract: 8,
  closed: 0,
};

const i18n = {
  en: {
    "app.title": "Xenta Field CRM",
    "app.subtitle": "Distributor intelligence & follow-up toolkit",
    "actions.export": "Export JSON",
    "actions.import": "Import JSON",
    "actions.reset": "Reset",
    "actions.new": "+ New record",
    "actions.save": "Save",
    "actions.delete": "Delete",
    "actions.generate": "Generate",
    "actions.copy": "Copy",
    "actions.sortByScore": "Sort by Score",
    "actions.sortByLastContact": "Sort by Last Contact",
    "common.all": "All",
    "filters.title": "Filters",
    "filters.search": "Search",
    "filters.country": "Country",
    "filters.region": "Region",
    "filters.stage": "Stage",
    "filters.grade": "Grade",
    "filters.focus": "Product Focus",
    "filters.hint": "Tip: Use Export/Import to sync across devices.",
    "table.title": "Leads",
    "table.name": "Name",
    "table.country": "Country",
    "table.region": "Region",
    "table.stage": "Stage",
    "table.grade": "Grade",
    "table.last": "Last",
    "table.score": "🔥 Score",
    "table.actions": "Actions",
    "top5.title": "🔥 Today’s Top 5",
    "top5.empty": "No due follow-ups today.",
    "status.due": "Due",
    "details.title": "Details",
    "details.emptyTitle": "Select a record",
    "details.emptySub": "or create a new one to start.",
    "details.name": "Name",
    "details.country": "Country",
    "details.region": "Region",
    "details.stage": "Stage",
    "details.grade": "Grade",
    "details.contacts": "Contacts",
    "details.brands": "Brands / Competitors",
    "details.focus": "Product Focus",
    "details.notes": "Notes",
    "details.communicationLog": "Communication Log",
    "details.last": "Last contact",
    "details.next": "Next step",
    "msg.title": "Message Generator",
    "msg.channel": "Channel",
    "msg.lang": "Language",
    "msg.type": "Template",
    "msg.output": "Output",
    "msg.tpl.intro": "Intro / Product pitch",
    "msg.tpl.catalog": "Send catalog",
    "msg.tpl.meeting": "Schedule meeting",
    "msg.tpl.followup": "Follow up",
    "footer.note": "Local-only. No backend. Data stays in your browser.",
    "weekly.title": "Weekly Progress Export",
    "weekly.close": "Close",
    "weekly.weekLabel": "Week Label",
    "weekly.progressSource": "Progress Source",
    "weekly.source.latest": "A) Latest progress field",
    "weekly.source.weekly": "B) Weekly notes field",
    "weekly.source.manual": "C) Manual paste mapping",
    "weekly.manualMapping": "Manual Mapping (Client ID[TAB]Update per line)",
    "weekly.uploadExisting": "Upload last report (optional)",
    "weekly.hint": "Exports without page refresh.",
    "weekly.generateSummary": "Generate Weekly Summary Text",
    "weekly.export": "Export",
    "weekly.summaryText": "Weekly Summary Text",
    "weekly.copySummary": "Copy Summary",
  },
  zh: {
    "app.title": "Xenta 外勤 CRM",
    "app.subtitle": "经销商情报与跟进工具",
    "actions.export": "导出 JSON",
    "actions.import": "导入 JSON",
    "actions.reset": "重置",
    "actions.new": "+ 新建客户",
    "actions.save": "保存",
    "actions.delete": "删除",
    "actions.generate": "生成",
    "actions.copy": "复制",
    "actions.sortByScore": "按评分排序",
    "actions.sortByLastContact": "按最近联系排序",
    "common.all": "全部",
    "filters.title": "筛选",
    "filters.search": "搜索",
    "filters.country": "国家",
    "filters.region": "区域",
    "filters.stage": "阶段",
    "filters.grade": "等级",
    "filters.focus": "产品方向",
    "filters.hint": "提示：可用导出/导入在设备间同步。",
    "table.title": "客户列表",
    "table.name": "名称",
    "table.country": "国家",
    "table.region": "区域",
    "table.stage": "阶段",
    "table.grade": "等级",
    "table.last": "最近联系",
    "table.score": "🔥 评分",
    "table.actions": "操作",
    "top5.title": "🔥 今日优先 Top 5",
    "top5.empty": "今天没有到期跟进客户。",
    "status.due": "到期",
    "details.title": "详情",
    "details.emptyTitle": "请选择一条记录",
    "details.emptySub": "或先新建一条记录。",
    "details.name": "名称",
    "details.country": "国家",
    "details.region": "区域",
    "details.stage": "阶段",
    "details.grade": "等级",
    "details.contacts": "联系方式",
    "details.brands": "品牌 / 竞品",
    "details.focus": "产品方向",
    "details.notes": "备注",
    "details.last": "最近联系",
    "details.next": "下一步",
    "msg.title": "消息生成",
    "msg.channel": "渠道",
    "msg.lang": "语言",
    "msg.type": "模板",
    "msg.output": "输出",
    "msg.tpl.intro": "介绍 / 产品推荐",
    "msg.tpl.catalog": "发送目录",
    "msg.tpl.meeting": "安排会议",
    "msg.tpl.followup": "跟进",
    "footer.note": "纯本地运行，无后端，数据保存在浏览器。",
    "weekly.title": "周进展导出",
    "weekly.close": "关闭",
    "weekly.weekLabel": "周标签",
    "weekly.progressSource": "进展来源",
    "weekly.source.latest": "A) 最新进展字段",
    "weekly.source.weekly": "B) 周备注字段",
    "weekly.source.manual": "C) 手动粘贴映射",
    "weekly.manualMapping": "手动映射（每行：客户ID[TAB]更新）",
    "weekly.uploadExisting": "上传上周报表（可选）",
    "weekly.hint": "无需刷新页面即可导出。",
    "weekly.generateSummary": "生成周总结文本",
    "weekly.export": "导出",
    "weekly.summaryText": "周总结文本",
    "weekly.copySummary": "复制总结",
  },
};

// ---------------------------
// Utilities
// ---------------------------
function uid() {
  return "id_" + Math.random().toString(16).slice(2) + "_" + Date.now();
}

function nowISO() {
  return new Date().toISOString();
}

function formatDate(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  return d.toLocaleString();
}

function safeText(s) {
  return (s || "").toString().trim();
}

function t(key) {
  const dict = i18n[currentLang] || i18n.en;
  return dict[key] || i18n.en[key] || key;
}

function applyI18n() {
  document.querySelectorAll("[data-i18n]").forEach((node) => {
    const key = node.getAttribute("data-i18n");
    if (!key) return;
    node.textContent = t(key);
  });
  document.documentElement.lang = currentLang;
}

function setLanguage(lang, options = {}) {
  const normalized = safeText(lang).toLowerCase();
  currentLang = i18n[normalized] ? normalized : "en";
  if (options.persist !== false) {
    try {
      localStorage.setItem(LANG_KEY, currentLang);
    } catch (_) {}
  }
  applyI18n();
  if (options.rerender) render();
}

function detectInitialLanguage() {
  let saved = "";
  try {
    saved = safeText(localStorage.getItem(LANG_KEY)).toLowerCase();
  } catch (_) {}
  if (saved && i18n[saved]) return saved;
  return "en";
}

function todayDateISO() {
  return new Date().toISOString().slice(0, 10);
}

function normalizeGrade(rawGrade) {
  const g = safeText(rawGrade).toUpperCase();
  return followUpDays[g] ? g : "C";
}

function ensureClientSafety(rawClient) {
  const client = { ...(rawClient || {}) };
  const safeGrade = normalizeGrade(client.grade || client.level);
  const safeStage = safeText(client.stage || client.status) || "New";
  const safePriority = safeText(client.priority) || "Normal";
  const safeLastContactDate = safeText(
    client.lastContactDate || client.lastContact || client.updatedAt || client.updated_at
  ) || todayDateISO();

  client.name = safeText(client.name || client.company);
  client.priority = safePriority;
  client.grade = safeGrade;
  client.stage = safeStage;
  client.lastContactDate = safeLastContactDate;

  return client;
}

function getDaysSinceLastContact(rawClient) {
  const client = ensureClientSafety(rawClient);
  const lastDate = new Date(client.lastContactDate);
  if (!Number.isFinite(lastDate.getTime())) return 0;
  const ms = Date.now() - lastDate.getTime();
  const days = Math.floor(ms / (1000 * 60 * 60 * 24));
  return Math.max(0, days);
}

function isFollowUpDue(rawClient) {
  const client = ensureClientSafety(rawClient);
  const cadenceDays = followUpDays[client.grade] || followUpDays.C;
  return getDaysSinceLastContact(client) >= cadenceDays;
}

function calculateFollowUpScore(rawClient) {
  const client = ensureClientSafety(rawClient);
  const daysSinceLastContact = getDaysSinceLastContact(client);
  const gradeWeight = gradeWeightMap[client.grade] || gradeWeightMap.C;
  const stageWeight = stageWeightMap[safeText(client.stage).toLowerCase()] ?? 0;
  const priorityBonus = safeText(client.priority).toLowerCase() === "high" ? 5 : 0;

  const score = (daysSinceLastContact * gradeWeight) + stageWeight + priorityBonus;
  return Math.round(score);
}

function enrichClient(rawClient) {
  const client = ensureClientSafety(rawClient);
  const daysSinceLastContact = getDaysSinceLastContact(client);
  const score = calculateFollowUpScore(client);
  const isDue = isFollowUpDue(client);

  return {
    ...client,
    daysSinceLastContact,
    score,
    followUpScore: score,
    isDue,
  };
}

function normalizeClientCollection(rawClients) {
  const list = Array.isArray(rawClients) ? rawClients : [];
  let changed = false;
  const normalized = list.map((raw) => {
    const before = raw || {};
    const after = ensureClientSafety(before);
    if (
      safeText(before.priority) !== safeText(after.priority) ||
      safeText(before.grade) !== safeText(after.grade) ||
      safeText(before.stage) !== safeText(after.stage) ||
      safeText(before.lastContactDate) !== safeText(after.lastContactDate) ||
      safeText(before.name) !== safeText(after.name)
    ) {
      changed = true;
    }
    return after;
  });
  return { normalized, changed };
}

function parseCommunicationLog(raw) {
  const text = safeText(raw);
  if (!text) return [];

  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  return lines.map((line) => {
    let date = todayDateISO();
    let type = "Note";
    let detail = line;

    const bracketMatch = line.match(/^\[([^\]]+)\](?:\s*\[([^\]]+)\])?\s*(.*)$/);
    if (bracketMatch) {
      const maybeDate = safeText(bracketMatch[1]);
      const maybeType = safeText(bracketMatch[2]);
      const maybeDetail = safeText(bracketMatch[3]);
      if (maybeDate) date = maybeDate;
      if (maybeType) type = maybeType;
      if (maybeDetail) detail = maybeDetail;
    }

    return {
      date,
      type,
      detail,
    };
  });
}

function historyToCommunicationLog(history) {
  const list = Array.isArray(history) ? history : [];
  return list
    .map((h) => {
      const date = safeText(h?.date || h?.timestamp || h?.createdAt || h?.updatedAt) || todayDateISO();
      const type = safeText(h?.type || h?.action || h?.event || "Note");
      const detail = safeText(h?.detail || h?.note || h?.notes || h?.description);
      if (!detail) return "";
      return `[${date}][${type}] ${detail}`;
    })
    .filter(Boolean)
    .join("\n");
}

function xmlEscape(value) {
  return safeText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function saveDB() {
  localStorage.setItem(DB_KEY, JSON.stringify(distributors));
}

function loadDB() {
  const raw = localStorage.getItem(DB_KEY);
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

async function loadInitialClients() {
  const local = loadDB();
  const localNormalized = normalizeClientCollection(local).normalized;

  if (!(window.dataAdapter && typeof window.dataAdapter.getAllClients === "function")) {
    return localNormalized;
  }

  try {
    const adapterClients = await window.dataAdapter.getAllClients();
    const adapterList = Array.isArray(adapterClients) ? adapterClients : [];
    const adapterNormalized = normalizeClientCollection(adapterList).normalized;

    if (!adapterNormalized.length) {
      return localNormalized;
    }

    const mergedById = new Map();
    const merged = [];
    const pushIfNew = (item) => {
      if (!item || typeof item !== "object") return;
      const id = safeText(item.id);
      if (id) {
        if (mergedById.has(id)) return;
        mergedById.set(id, true);
      }
      merged.push(item);
    };

    adapterNormalized.forEach(pushIfNew);
    localNormalized.forEach(pushIfNew);
    return merged;
  } catch (err) {
    console.warn("[app] Failed to read clients from dataAdapter during init; using localStorage only.", err);
    return localNormalized;
  }
}

function normalizePhone(phone) {
  let p = safeText(phone);
  p = p.replace(/\s+/g, "");
  if (p.startsWith("00")) p = "+" + p.slice(2);
  return p;
}

function openWhatsApp(phone, message) {
  const p = normalizePhone(phone).replace("+", "");
  const url = `https://wa.me/${p}?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

function openMail(email, subject, body) {
  const url = `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.open(url, "_blank");
}

function tagClass(level) {
  if (level === "A") return "tag good";
  if (level === "B") return "tag mid";
  if (level === "C") return "tag bad";
  return "tag";
}

function statusClass(status) {
  const s = safeText(status).toLowerCase();
  if (s.includes("deal") || s.includes("contract")) return "tag good";
  if (s.includes("quotation") || s.includes("sample")) return "tag mid";
  if (s.includes("lost") || s.includes("rejected")) return "tag bad";
  return "tag";
}

// ---------------------------
// Message Templates
// ---------------------------
function generateWhatsAppMessage(d) {
  const name = safeText(d.contactName) || "Sir/Madam";
  const company = safeText(d.company) || "your company";

  return `Hi ${name}, hope you're doing well.

This is Dintom from Xenta Biotech (Guangzhou). I wanted to briefly follow up and share a key POCT solution that we believe fits the local market very well.

✅ Homogeneous CLIA POCT (no separation step)
✅ Whole blood testing (no centrifugation)
✅ Room-temperature storage (2 years shelf life)
✅ Single-test cassette, very easy operation
✅ Fast turnaround (3–10 minutes)
✅ Compact & portable (only ~8.5 kg)

This platform is ideal for emergency departments, outpatient clinics, and near-patient testing scenarios.

May I know if you are currently supplying POCT analyzers to hospitals/clinics? If yes, I would be happy to share our catalog and quotation for your evaluation.`;
}

function generateEmail(d) {
  const name = safeText(d.contactName) || "Sir/Madam";
  const company = safeText(d.company) || "your company";

  const subject = `Xenta Biotech – CLIA POCT Solution for Emergency & Near-Patient Testing`;

  const body = `Dear ${name},

I hope this email finds you well.

This is Dintom Li from Xenta Biotech (Guangzhou) Medical Technology Co., Ltd. We specialize in POCT and IVD solutions, with a strong focus on emergency and near-patient diagnostics.

I would like to briefly introduce a product that we believe has strong potential for ${company}:

Our key POCT platform is a Light-Initiated Chemiluminescence (CLIA POCT) system with the following advantages:

1) Room-temperature reagent storage (up to 2 years), suitable for limited cold-chain regions  
2) Whole blood testing (no centrifugation required)  
3) Single-test cassette design, easy and standardized operation  
4) Fast turnaround time (typically 3–10 minutes)  
5) High analytical consistency compared to central-lab CLIA platforms  
6) Compact & portable analyzer (approx. 8.5 kg)

This system is well suited for emergency departments, outpatient clinics, and decentralized healthcare settings.

If you are interested, I would be happy to arrange an online meeting or provide a quotation and product catalog for your evaluation.

Best regards,
Dintom Li  
International Sales Department  
Xenta Biotech (Guangzhou) Medical Technology Co., Ltd.`;

  return { subject, body };
}

// ---------------------------
// DOM Elements
// ---------------------------
function pickEl(...selectors) {
  for (const s of selectors) {
    const node = document.querySelector(s);
    if (node) return node;
  }
  return null;
}

const el = {
  form: pickEl("#distributorForm"),
  detailForm: pickEl("#detailForm"),
  detailsPanel: pickEl("#detailsPanel"),
  detailsBackdrop: pickEl("#detailsBackdrop"),
  btnCloseDetails: pickEl("#btnCloseDetails"),
  company: pickEl("#company", "#fName"),
  country: pickEl("#country", "#fCountry"),
  city: pickEl("#city"),
  contactName: pickEl("#contactName"),
  role: pickEl("#role"),
  phone: pickEl("#phone", "#fWhatsapp"),
  email: pickEl("#email", "#fEmail"),
  website: pickEl("#website", "#fLinkedIn"),
  brands: pickEl("#brands", "#fBrands"),
  interest: pickEl("#interest", "#fFocus"),
  level: pickEl("#level", "#fGrade"),
  status: pickEl("#status", "#fStage"),
  notes: pickEl("#notes", "#fNotes"),
  communicationLog: pickEl("#fCommunicationLog"),
  priority: pickEl("#priority", "#fPriority"),
  lastContactDate: pickEl("#lastContactDate", "#fLastContact"),
  nextStep: pickEl("#fNextStep"),

  btnSave: pickEl("#btnSave"),
  btnReset: pickEl("#btnReset"),
  btnNew: pickEl("#btnNew"),
  langSelect: pickEl("#langSelect"),
  btnSortMode: pickEl("#btnSortMode"),

  search: pickEl("#search", "#qSearch"),
  filterCountry: pickEl("#filterCountry", "#qCountry"),
  filterLevel: pickEl("#filterLevel", "#qGrade"),
  filterStatus: pickEl("#filterStatus", "#qStage"),

  tableBody: pickEl("#tableBody", "#leadTbody"),
  emptyState: pickEl("#emptyState"),
  totalCount: pickEl("#totalCount", "#countPill"),
  todayTop5List: pickEl("#todayTop5List"),

  btnExport: pickEl("#btnExport"),
  btnExportWeekly: pickEl("#btnExportWeekly"),
  weeklyExportPanel: pickEl("#weeklyExportPanel"),
  weekLabelInput: pickEl("#weekLabelInput"),
  progressSourceSelect: pickEl("#progressSourceSelect"),
  manualProgressWrap: pickEl("#manualProgressWrap"),
  manualProgressInput: pickEl("#manualProgressInput"),
  existingWeeklyFile: pickEl("#existingWeeklyFile"),
  btnGenerateWeeklySummary: pickEl("#btnGenerateWeeklySummary"),
  weeklySummaryText: pickEl("#weeklySummaryText"),
  btnCopyWeeklySummary: pickEl("#btnCopyWeeklySummary"),
  btnRunWeeklyExport: pickEl("#btnRunWeeklyExport"),
  btnCloseWeeklyExport: pickEl("#btnCloseWeeklyExport"),
  btnImport: pickEl("#btnImport"),
  importFile: pickEl("#importFile", "#fileImport"),
  btnClearAll: pickEl("#btnClearAll"),

  btnQuickWhatsApp: pickEl("#btnQuickWhatsApp"),
  btnQuickEmail: pickEl("#btnQuickEmail"),

  messagePreview: pickEl("#messagePreview", "#mOutput"),
  previewTitle: pickEl("#previewTitle"),
};

function openDetailsPanel() {
  if (el.detailsPanel) {
    el.detailsPanel.classList.remove("hidden");
    el.detailsPanel.classList.add("open");
  }
  if (el.detailsBackdrop) {
    el.detailsBackdrop.classList.remove("hidden");
  }
  document.body.classList.add("details-open");
}

function closeDetailsPanel() {
  if (el.detailsPanel) {
    el.detailsPanel.classList.remove("open");
    el.detailsPanel.classList.add("hidden");
  }
  if (el.detailsBackdrop) {
    el.detailsBackdrop.classList.add("hidden");
  }
  document.body.classList.remove("details-open");
}

// ---------------------------
// CRUD
// ---------------------------
function getFormData() {
  const companyOrName = safeText(el.company ? el.company.value : "");
  const gradeValue = normalizeGrade(el.level ? el.level.value : "");
  const stageValue = safeText(el.status ? el.status.value : "") || "New";
  const lastContactDateValue = safeText(el.lastContactDate ? el.lastContactDate.value : "") || todayDateISO();
  const priorityValue = safeText(el.priority ? el.priority.value : "") || "Normal";
  const communicationLogValue = safeText(el.communicationLog ? el.communicationLog.value : "");

  return {
    company: companyOrName,
    name: companyOrName,
    country: safeText(el.country ? el.country.value : ""),
    city: safeText(el.city ? el.city.value : ""),
    contactName: safeText(el.contactName ? el.contactName.value : ""),
    role: safeText(el.role ? el.role.value : ""),
    phone: safeText(el.phone ? el.phone.value : ""),
    email: safeText(el.email ? el.email.value : ""),
    website: safeText(el.website ? el.website.value : ""),
    brands: safeText(el.brands ? el.brands.value : ""),
    interest: safeText(el.interest ? el.interest.value : ""),
    level: gradeValue,
    status: stageValue,
    grade: gradeValue,
    stage: stageValue,
    priority: priorityValue,
    lastContactDate: lastContactDateValue,
    lastContact: lastContactDateValue,
    notes: safeText(el.notes ? el.notes.value : ""),
    nextStep: safeText(el.nextStep ? el.nextStep.value : ""),
    nextAction: safeText(el.nextStep ? el.nextStep.value : ""),
    communicationLog: communicationLogValue,
    history: communicationLogValue ? parseCommunicationLog(communicationLogValue) : null,
  };
}

function resetForm() {
  const fields = [
    "company", "country", "city", "contactName", "role", "phone",
    "email", "website", "brands", "interest", "level", "status", "notes", "priority", "lastContactDate", "nextStep", "communicationLog"
  ];
  currentEditId = null;
  if (el.form) el.form.reset();
  if (!el.form) {
    fields.forEach((k) => {
      if (el[k]) el[k].value = "";
    });
  }
  if (el.lastContactDate) el.lastContactDate.value = todayDateISO();
  if (el.priority) el.priority.value = "Normal";
  if (el.detailForm) el.detailForm.classList.remove("hidden");
  if (el.emptyState) el.emptyState.classList.add("hidden");
  if (el.btnSave) el.btnSave.textContent = "Save Distributor";
  openDetailsPanel();
}

function validateData(d) {
  if (!d.company) return "Company name is required.";
  if (!d.country) return "Country is required.";
  return null;
}

function createDistributor(d) {
  const history = Array.isArray(d.history) ? d.history : [];
  const item = ensureClientSafety({
    id: uid(),
    ...d,
    history,
    createdAt: nowISO(),
    updatedAt: nowISO(),
  });
  distributors.push(item);
  saveDB();
}

function updateDistributor(id, d) {
  const idx = distributors.findIndex(x => x.id === id);
  if (idx === -1) return;

  const existing = distributors[idx] || {};
  const history = Array.isArray(d.history)
    ? d.history
    : (Array.isArray(existing.history) ? existing.history : []);

  distributors[idx] = ensureClientSafety({
    ...existing,
    ...d,
    history,
    updatedAt: nowISO(),
  });
  saveDB();
}

function deleteDistributor(id) {
  distributors = distributors.filter(x => x.id !== id);
  saveDB();
}

function editDistributor(id) {
  const d = distributors.find(x => x.id === id);
  if (!d) return;

  currentEditId = id;
  if (el.company) el.company.value = d.company || "";
  if (el.country) el.country.value = d.country || "";
  if (el.city) el.city.value = d.city || "";
  if (el.contactName) el.contactName.value = d.contactName || "";
  if (el.role) el.role.value = d.role || "";
  if (el.phone) el.phone.value = d.phone || "";
  if (el.email) el.email.value = d.email || "";
  if (el.website) el.website.value = d.website || "";
  if (el.brands) el.brands.value = d.brands || "";
  if (el.interest) el.interest.value = d.interest || "";
  if (el.level) el.level.value = normalizeGrade(d.grade || d.level);
  if (el.status) el.status.value = safeText(d.stage || d.status);
  if (el.priority) el.priority.value = safeText(d.priority || "Normal");
  if (el.lastContactDate) el.lastContactDate.value = safeText(d.lastContactDate || d.lastContact || todayDateISO());
  if (el.notes) el.notes.value = d.notes || "";
  if (el.nextStep) el.nextStep.value = safeText(d.nextStep || d.nextAction);
  if (el.communicationLog) {
    const savedLog = safeText(d.communicationLog);
    el.communicationLog.value = savedLog || historyToCommunicationLog(d.history);
  }
  if (el.detailForm) el.detailForm.classList.remove("hidden");
  if (el.emptyState) el.emptyState.classList.add("hidden");

  if (el.btnSave) el.btnSave.textContent = "Update Distributor";
  openDetailsPanel();
}

// ---------------------------
// Filtering & Sorting
// ---------------------------
function getFilteredData() {
  const searchEl = el.search || document.getElementById("search");
  const filterCountryEl = el.filterCountry || document.getElementById("filterCountry");
  const filterLevelEl = el.filterLevel || document.getElementById("filterLevel");
  const filterStatusEl = el.filterStatus || document.getElementById("filterStatus");

  const searchValue = searchEl ? searchEl.value : "";
  const filterCountryValue = filterCountryEl ? filterCountryEl.value : "";
  const filterLevelValue = filterLevelEl ? filterLevelEl.value : "";
  const filterStatusValue = filterStatusEl ? filterStatusEl.value : "";

  const q = safeText(searchValue).toLowerCase();
  const countryFilter = safeText(filterCountryValue).toLowerCase();
  const lvl = safeText(filterLevelValue);
  const st = safeText(filterStatusValue).toLowerCase();

  return distributors.filter(d => {
    const clientSafe = ensureClientSafety(d);
    const text = [
      clientSafe.name, clientSafe.company, clientSafe.country, clientSafe.city, clientSafe.region, clientSafe.contactName, clientSafe.role,
      clientSafe.phone, clientSafe.email, clientSafe.website, clientSafe.brands, clientSafe.interest, clientSafe.priority,
      clientSafe.grade, clientSafe.level, clientSafe.stage, clientSafe.status, clientSafe.notes,
      clientSafe.nextStep, clientSafe.nextAction, clientSafe.communicationLog,
      ...(Array.isArray(clientSafe.history) ? clientSafe.history.map((h) => safeText(h && (h.detail || h.note || h.notes || h.description))) : [])
    ].join(" ").toLowerCase();

    const matchQ = !q || text.includes(q);
    const matchCountry = !countryFilter || safeText(clientSafe.country).toLowerCase() === countryFilter;
    const matchLevel = !lvl || normalizeGrade(clientSafe.grade || clientSafe.level) === normalizeGrade(lvl);
    const matchStatus = !st || safeText(clientSafe.stage || clientSafe.status).toLowerCase() === st;

    return matchQ && matchCountry && matchLevel && matchStatus;
  });
}

function sortData(data) {
  const { key, direction } = sortState;
  const dir = direction === "asc" ? 1 : -1;

  return [...data].sort((a, b) => {
    if (key === "score") {
      return ((a.score || 0) - (b.score || 0)) * dir;
    }
    if (key === "lastContactDate") {
      const ta = new Date(safeText(a.lastContactDate)).getTime() || 0;
      const tb = new Date(safeText(b.lastContactDate)).getTime() || 0;
      return (ta - tb) * dir;
    }
    const va = safeText(a[key]);
    const vb = safeText(b[key]);

    if (key.includes("At")) {
      return (new Date(va) - new Date(vb)) * dir;
    }
    return va.localeCompare(vb) * dir;
  });
}

function setSort(key) {
  if (sortState.key === key) {
    sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
  } else {
    sortState.key = key;
    sortState.direction = key === "score" ? "desc" : "asc";
  }
  if (key === "score") sortMode = "score";
  if (key === "lastContactDate") sortMode = "lastContact";
  render();
}

function toggleSortMode() {
  if (sortMode === "score") {
    sortMode = "lastContact";
    sortState = { key: "lastContactDate", direction: "desc" };
  } else {
    sortMode = "score";
    sortState = { key: "score", direction: "desc" };
  }
  render();
}

// ---------------------------
// Rendering
// ---------------------------
function renderFilters(data) {
  if (!el.filterCountry) return;
  const countries = [...new Set(data.map(x => safeText(x.country)).filter(Boolean))].sort();

  const current = safeText(el.filterCountry.value);
  if (el.filterCountry.tagName === "SELECT") {
    el.filterCountry.innerHTML = `<option value="">All Countries</option>` +
      countries.map(c => `<option value="${c}">${c}</option>`).join("");

    if (countries.includes(current)) {
      el.filterCountry.value = current;
    }
  }
}

function scoreClass(score) {
  if (score >= 25) return "score-badge score-red";
  if (score >= 15) return "score-badge score-orange";
  if (score >= 8) return "score-badge score-yellow";
  return "score-badge score-gray";
}

function renderTodayTop5(data) {
  if (!el.todayTop5List) return;
  const dueTop5 = [...data]
    .filter((c) => c.isDue)
    .sort((a, b) => (b.score || 0) - (a.score || 0))
    .slice(0, 5);

  if (!dueTop5.length) {
    el.todayTop5List.innerHTML = `<div class="top5-empty">${t("top5.empty")}</div>`;
    return;
  }

  el.todayTop5List.innerHTML = dueTop5.map((c) => `
    <div class="top5-item">
      <div class="top5-main">${safeText(c.name) || "-"}</div>
      <div class="top5-meta">${safeText(c.country) || "-"} | ${safeText(c.stage) || "-"} | ${c.daysSinceLastContact}d</div>
    </div>
  `).join("");
}

function updateSortModeButton() {
  if (!el.btnSortMode) return;
  el.btnSortMode.textContent = sortMode === "score" ? t("actions.sortByLastContact") : t("actions.sortByScore");
}

function renderTable(data) {
  if (!el.tableBody) return;
  el.tableBody.innerHTML = "";

  if (!data.length) {
    if (el.emptyState) el.emptyState.classList.remove("hidden");
    return;
  }

  if (el.emptyState) el.emptyState.classList.add("hidden");

  data.forEach(client => {
    const tr = document.createElement("tr");
    tr.dataset.id = client.id;
    const name = safeText(client.name || client.company);
    const country = safeText(client.country);
    const region = safeText(client.region || client.city);
    const stage = safeText(client.stage || client.status);
    const grade = normalizeGrade(client.grade || client.level);
    const lastContact = safeText(client.lastContactDate || client.lastContact || client.updatedAt);
    const dueHtml = client.isDue ? `<span class="due-flag" title="${t("status.due")}">&#9888; <span class="due-badge">${t("status.due")}</span></span>` : "";

    tr.innerHTML = `
      <td>${name || "-"} ${dueHtml}</td>
      <td>${country || "-"}</td>
      <td>${region || "-"}</td>
      <td>${stage || "-"}</td>
      <td>${grade || "-"}</td>
      <td class="muted">${lastContact || "-"}</td>
      <td><span class="${scoreClass(client.score)}">${client.score}</span></td>
      <td>
        <div class="row-actions">
          <button class="btn ghost" data-action="wa" data-id="${client.id}">WhatsApp</button>
          <button class="btn ghost" data-action="mail" data-id="${client.id}">Email</button>
          <button class="btn ghost" data-action="edit" data-id="${client.id}">Edit</button>
          <button class="btn danger" data-action="del" data-id="${client.id}">Delete</button>
        </div>
      </td>
    `;

    el.tableBody.appendChild(tr);
  });
}

function renderStats(data) {
  if (el.totalCount) el.totalCount.textContent = data.length;
}

function render() {
  const safeAll = distributors.map(enrichClient);
  const filtered = getFilteredData().map(enrichClient);
  const sorted = sortData(filtered);

  renderFilters(safeAll);
  renderStats(filtered);
  renderTodayTop5(safeAll);
  updateSortModeButton();
  renderTable(sorted);
}

// ---------------------------
// Import / Export
// ---------------------------
async function getExportClients() {
  if (window.dataAdapter && typeof window.dataAdapter.getAllClients === "function") {
    try {
      const clients = await window.dataAdapter.getAllClients();
      if (Array.isArray(clients)) return clients;
    } catch (err) {
      console.warn("[app] Failed to load clients from dataAdapter; using in-memory records.", err);
    }
  }
  return distributors;
}

function getProgressUpdatesByClient(clients, options = {}) {
  const source = safeText(options.progressSource || "latest").toLowerCase();
  const manualMap = parseManualMapping(options.manualMappingText || "");

  const pickField = (d, keys) => {
    for (const key of keys) {
      const value = safeText(d[key]);
      if (value) return value;
    }
    return "";
  };

  const updates = new Map();
  clients.forEach((d) => {
    const id = safeText(d.id || d.clientId);
    if (!id) return;

    let progressText = "";
    if (source === "weekly") {
      progressText = pickField(d, [
        "weeklyNotes",
        "weekly_notes",
        "weeklyProgress",
        "weekly_progress",
        "weekNotes",
        "notesWeekly",
        "weeklyUpdate",
        "weekly_update",
      ]);
    } else if (source === "manual") {
      progressText = safeText(manualMap.get(id));
    } else {
      progressText = pickField(d, [
        "latestProgress",
        "latest_progress",
        "progress",
        "currentProgress",
        "status",
        "stage",
      ]);
    }

    if (safeText(progressText)) {
      updates.set(id, safeText(progressText));
    }
  });

  return updates;
}

function getDefaultWeekLabel() {
  return `Week ${new Date().toISOString().slice(0, 10)}`;
}

function parseManualMapping(raw) {
  const map = new Map();
  safeText(raw)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const parts = line.split("\t");
      if (parts.length < 2) return;
      const id = safeText(parts[0]);
      const update = safeText(parts.slice(1).join("\t"));
      if (!id || !update) return;
      map.set(id, update);
    });
  return map;
}

function sanitizeFilenameToken(value) {
  const cleaned = safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_");
  return cleaned || "Week";
}

function toggleWeeklyPanel(visible) {
  if (!el.weeklyExportPanel) return;
  el.weeklyExportPanel.classList.toggle("hidden", !visible);
}

function syncManualMappingVisibility() {
  if (!el.progressSourceSelect || !el.manualProgressWrap) return;
  const showManual = el.progressSourceSelect.value === "manual";
  el.manualProgressWrap.classList.toggle("hidden", !showManual);
}

function parseWeekDate(label) {
  const m = safeText(label).match(/^(?:Week|周)\s+(\d{4}-\d{2}-\d{2})$/i);
  if (!m) return null;
  const d = new Date(`${m[1]}T00:00:00Z`);
  if (!Number.isFinite(d.getTime())) return null;
  return d;
}

function getClientId(client) {
  return safeText(client.id || client.clientId);
}

function getClientName(client) {
  return safeText(client.contactName || client.name || client.clientName || client.company || "(Unknown)");
}

function getClientCountry(client) {
  return safeText(client.country || "-");
}

function getClientStage(client) {
  return safeText(client.stage || client.status);
}

function getClientPriority(client) {
  return safeText(client.priority || client.level);
}

async function readUploadedClientsHistory(file) {
  if (!file) return null;
  if (!window.XLSX) {
    console.warn("[app] SheetJS not found; history analysis skipped.");
    return null;
  }

  try {
    const data = await file.arrayBuffer();
    const wb = window.XLSX.read(data, { type: "array" });
    const ws = wb.Sheets["Clients"];
    if (!ws) return null;

    const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows.length) return null;

    const headers = rows[0].map((h) => safeText(h));
    const idIdx = headers.findIndex((h) => h === "Client ID" || h === "客户ID");
    if (idIdx === -1) return null;

    const weeklyColumns = headers.filter((h) => /^Week\s+/i.test(h) || /^周\s*/.test(h) || /^第.*周$/.test(h));
    const rowsById = new Map();

    rows.slice(1).forEach((r) => {
      const id = safeText(r[idIdx]);
      if (!id) return;
      const rowObj = {};
      headers.forEach((h, i) => {
        rowObj[h] = safeText(r[i]);
      });
      rowsById.set(id, rowObj);
    });

    return { weeklyColumns, rowsById };
  } catch (err) {
    console.warn("[app] Failed reading uploaded history workbook.", err);
    return null;
  }
}

function buildWeeklySummaryText(options) {
  const weekLabel = safeText(options.weekLabel || getDefaultWeekLabel());
  const clients = Array.isArray(options.clients) ? options.clients : [];
  const updatesById = options.updatesById instanceof Map ? options.updatesById : new Map();
  const history = options.history || null;

  const crmById = new Map();
  clients.forEach((c) => {
    const id = getClientId(c);
    if (!id) return;
    crmById.set(id, c);
  });

  const updatedLines = [];
  Array.from(updatesById.entries()).forEach(([id, update]) => {
    const c = crmById.get(id) || {};
    updatedLines.push(`- ${getClientName(c)} (${getClientCountry(c)}): ${safeText(update)}`);
  });
  updatedLines.sort((a, b) => a.localeCompare(b));

  const stuckLines = [];
  if (history && history.weeklyColumns && history.weeklyColumns.length >= 2) {
    const weekCols = [...history.weeklyColumns];
    if (!weekCols.includes(weekLabel)) weekCols.push(weekLabel);
    weekCols.sort((a, b) => {
      const da = parseWeekDate(a);
      const db = parseWeekDate(b);
      if (!da || !db) return a.localeCompare(b);
      return da - db;
    });

    const lastTwo = weekCols.slice(-2);
    const allIds = new Set([
      ...Array.from(history.rowsById.keys()),
      ...Array.from(crmById.keys()),
    ]);

    Array.from(allIds).forEach((id) => {
      const row = history.rowsById.get(id) || {};
      const hasRecent = lastTwo.some((wk) => {
        if (wk === weekLabel) return safeText(updatesById.get(id));
        return safeText(row[wk]);
      });
      if (hasRecent) return;

      const c = crmById.get(id) || {};
      const name = safeText(
        (crmById.has(id) ? getClientName(c) : "") ||
        pickField(row, ["Client Name", "客户名称", "Organization", "机构名称"]) ||
        id
      );
      const country = safeText((crmById.has(id) ? getClientCountry(c) : "") || pickField(row, ["Country", "国家"]) || "-");
      stuckLines.push(`- ${name} (${country})`);
    });
    stuckLines.sort((a, b) => a.localeCompare(b));
  }

  const reminderLines = [];
  clients.forEach((c) => {
    const priority = getClientPriority(c);
    const stage = getClientStage(c);
    const isPriority = priority.toLowerCase() === "high";
    const isNegotiation = stage.toLowerCase() === "negotiation";
    if (!isPriority && !isNegotiation) return;
    reminderLines.push(`- ${getClientName(c)} (${getClientCountry(c)}) | Priority: ${priority || "-"} | Stage: ${stage || "-"}`);
  });
  reminderLines.sort((a, b) => a.localeCompare(b));

  return [
    `${weekLabel} Summary`,
    "",
    "Updated clients:",
    ...(updatedLines.length ? updatedLines : ["- None"]),
    "",
    "Stuck clients (no update for 2+ weeks):",
    ...(stuckLines.length ? stuckLines : [history ? "- None" : "- History not available (upload last report to enable)."]),
    "",
    "High priority reminders:",
    ...(reminderLines.length ? reminderLines : ["- None"]),
  ].join("\n");
}

async function generateWeeklySummaryText() {
  const weekLabel = el.weekLabelInput ? el.weekLabelInput.value : getDefaultWeekLabel();
  const progressSource = el.progressSourceSelect ? el.progressSourceSelect.value : "latest";
  const manualMappingText = el.manualProgressInput ? el.manualProgressInput.value : "";
  const existingFile = el.existingWeeklyFile && el.existingWeeklyFile.files ? el.existingWeeklyFile.files[0] : null;

  const clients = await getExportClients();
  const updatesById = getProgressUpdatesByClient(clients, {
    progressSource,
    manualMappingText,
  });
  const history = await readUploadedClientsHistory(existingFile);
  const summaryText = buildWeeklySummaryText({
    weekLabel,
    clients,
    updatesById,
    history,
  });

  if (el.weeklySummaryText) {
    el.weeklySummaryText.value = summaryText;
  }
}

async function exportWeeklyProgressExcel(options = {}) {
  try {
    const weekLabel = safeText(options.weekLabel || getDefaultWeekLabel());
    const progressSource = safeText(options.progressSource || "latest");
    const manualMappingText = options.manualMappingText || "";
    const existingFile = options.existingFile || null;

    const clients = await getExportClients();
    const enrichedClients = clients.map(enrichClient);
    const updatesById = getProgressUpdatesByClient(clients, {
      progressSource,
      manualMappingText,
    });
    if (window.excelExport && typeof window.excelExport.exportWeeklyWorkbook === "function") {
      await window.excelExport.exportWeeklyWorkbook({
        clients: enrichedClients,
        updatesById,
        weekLabel,
        existingFile,
        includeSummary: true,
      });
      return;
    }

    console.warn("[app] excelExport module missing; falling back to JSON export.");
    const blob = new Blob([JSON.stringify(enrichedClients, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Weekly_Client_Progress_${sanitizeFilenameToken(weekLabel)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    console.warn("[app] Weekly export failed.", err);
    alert("Weekly Excel export failed.");
  }
}

async function exportJSON() {
  try {
    const clients = await getExportClients();
    const blob = new Blob([JSON.stringify(clients, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `xenta_distributors_${new Date().toISOString().slice(0,10)}.json`;
    a.click();

    URL.revokeObjectURL(url);
  } catch (err) {
    console.warn("[app] JSON export failed.", err);
    alert("JSON export failed.");
  }
}

function importJSON(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      if (!Array.isArray(data)) {
        alert("Invalid JSON format: should be an array.");
        return;
      }

      // minimal validation
      const cleaned = data.map(x => ensureClientSafety({
        ...(x || {}),
        id: x.id || uid(),
        company: safeText(x.company),
        name: safeText(x.name || x.company),
        country: safeText(x.country),
        city: safeText(x.city),
        contactName: safeText(x.contactName),
        role: safeText(x.role),
        phone: safeText(x.phone),
        email: safeText(x.email),
        website: safeText(x.website),
        brands: safeText(x.brands),
        interest: safeText(x.interest),
        level: normalizeGrade(x.level || x.grade),
        status: safeText(x.status || x.stage),
        grade: normalizeGrade(x.grade || x.level),
        stage: safeText(x.stage || x.status) || "New",
        priority: safeText(x.priority) || "Normal",
        lastContactDate: safeText(x.lastContactDate || x.lastContact || x.updatedAt) || todayDateISO(),
        lastContact: safeText(x.lastContact || x.lastContactDate || x.updatedAt) || todayDateISO(),
        notes: safeText(x.notes),
        nextStep: safeText(x.nextStep || x.nextAction || x.next_step || x.plan),
        nextAction: safeText(x.nextAction || x.nextStep || x.next_step || x.plan),
        communicationLog: safeText(x.communicationLog),
        history: Array.isArray(x.history) ? x.history : parseCommunicationLog(x.communicationLog),
        createdAt: x.createdAt || nowISO(),
        updatedAt: x.updatedAt || nowISO(),
      }));

      distributors = cleaned;
      saveDB();
      render();
      alert("Import successful.");
    } catch (e) {
      alert("Failed to import JSON. Invalid file.");
    }
  };
  reader.readAsText(file);
}

function clearAll() {
  if (!confirm("Are you sure you want to delete ALL records? This cannot be undone.")) return;
  distributors = [];
  saveDB();
  render();
}

// ---------------------------
// Event Listeners
// ---------------------------
function handleSave(e) {
  if (e && typeof e.preventDefault === "function") e.preventDefault();

  const data = getFormData();
  const err = validateData(data);
  if (err) {
    alert(err);
    return;
  }

  if (currentEditId) {
    updateDistributor(currentEditId, data);
  } else {
    createDistributor(data);
  }

  resetForm();
  render();
  closeDetailsPanel();
}

if (el.form) {
  el.form.addEventListener("submit", handleSave);
}

if (el.btnSave && !el.form) {
  el.btnSave.addEventListener("click", handleSave);
}

if (el.btnNew) {
  el.btnNew.addEventListener("click", (e) => {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    resetForm();
  });
}

if (el.btnReset) {
  el.btnReset.addEventListener("click", () => {
    resetForm();
  });
}

if (el.btnCloseDetails) {
  el.btnCloseDetails.addEventListener("click", () => {
    closeDetailsPanel();
  });
}

if (el.detailsBackdrop) {
  el.detailsBackdrop.addEventListener("click", () => {
    closeDetailsPanel();
  });
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeDetailsPanel();
});

if (el.langSelect) {
  el.langSelect.addEventListener("change", (e) => {
    setLanguage(e.target ? e.target.value : "en", { persist: true, rerender: true });
  });
}

if (el.btnSortMode) {
  el.btnSortMode.addEventListener("click", () => {
    toggleSortMode();
  });
}

if (el.search) el.search.addEventListener("input", render);
if (el.filterCountry) el.filterCountry.addEventListener("change", render);
if (el.filterLevel) el.filterLevel.addEventListener("change", render);
if (el.filterStatus) el.filterStatus.addEventListener("change", render);

if (el.btnExport) el.btnExport.addEventListener("click", exportJSON);
if (el.btnExportWeekly) {
  el.btnExportWeekly.addEventListener("click", () => {
    if (el.weekLabelInput && !safeText(el.weekLabelInput.value)) {
      el.weekLabelInput.value = getDefaultWeekLabel();
    }
    syncManualMappingVisibility();
    toggleWeeklyPanel(true);
  });
}

if (el.btnCloseWeeklyExport) {
  el.btnCloseWeeklyExport.addEventListener("click", () => toggleWeeklyPanel(false));
}

if (el.progressSourceSelect) {
  el.progressSourceSelect.addEventListener("change", syncManualMappingVisibility);
}

if (el.btnRunWeeklyExport) {
  el.btnRunWeeklyExport.addEventListener("click", () => {
    exportWeeklyProgressExcel({
      weekLabel: el.weekLabelInput ? el.weekLabelInput.value : "",
      progressSource: el.progressSourceSelect ? el.progressSourceSelect.value : "latest",
      manualMappingText: el.manualProgressInput ? el.manualProgressInput.value : "",
      existingFile: el.existingWeeklyFile && el.existingWeeklyFile.files ? el.existingWeeklyFile.files[0] : null,
    });
  });
}

if (el.btnGenerateWeeklySummary) {
  el.btnGenerateWeeklySummary.addEventListener("click", () => {
    generateWeeklySummaryText();
  });
}

if (el.btnCopyWeeklySummary && el.weeklySummaryText) {
  el.btnCopyWeeklySummary.addEventListener("click", async () => {
    const text = safeText(el.weeklySummaryText.value);
    if (!text) {
      alert("No summary text to copy.");
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      alert("Summary copied.");
    } catch (_) {
      el.weeklySummaryText.select();
      document.execCommand("copy");
      alert("Summary copied.");
    }
  });
}

if (el.btnImport && el.importFile) {
  el.btnImport.addEventListener("click", () => {
    el.importFile.click();
  });
}

if (el.importFile) {
  el.importFile.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    importJSON(file);
    e.target.value = "";
  });
}

if (el.btnClearAll) el.btnClearAll.addEventListener("click", clearAll);

// Table actions
if (el.tableBody) {
  el.tableBody.addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn) {
      const row = e.target.closest("tr[data-id]");
      if (!row) return;
      editDistributor(row.dataset.id);
      return;
    }

    const action = btn.dataset.action;
    const id = btn.dataset.id;

    const d = distributors.find(x => x.id === id);
    if (!d) return;

    if (action === "edit") {
      editDistributor(id);
    }

    if (action === "del") {
      if (!confirm(`Delete ${d.company}?`)) return;
      deleteDistributor(id);
      render();
    }

    if (action === "wa") {
      if (!d.phone) {
        alert("No phone number available.");
        return;
      }
      const msg = generateWhatsAppMessage(d);
      openWhatsApp(d.phone, msg);
    }

    if (action === "mail") {
      if (!d.email) {
        alert("No email available.");
        return;
      }
      const email = generateEmail(d);
      openMail(d.email, email.subject, email.body);
    }
  });
}

// Sort by clicking table headers
document.querySelectorAll("th[data-sort]").forEach(th => {
  th.addEventListener("click", () => {
    setSort(th.dataset.sort);
  });
});

// Quick message preview buttons
if (el.btnQuickWhatsApp && el.previewTitle && el.messagePreview) {
  el.btnQuickWhatsApp.addEventListener("click", () => {
    const data = getFormData();
    const msg = generateWhatsAppMessage(data);

    el.previewTitle.textContent = "WhatsApp Message Preview";
    el.messagePreview.value = msg;
  });
}

if (el.btnQuickEmail && el.previewTitle && el.messagePreview) {
  el.btnQuickEmail.addEventListener("click", () => {
    const data = getFormData();
    const email = generateEmail(data);

    el.previewTitle.textContent = "Email Preview";
    el.messagePreview.value = `Subject: ${email.subject}\n\n${email.body}`;
  });
}

// ---------------------------
// Init
// ---------------------------
async function init() {
  const initialLang = detectInitialLanguage();
  setLanguage(initialLang, { persist: false, rerender: false });
  if (el.langSelect) el.langSelect.value = currentLang;

  const loaded = await loadInitialClients();
  const normalizedResult = normalizeClientCollection(loaded);
  distributors = normalizedResult.normalized;
  saveDB();
  if (el.weekLabelInput && !safeText(el.weekLabelInput.value)) {
    el.weekLabelInput.value = getDefaultWeekLabel();
  }
  closeDetailsPanel();
  updateSortModeButton();
  syncManualMappingVisibility();
  if (el.tableBody) {
    render();
  }
}

init();
  const pickField = (row, keys) => {
    if (!row || typeof row !== "object") return "";
    for (const key of keys) {
      const value = safeText(row[key]);
      if (value) return value;
    }
    return "";
  };
