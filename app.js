/**************************************/

const DB_KEY = "xenta_crm_distributors_v1";

let distributors = [];
let currentEditId = null;

let sortState = { key: "updatedAt", direction: "desc" };

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
const el = {
  form: document.querySelector("#distributorForm"),
  company: document.querySelector("#company"),
  country: document.querySelector("#country"),
  city: document.querySelector("#city"),
  contactName: document.querySelector("#contactName"),
  role: document.querySelector("#role"),
  phone: document.querySelector("#phone"),
  email: document.querySelector("#email"),
  website: document.querySelector("#website"),
  brands: document.querySelector("#brands"),
  interest: document.querySelector("#interest"),
  level: document.querySelector("#level"),
  status: document.querySelector("#status"),
  notes: document.querySelector("#notes"),

  btnSave: document.querySelector("#btnSave"),
  btnReset: document.querySelector("#btnReset"),

  search: document.querySelector("#search"),
  filterCountry: document.querySelector("#filterCountry"),
  filterLevel: document.querySelector("#filterLevel"),
  filterStatus: document.querySelector("#filterStatus"),

  tableBody: document.querySelector("#tableBody"),
  emptyState: document.querySelector("#emptyState"),
  totalCount: document.querySelector("#totalCount"),

  btnExport: document.querySelector("#btnExport"),
  btnImport: document.querySelector("#btnImport"),
  importFile: document.querySelector("#importFile"),
  btnClearAll: document.querySelector("#btnClearAll"),

  btnQuickWhatsApp: document.querySelector("#btnQuickWhatsApp"),
  btnQuickEmail: document.querySelector("#btnQuickEmail"),

  messagePreview: document.querySelector("#messagePreview"),
  previewTitle: document.querySelector("#previewTitle"),
};

// ---------------------------
// CRUD
// ---------------------------
function getFormData() {
  return {
    company: safeText(el.company.value),
    country: safeText(el.country.value),
    city: safeText(el.city.value),
    contactName: safeText(el.contactName.value),
    role: safeText(el.role.value),
    phone: safeText(el.phone.value),
    email: safeText(el.email.value),
    website: safeText(el.website.value),
    brands: safeText(el.brands.value),
    interest: safeText(el.interest.value),
    level: safeText(el.level.value),
    status: safeText(el.status.value),
    notes: safeText(el.notes.value),
  };
}

function resetForm() {
  currentEditId = null;
  el.form.reset();
  el.btnSave.textContent = "Save Distributor";
}

function validateData(d) {
  if (!d.company) return "Company name is required.";
  if (!d.country) return "Country is required.";
  return null;
}

function createDistributor(d) {
  const item = {
    id: uid(),
    ...d,
    createdAt: nowISO(),
    updatedAt: nowISO(),
  };
  distributors.push(item);
  saveDB();
}

function updateDistributor(id, d) {
  const idx = distributors.findIndex(x => x.id === id);
  if (idx === -1) return;

  distributors[idx] = {
    ...distributors[idx],
    ...d,
    updatedAt: nowISO(),
  };
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
  el.company.value = d.company || "";
  el.country.value = d.country || "";
  el.city.value = d.city || "";
  el.contactName.value = d.contactName || "";
  el.role.value = d.role || "";
  el.phone.value = d.phone || "";
  el.email.value = d.email || "";
  el.website.value = d.website || "";
  el.brands.value = d.brands || "";
  el.interest.value = d.interest || "";
  el.level.value = d.level || "";
  el.status.value = d.status || "";
  el.notes.value = d.notes || "";

  el.btnSave.textContent = "Update Distributor";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// ---------------------------
// Filtering & Sorting
// ---------------------------
function getFilteredData() {
  const q = safeText(el.search.value).toLowerCase();
  const c = safeText(el.filterCountry.value).toLowerCase();
  const lvl = safeText(el.filterLevel.value);
  const st = safeText(el.filterStatus.value).toLowerCase();

  return distributors.filter(d => {
    const text = [
      d.company, d.country, d.city, d.contactName, d.role,
      d.phone, d.email, d.website, d.brands, d.interest,
      d.level, d.status, d.notes
    ].join(" ").toLowerCase();

    const matchQ = !q || text.includes(q);
    const matchCountry = !c || safeText(d.country).toLowerCase() === c;
    const matchLevel = !lvl || safeText(d.level) === lvl;
    const matchStatus = !st || safeText(d.status).toLowerCase() === st;

    return matchQ && matchCountry && matchLevel && matchStatus;
  });
}

function sortData(data) {
  const { key, direction } = sortState;
  const dir = direction === "asc" ? 1 : -1;

  return [...data].sort((a, b) => {
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
    sortState.direction = "asc";
  }
  render();
}

// ---------------------------
// Rendering
// ---------------------------
function renderFilters(data) {
  const countries = [...new Set(data.map(x => safeText(x.country)).filter(Boolean))].sort();

  const current = el.filterCountry.value;
  el.filterCountry.innerHTML = `<option value="">All Countries</option>` +
    countries.map(c => `<option value="${c}">${c}</option>`).join("");

  if (countries.includes(current)) {
    el.filterCountry.value = current;
  }
}

function renderTable(data) {
  el.tableBody.innerHTML = "";

  if (!data.length) {
    el.emptyState.classList.remove("hidden");
    return;
  }

  el.emptyState.classList.add("hidden");

  data.forEach(d => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td><strong>${d.company || "-"}</strong><br><span class="muted">${d.city || ""}</span></td>
      <td>${d.country || "-"}</td>
      <td>${d.contactName || "-"}<br><span class="muted">${d.role || ""}</span></td>
      <td>${d.phone || "-"}</td>
      <td>${d.email || "-"}</td>
      <td>${d.brands || "-"}</td>
      <td><span class="${tagClass(d.level)}">${d.level || "-"}</span></td>
      <td><span class="${statusClass(d.status)}">${d.status || "-"}</span></td>
      <td class="muted">${formatDate(d.updatedAt)}</td>
      <td>
        <div class="row-actions">
          <button class="btn ghost" data-action="wa" data-id="${d.id}">WhatsApp</button>
          <button class="btn ghost" data-action="mail" data-id="${d.id}">Email</button>
          <button class="btn ghost" data-action="edit" data-id="${d.id}">Edit</button>
          <button class="btn danger" data-action="del" data-id="${d.id}">Delete</button>
        </div>
      </td>
    `;

    el.tableBody.appendChild(tr);
  });
}

function renderStats(data) {
  el.totalCount.textContent = data.length;
}

function render() {
  const filtered = getFilteredData();
  const sorted = sortData(filtered);

  renderFilters(distributors);
  renderStats(filtered);
  renderTable(sorted);
}

// ---------------------------
// Import / Export
// ---------------------------
function exportJSON() {
  const blob = new Blob([JSON.stringify(distributors, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `xenta_distributors_${new Date().toISOString().slice(0,10)}.json`;
  a.click();

  URL.revokeObjectURL(url);
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
      const cleaned = data.map(x => ({
        id: x.id || uid(),
        company: safeText(x.company),
        country: safeText(x.country),
        city: safeText(x.city),
        contactName: safeText(x.contactName),
        role: safeText(x.role),
        phone: safeText(x.phone),
        email: safeText(x.email),
        website: safeText(x.website),
        brands: safeText(x.brands),
        interest: safeText(x.interest),
        level: safeText(x.level),
        status: safeText(x.status),
        notes: safeText(x.notes),
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
el.form.addEventListener("submit", (e) => {
  e.preventDefault();

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
});

el.btnReset.addEventListener("click", () => {
  resetForm();
});

el.search.addEventListener("input", render);
el.filterCountry.addEventListener("change", render);
el.filterLevel.addEventListener("change", render);
el.filterStatus.addEventListener("change", render);

el.btnExport.addEventListener("click", exportJSON);

el.btnImport.addEventListener("click", () => {
  el.importFile.click();
});

el.importFile.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  importJSON(file);
  e.target.value = "";
});

el.btnClearAll.addEventListener("click", clearAll);

// Table actions
el.tableBody.addEventListener("click", (e) => {
  const btn = e.target.closest("button");
  if (!btn) return;

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

// Sort by clicking table headers
document.querySelectorAll("th[data-sort]").forEach(th => {
  th.addEventListener("click", () => {
    setSort(th.dataset.sort);
  });
});

// Quick message preview buttons
el.btnQuickWhatsApp.addEventListener("click", () => {
  const data = getFormData();
  const msg = generateWhatsAppMessage(data);

  el.previewTitle.textContent = "WhatsApp Message Preview";
  el.messagePreview.value = msg;
});

el.btnQuickEmail.addEventListener("click", () => {
  const data = getFormData();
  const email = generateEmail(data);

  el.previewTitle.textContent = "Email Preview";
  el.messagePreview.value = `Subject: ${email.subject}\n\n${email.body}`;
});

// ---------------------------
// Init
// ---------------------------
function init() {
  distributors = loadDB();
  render();
}

init();