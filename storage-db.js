// Minimal IndexedDB wrapper (no external deps)
// Stores: leads

const DB_NAME = "xenta_crm";
const DB_VER = 1;

let db = null;

export async function dbOpen() {
  if (db) return db;

  db = await new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VER);

    req.onupgradeneeded = (e) => {
      const d = req.result;

      if (!d.objectStoreNames.contains("leads")) {
        const store = d.createObjectStore("leads", { keyPath: "id" });
        store.createIndex("by_updated_at", "updated_at", { unique: false });
        store.createIndex("by_email", "email", { unique: false });
        store.createIndex("by_phone", "phone", { unique: false });
        store.createIndex("by_key_hint", "key_hint", { unique: false });
      }
    };

    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });

  return db;
}

function tx(storeName, mode = "readonly") {
  if (!db) throw new Error("DB not opened. Call dbOpen() first.");
  return db.transaction(storeName, mode).objectStore(storeName);
}

function uid() {
  // sortable-ish id
  return `L_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

async function findExistingLeadIdByKeyHint(key_hint) {
  return await new Promise((resolve, reject) => {
    const store = tx("leads", "readonly");
    const idx = store.index("by_key_hint");
    const req = idx.get(key_hint);

    req.onsuccess = () => resolve(req.result?.id || null);
    req.onerror = () => reject(req.error);
  });
}

export async function upsertLead(lead) {
  if (!lead) throw new Error("lead is required");
  await dbOpen();

  const now = new Date().toISOString();
  const key_hint = lead.key_hint || lead.email || lead.phone || `${lead.company}:${lead.name}`;

  let id = lead.id || null;

  if (!id && key_hint) {
    id = await findExistingLeadIdByKeyHint(key_hint);
  }
  if (!id) id = uid();

  const record = {
    id,
    key_hint,
    created_at: lead.created_at || now,
    updated_at: now,
    ...lead,
  };

  await new Promise((resolve, reject) => {
    const store = tx("leads", "readwrite");
    const req = store.put(record);
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });

  return record;
}
