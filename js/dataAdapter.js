(function () {
  const ADAPTER_TAG = "[dataAdapter]";
  const LOCAL_KEY_HINTS = [
    "xenta_crm_distributors_v1",
    "clients",
    "crm_clients",
    "xenta_clients",
    "xenta_crm_clients",
    "crm_leads",
    "leads",
    "distributors",
  ];

  const RECOMMENDED_EXPORT_FIELDS = [
    "id",
    "company",
    "country",
    "city",
    "contactName",
    "role",
    "email",
    "phone",
    "website",
    "brands",
    "interest",
    "level",
    "status",
    "notes",
    "createdAt",
    "updatedAt",
  ];

  function logWarn(message, extra) {
    if (extra !== undefined) {
      console.warn(`${ADAPTER_TAG} ${message}`, extra);
      return;
    }
    console.warn(`${ADAPTER_TAG} ${message}`);
  }

  function safeParseArray(raw) {
    if (!raw) return null;
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {
      return null;
    }
    return null;
  }

  function getLikelyLocalStorageKeys() {
    const all = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const k = localStorage.key(i);
      if (!k) continue;
      all.push(k);
    }

    const dynamic = all.filter((k) => /(client|crm|lead|distributor)/i.test(k));
    const merged = [...LOCAL_KEY_HINTS, ...dynamic];
    return Array.from(new Set(merged));
  }

  function detectClientsFromLocalStorage() {
    const keys = getLikelyLocalStorageKeys();

    for (const key of keys) {
      const arr = safeParseArray(localStorage.getItem(key));
      if (!arr || !arr.length) continue;
      if (typeof arr[0] !== "object" || arr[0] === null) continue;
      return { source: "localStorage", key, clients: arr };
    }

    return { source: "localStorage", key: null, clients: [] };
  }

  async function fetchTextIfExists(path) {
    try {
      const res = await fetch(path, { cache: "no-store" });
      if (!res.ok) return "";
      return await res.text();
    } catch (_) {
      return "";
    }
  }

  async function detectIndexedDbConfigFromCode() {
    const scriptPaths = Array.from(document.querySelectorAll("script[src]"))
      .map((n) => n.getAttribute("src"))
      .filter(Boolean);
    const candidatePaths = Array.from(new Set([
      ...scriptPaths,
      "storage-db.js",
      "app.js",
      "js/storage-db.js",
      "js/app.js",
    ]));

    let dbName = "";
    let storeName = "";

    for (const path of candidatePaths) {
      const code = await fetchTextIfExists(path);
      if (!code) continue;

      const dbMatch = code.match(/const\s+DB_NAME\s*=\s*["']([^"']+)["']/);
      const storeMatch = code.match(/createObjectStore\s*\(\s*["']([^"']+)["']/);

      if (!dbName && dbMatch) dbName = dbMatch[1];
      if (!storeName && storeMatch) storeName = storeMatch[1];

      if (dbName && storeName) break;
    }

    if (!dbName || !storeName) {
      logWarn("Unable to detect IndexedDB config from code constants. Using placeholder fallback.");
      return {
        dbName: "__CRM_DB_NAME__",
        storeName: "__CRM_OBJECT_STORE__",
        isPlaceholder: true,
      };
    }

    return { dbName, storeName, isPlaceholder: false };
  }

  async function readAllFromIndexedDb() {
    if (!("indexedDB" in window)) return { source: "indexedDB", clients: [] };

    const cfg = await detectIndexedDbConfigFromCode();
    if (cfg.isPlaceholder) {
      return { source: "indexedDB", clients: [], config: cfg };
    }

    return await new Promise((resolve) => {
      const openReq = indexedDB.open(cfg.dbName);

      openReq.onerror = () => {
        logWarn(`Failed opening IndexedDB "${cfg.dbName}".`, openReq.error);
        resolve({ source: "indexedDB", clients: [], config: cfg });
      };

      openReq.onsuccess = () => {
        const db = openReq.result;
        if (!db.objectStoreNames.contains(cfg.storeName)) {
          logWarn(`IndexedDB store "${cfg.storeName}" not found in "${cfg.dbName}".`);
          db.close();
          resolve({ source: "indexedDB", clients: [], config: cfg });
          return;
        }

        const tx = db.transaction(cfg.storeName, "readonly");
        const store = tx.objectStore(cfg.storeName);
        const getAllReq = store.getAll();

        getAllReq.onerror = () => {
          logWarn("Failed reading clients from IndexedDB.", getAllReq.error);
          db.close();
          resolve({ source: "indexedDB", clients: [], config: cfg });
        };

        getAllReq.onsuccess = () => {
          const result = Array.isArray(getAllReq.result) ? getAllReq.result : [];
          db.close();
          resolve({ source: "indexedDB", clients: result, config: cfg });
        };
      };
    });
  }

  function mergeClients(localClients, idbClients) {
    if (!idbClients.length) return localClients;
    if (!localClients.length) return idbClients;

    const out = [];
    const seen = new Set();
    const add = (client) => {
      if (!client || typeof client !== "object") return;
      const k = client.id || client.key_hint || `${client.company || ""}:${client.name || client.contactName || ""}`;
      if (!k || seen.has(k)) return;
      seen.add(k);
      out.push(client);
    };

    idbClients.forEach(add);
    localClients.forEach(add);
    return out;
  }

  async function getAllClients() {
    const local = detectClientsFromLocalStorage();
    const idb = await readAllFromIndexedDb();
    return mergeClients(local.clients, idb.clients);
  }

  async function getClientById(id) {
    if (!id) return null;
    const all = await getAllClients();
    return all.find((c) => c && c.id === id) || null;
  }

  async function getClientSchema() {
    const all = await getAllClients();
    const keySet = new Set();

    all.forEach((c) => {
      if (!c || typeof c !== "object") return;
      Object.keys(c).forEach((k) => keySet.add(k));
    });

    return {
      keysFound: Array.from(keySet).sort(),
      recommendedExportFields: RECOMMENDED_EXPORT_FIELDS,
    };
  }

  window.dataAdapter = {
    getAllClients,
    getClientById,
    getClientSchema,
  };

  // Optional direct aliases for convenience.
  window.getAllClients = getAllClients;
  window.getClientById = getClientById;
  window.getClientSchema = getClientSchema;
})();
