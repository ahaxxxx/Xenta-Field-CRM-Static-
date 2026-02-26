(function () {
  const TAG = "[cloudSync]";

  function safeText(value) {
    return (value || "").toString().trim();
  }

  function normalizeClient(raw) {
    if (!raw || typeof raw !== "object") return null;
    const client = { ...raw };
    if (!safeText(client.id)) return null;
    return client;
  }

  function normalizeConfig(raw) {
    const cfg = raw && typeof raw === "object" ? raw : {};
    const url = safeText(cfg.url);
    const anonKey = safeText(cfg.anonKey);
    const table = safeText(cfg.table) || "crm_clients";
    const enabled = Boolean(url && anonKey && window.supabase && typeof window.supabase.createClient === "function");
    return { url, anonKey, table, enabled };
  }

  const cfg = normalizeConfig(window.XENTA_SUPABASE_CONFIG);
  let client = null;

  if (cfg.enabled) {
    client = window.supabase.createClient(cfg.url, cfg.anonKey, {
      auth: { persistSession: false },
    });
    console.info(`${TAG} enabled with table "${cfg.table}".`);
  } else {
    console.info(`${TAG} disabled. Fill js/cloud-config.js to enable.`);
  }

  async function getAllClients() {
    if (!cfg.enabled || !client) return [];

    const { data, error } = await client
      .from(cfg.table)
      .select("id,payload,updated_at")
      .order("updated_at", { ascending: false });

    if (error) throw error;
    if (!Array.isArray(data)) return [];

    return data
      .map((row) => {
        const payload = row && typeof row.payload === "object" ? row.payload : {};
        const normalized = normalizeClient({
          ...payload,
          id: safeText(row.id || payload.id),
          updatedAt: safeText(payload.updatedAt || row.updated_at),
        });
        return normalized;
      })
      .filter(Boolean);
  }

  async function upsertClient(rawClient) {
    if (!cfg.enabled || !client) return;
    const normalized = normalizeClient(rawClient);
    if (!normalized) return;

    const updatedAt = safeText(normalized.updatedAt) || new Date().toISOString();
    const row = {
      id: normalized.id,
      payload: { ...normalized, updatedAt },
      updated_at: updatedAt,
    };

    const { error } = await client.from(cfg.table).upsert(row, { onConflict: "id" });
    if (error) throw error;
  }

  async function deleteClient(id) {
    if (!cfg.enabled || !client) return;
    const idSafe = safeText(id);
    if (!idSafe) return;

    const { error } = await client.from(cfg.table).delete().eq("id", idSafe);
    if (error) throw error;
  }

  async function replaceAllClients(rawClients) {
    if (!cfg.enabled || !client) return;
    const input = Array.isArray(rawClients) ? rawClients : [];
    const clients = input.map(normalizeClient).filter(Boolean);

    const { error: deleteError } = await client.from(cfg.table).delete().not("id", "is", null);
    if (deleteError) throw deleteError;

    if (!clients.length) return;

    const rows = clients.map((clientRow) => {
      const updatedAt = safeText(clientRow.updatedAt) || new Date().toISOString();
      return {
        id: clientRow.id,
        payload: { ...clientRow, updatedAt },
        updated_at: updatedAt,
      };
    });

    const { error: insertError } = await client.from(cfg.table).upsert(rows, { onConflict: "id" });
    if (insertError) throw insertError;
  }

  window.cloudSync = {
    enabled: cfg.enabled,
    table: cfg.table,
    getAllClients,
    upsertClient,
    deleteClient,
    replaceAllClients,
  };
})();

