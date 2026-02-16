// Heuristic text extraction for business cards / contact notes.
// You can iterate as you see more PDFs.

function uniq(arr) {
  return Array.from(new Set(arr.filter(Boolean)));
}

function pickBestName(lines) {
  // Try to find a likely "person name" line:
  // - not containing @
  // - not too long
  // - contains letters and maybe title
  const candidates = lines
    .map(s => s.trim())
    .filter(s => s.length >= 3 && s.length <= 40)
    .filter(s => !s.includes("@"))
    .filter(s => !/https?:\/\//i.test(s))
    .filter(s => !/\b(inc|ltd|llc|company|medical|laboratory|clinic|hospital)\b/i.test(s))
    .filter(s => /[A-Za-z]/.test(s));

  // Often the name is one of the first 3–5 non-empty lines
  return candidates[0] || "";
}

function pickTitle(text) {
  const titles = [
    "CEO", "Founder", "Co-Founder", "Managing Director", "Director",
    "Head of Sales", "Sales Manager", "Business Development", "BD",
    "Service Engineer", "Sales Engineer", "Product Manager", "Manager",
    "Dr", "Professor"
  ];

  for (const t of titles) {
    const re = new RegExp(`\\b${t.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&")}\\b`, "i");
    if (re.test(text)) return t;
  }

  // Try generic pattern "Head of ..."
  const m = text.match(/\b(Head of [A-Za-z ]{2,30})\b/i);
  if (m) return m[1].trim();

  return "";
}

function pickCompany(lines, website) {
  // If we have website, infer from domain
  if (website) {
    try {
      const u = new URL(website.startsWith("http") ? website : `https://${website}`);
      const host = u.hostname.replace(/^www\./, "");
      const main = host.split(".")[0];
      if (main && main.length >= 3) return main.toUpperCase();
    } catch {}
  }

  // Else: look for lines with company-ish keywords
  const candidates = lines
    .map(s => s.trim())
    .filter(Boolean)
    .filter(s => !s.includes("@"))
    .filter(s => !/https?:\/\//i.test(s))
    .filter(s => s.length <= 60)
    .filter(s =>
      /\b(ltd|llc|inc|company|medical|lab|laboratory|diagnostic|pharma|health|clinic)\b/i.test(s)
    );

  return candidates[0] || "";
}

function pickWebsite(text) {
  const m = text.match(/(https?:\/\/[^\s]+)|(www\.[^\s]+)|([a-z0-9.-]+\.[a-z]{2,})(\/[^\s]*)?/i);
  if (!m) return "";
  const url = m[1] || m[2] || m[3] || "";
  if (!url) return "";
  if (url.includes("@")) return "";
  // avoid grabbing random "something.pdf"
  if (/\.pdf\b/i.test(url)) return "";
  return url.startsWith("http") ? url : `https://${url.replace(/^www\./, "www.")}`;
}

function pickEmail(text) {
  const m = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : "";
}

function pickPhone(text) {
  // Liberal phone extraction, then keep the longest
  const matches = text.match(/(\+?\d[\d\s\-()]{6,}\d)/g);
  if (!matches || !matches.length) return "";
  const cleaned = matches
    .map(s => s.replace(/\s+/g, " ").trim())
    .map(s => s.replace(/[^\d+()\- ]/g, "").trim());
  cleaned.sort((a, b) => b.length - a.length);
  return cleaned[0] || "";
}

function guessCountry(text) {
  // Keep minimal; you can extend.
  const map = [
    ["UAE", /\b(UAE|Dubai|Abu Dhabi)\b/i],
    ["Kenya", /\bKenya\b/i],
    ["Ethiopia", /\bEthiopia|Addis Ababa\b/i],
    ["Rwanda", /\bRwanda|Kigali\b/i],
    ["Uganda", /\bUganda|Kampala\b/i],
    ["Tanzania", /\bTanzania|Dar es salaam\b/i],
  ];
  for (const [c, re] of map) {
    if (re.test(text)) return c;
  }
  return "";
}

export function extractLeadFromText(text) {
  const raw = (text || "").replace(/\r/g, "\n");
  const lines = raw.split("\n").map(s => s.trim()).filter(Boolean);

  const email = pickEmail(raw);
  const website = pickWebsite(raw);
  const phone = pickPhone(raw);
  const name = pickBestName(lines);
  const title = pickTitle(raw);
  const company = pickCompany(lines, website);
  const country = guessCountry(raw);

  const tags = uniq([
    /chemilum/i.test(raw) ? "CLIA" : "",
    /\bPCR\b/i.test(raw) ? "PCR" : "",
    /\bdistributor\b/i.test(raw) ? "Distributor" : "",
    /\bhospital\b/i.test(raw) ? "Hospital" : "",
    /\blab\b|\blaboratory\b/i.test(raw) ? "Lab" : "",
  ]);

  return {
    name,
    title,
    company,
    country,
    email,
    phone,
    website,
    tags,
    notes: "",
  };
}
