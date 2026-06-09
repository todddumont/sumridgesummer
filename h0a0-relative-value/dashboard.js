const FILES = {
  bonds: "output/bond_rv_latest.csv",
  issuers: "output/issuer_rv_latest.csv"
};

const state = {
  bonds: [],
  issuers: []
};

const formatNumber = (value, decimals = 0) => {
  const num = Number(value);
  if (!Number.isFinite(num)) return value || "";
  return num.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
};

const firstValue = (row, keys) => {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") return row[key];
  }
  return "";
};

function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      i++;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i++;
      row.push(cell);
      if (row.some(v => v.trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  row.push(cell);
  if (row.some(v => v.trim() !== "")) rows.push(row);
  if (rows.length < 2) return [];

  const headers = rows[0].map(h => h.trim());
  return rows.slice(1).map(values => {
    const obj = {};
    headers.forEach((header, index) => { obj[header] = (values[index] || "").trim(); });
    return obj;
  });
}

async function fetchCSV(path) {
  const response = await fetch(path, { cache: "no-store" });
  if (!response.ok) throw new Error(`${path}: ${response.status}`);
  return parseCSV(await response.text());
}

async function loadFromOutputFolder() {
  setStatus("Loading CSVs...");
  const [bonds, issuers] = await Promise.all([
    fetchCSV(FILES.bonds),
    fetchCSV(FILES.issuers)
  ]);
  state.bonds = bonds;
  state.issuers = issuers;
  renderAll();
  setStatus("Data loaded");
}

function setStatus(text) {
  document.getElementById("loadStatus").textContent = text;
}

function setView(view) {
  document.querySelectorAll(".nav-item").forEach(btn => btn.classList.toggle("active", btn.dataset.view === view));
  document.querySelectorAll(".view-panel").forEach(panel => panel.classList.add("hidden"));
  document.getElementById(`${view}View`).classList.remove("hidden");
}

function renderMetrics() {
  const latestDate = firstValue(state.bonds[0] || state.issuers[0] || {}, ["as_of_date", "date"]);
  document.getElementById("metricAsOf").textContent = latestDate || "--";
  document.getElementById("metricBonds").textContent = formatNumber(state.bonds.length);
  document.getElementById("metricIssuers").textContent = formatNumber(state.issuers.length);
}

function scoreClass(signalText) {
  const s = String(signalText || "").toLowerCase();
  if (s.includes("cheap")) return "cheap";
  if (s.includes("rich")) return "rich";
  return "";
}

function getScore(row) {
  return Number(firstValue(row, ["rv_score", "issuer_rv_score", "score", "curve_pickup_score"])) || 0;
}


function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function couponText(row, prefix = "") {
  return firstValue(row, [prefix + "coupon", prefix + "coupon_rate", prefix + "cpn", "coupon", "coupon_rate", "cpn"]);
}

function maturityText(row, prefix = "") {
  return firstValue(row, [prefix + "maturity", prefix + "maturity_date", prefix + "maturity_dt", "maturity", "maturity_date", "maturity_dt"]);
}

function cusipText(row, prefix = "") {
  return firstValue(row, [prefix + "cusip", prefix + "cusip_number", prefix + "isin", "cusip", "cusip_number", "isin"]);
}

function shortMaturity(value) {
  const s = String(value || "").trim();
  if (!s) return "";
  const m = s.match(/(20\d{2}|19\d{2})/);
  return m ? m[1] : s;
}

function formatCoupon(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const n = Number(raw);
  if (!Number.isFinite(n)) return raw;
  const pct = n > 20 ? n / 100 : n;
  return pct.toFixed(pct % 1 === 0 ? 0 : 3).replace(/0+$/, "").replace(/\.$/, "");
}

function bondIdentity(row) {
  const cpn = formatCoupon(couponText(row));
  const mat = shortMaturity(maturityText(row));
  const cusip = cusipText(row);
  const bits = [];
  if (cpn || mat) bits.push(`${cpn || "?"} ${mat || "?"}`.trim());
  if (cusip) bits.push(`CUSIP ${cusip}`);
  return bits.join(" | ");
}

function confidenceLevel(row) {
  const peer = Number(firstValue(row, ["peer_count", "peer_group_count", "peer_count_used", "issuer_peer_count", "peer_issuer_count"])) || 0;
  const obs = Number(firstValue(row, ["history_obs", "bond_history_obs", "oas_obs", "oas_observations", "issuer_history_obs", "issuer_obs_count"])) || 0;
  const bondCount = Number(firstValue(row, ["bond_count", "issuer_bond_count"])) || 0;
  const curveResidual = firstValue(row, ["issuer_curve_residual", "curve_residual"]);
  const hasCurve = curveResidual !== "" && Number.isFinite(Number(curveResidual));

  let points = 0;
  if (peer >= 25) points += 2;
  else if (peer >= 10) points += 1;

  if (obs >= 180) points += 2;
  else if (obs >= 60) points += 1;

  if (bondCount >= 3 || hasCurve) points += 1;

  if (points >= 4) return { label: "High", cls: "high", detail: peer ? `${peer} peers` : "strong support" };
  if (points >= 2) return { label: "Medium", cls: "medium", detail: peer ? `${peer} peers` : "moderate support" };
  return { label: "Low", cls: "low", detail: peer ? `${peer} peers` : "limited support" };
}

function renderMiniList(id, rows, type) {
  const el = document.getElementById(id);
  if (!rows.length) {
    el.innerHTML = `<div class="mini-sub">No data loaded.</div>`;
    return;
  }
  el.innerHTML = rows.slice(0, 5).map(row => {
    const ticker = firstValue(row, ["ticker", "issuer", "company"]);
    const desc = firstValue(row, ["description", "issuer", "sector_l3"]);
    const signal = firstValue(row, ["rv_signal", "issuer_rv_signal", "signal"]);
    const confidence = confidenceLevel(row);
    const idLine = bondIdentity(row) || "CUSIP n/a";
    const oas = firstValue(row, ["oas"]);
    const score = getScore(row);
    return `
      <div class="mini-row">
        <div>
          <div class="mini-main">${escapeHtml(ticker || desc || "--")}</div>
          <div class="mini-sub">${escapeHtml(idLine)}${oas ? ` · OAS ${escapeHtml(formatNumber(oas, 0))}` : ""}</div>
          <div class="mini-sub">${escapeHtml(desc || signal || "")}</div>
          <div class="mini-sub">RV Score ${escapeHtml(formatNumber(score, 0))}</div>
        </div>
        <span class="confidence-badge ${confidence.cls}" title="${escapeHtml(confidence.detail)}">${confidence.label}</span>
      </div>`;
  }).join("");
}

function renderOverview() {
  const cheapBonds = [...state.bonds]
    .filter(r => String(firstValue(r, ["rv_signal"])).includes("Cheap"))
    .sort((a, b) => getScore(b) - getScore(a));
  const richBonds = [...state.bonds]
    .filter(r => String(firstValue(r, ["rv_signal"])).includes("Rich"))
    .sort((a, b) => getScore(a) - getScore(b));
  renderMiniList("overviewCheap", cheapBonds, "cheap");
  renderMiniList("overviewRich", richBonds, "rich");
}

function tableHTML(rows, columns) {
  if (!rows.length) return `<thead><tr><th>No data</th></tr></thead><tbody><tr><td>No rows loaded.</td></tr></tbody>`;
  const header = `<thead><tr>${columns.map(c => `<th>${escapeHtml(c.label)}</th>`).join("")}</tr></thead>`;
  const body = rows.map(row => `<tr>${columns.map(c => {
    const raw = c.get ? c.get(row) : firstValue(row, c.keys || [c.key]);
    const value = c.num ? formatNumber(raw, c.decimals || 0) : raw;
    const cls = [c.num ? "num" : "", c.signal ? `signal ${scoreClass(raw).replace(" ", "-")}` : ""].filter(Boolean).join(" ");
    return `<td class="${cls}">${c.html ? value : escapeHtml(value)}</td>`;
  }).join("")}</tr>`).join("");
  return `${header}<tbody>${body}</tbody>`;
}

function filterRows(rows, query, signalFilter = "") {
  const q = query.trim().toLowerCase();
  return rows.filter(row => {
    const matchesQ = !q || Object.values(row).some(v => String(v).toLowerCase().includes(q));
    const signal = firstValue(row, ["rv_signal", "issuer_rv_signal", "signal"]);
    const matchesSignal = !signalFilter || signal === signalFilter;
    return matchesQ && matchesSignal;
  });
}

function renderBondTable() {
  const rows = filterRows(state.bonds, document.getElementById("bondSearch").value, document.getElementById("bondSignalFilter").value)
    .sort((a, b) => getScore(b) - getScore(a))
    .slice(0, 500);
  const cols = [
    { label: "Ticker", keys: ["ticker"] },
    { label: "Description", keys: ["description"] },
    { label: "Signal", keys: ["rv_signal"], signal: true },
    { label: "Score", keys: ["rv_score"], num: true },
    { label: "Rating", keys: ["rating"] },
    { label: "Sector", keys: ["sector_l3", "sector"] },
    { label: "Price", keys: ["price"], num: true, decimals: 2 },
    { label: "OAS", keys: ["oas"], num: true },
    { label: "Bond OAS %", keys: ["bond_oas_percentile", "oas_pct_1y"], num: true },
    { label: "Peer Residual", keys: ["peer_oas_residual"], num: true },
    { label: "Curve Residual", keys: ["issuer_curve_residual"], num: true },
    { label: "AI Note", keys: ["ai_note", "rv_note"] }
  ];
  document.getElementById("bondTable").innerHTML = tableHTML(rows, cols);
}


function sameIssuerBondRows(issuerRow) {
  const issuerTicker = normalizeText(firstValue(issuerRow, ["ticker", "issuer", "company"]));
  if (!issuerTicker) return [];
  return state.bonds.filter(b => normalizeText(firstValue(b, ["ticker", "issuer", "company"])) === issuerTicker);
}

function bondCusipLine(row) {
  const desc = firstValue(row, ["description", "bond", "security"]);
  const cusip = cusipText(row) || "CUSIP n/a";
  const cpn = formatCoupon(couponText(row));
  const mat = shortMaturity(maturityText(row));
  const oas = firstValue(row, ["oas"]);
  const score = firstValue(row, ["rv_score"]);
  const bondBits = [];
  if (cpn || mat) bondBits.push(`${cpn || "?"} ${mat || "?"}`.trim());
  if (cusip) bondBits.push(cusip.startsWith("CUSIP") ? cusip : `CUSIP ${cusip}`);
  if (oas) bondBits.push(`OAS ${formatNumber(oas, 0)}`);
  if (score) bondBits.push(`Score ${formatNumber(score, 0)}`);
  const firstLine = bondBits.length ? bondBits.join(" | ") : (cusip || "--");
  return `<div class="bond-cusip-line"><strong>${escapeHtml(firstLine)}</strong>${desc ? `<span>${escapeHtml(desc)}</span>` : ""}</div>`;
}

function issuerSpecificBonds(issuerRow, side) {
  const bonds = sameIssuerBondRows(issuerRow);
  if (!bonds.length) return `<span class="muted">No current bonds matched</span>`;
  let filtered = bonds;
  if (side === "cheap") {
    filtered = bonds.filter(b => String(firstValue(b, ["rv_signal"])).toLowerCase().includes("cheap"));
    if (!filtered.length) filtered = bonds;
    filtered = filtered.sort((a, b) => getScore(b) - getScore(a));
  } else {
    filtered = bonds.filter(b => String(firstValue(b, ["rv_signal"])).toLowerCase().includes("rich"));
    if (!filtered.length) filtered = bonds;
    filtered = filtered.sort((a, b) => getScore(a) - getScore(b));
  }
  return filtered.slice(0, 3).map(bondCusipLine).join("");
}

function renderIssuerTable() {
  const rows = filterRows(state.issuers, document.getElementById("issuerSearch").value)
    .sort((a, b) => getScore(b) - getScore(a))
    .slice(0, 500);
  const cols = [
    { label: "Ticker", keys: ["ticker", "issuer"] },
    { label: "Signal", keys: ["issuer_rv_signal", "rv_signal"], signal: true },
    { label: "Score", keys: ["issuer_rv_score", "rv_score"], num: true },
    { label: "Bond Count", keys: ["bond_count"], num: true },
    { label: "Rating Bucket", keys: ["rating_bucket"] },
    { label: "Sector", keys: ["sector_l3", "sector"] },
    { label: "Issuer OAS", keys: ["issuer_oas", "issuer_median_oas"], num: true },
    { label: "Peer Residual", keys: ["issuer_peer_residual"], num: true },
    { label: "1Y %", keys: ["issuer_1y_percentile", "issuer_oas_pct_1y"], num: true },
    { label: "Vs Median", keys: ["issuer_oas_vs_1y_median"], num: true },
    { label: "Cheap Bond CUSIPs", get: row => issuerSpecificBonds(row, "cheap"), html: true },
    { label: "Rich Bond CUSIPs", get: row => issuerSpecificBonds(row, "rich"), html: true },
    { label: "Note", keys: ["ai_note", "rv_note"] }
  ];
  document.getElementById("issuerTable").innerHTML = tableHTML(rows, cols);
}

function renderAll() {
  renderMetrics();
  renderOverview();
  renderBondTable();
  renderIssuerTable();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function fileToRows(file) {
  if (!file) return [];
  return parseCSV(await file.text());
}

async function loadUploadedFiles() {
  state.bonds = await fileToRows(document.getElementById("bondFile").files[0]);
  state.issuers = await fileToRows(document.getElementById("issuerFile").files[0]);
  renderAll();
  setStatus("Uploaded files loaded");
}

function init() {
  document.querySelectorAll(".nav-item").forEach(btn => btn.addEventListener("click", () => setView(btn.dataset.view)));
  document.getElementById("reloadBtn").addEventListener("click", () => loadFromOutputFolder().catch(err => {
    console.error(err);
    setStatus("CSV fetch failed");
    document.getElementById("uploadPanel").classList.remove("hidden");
  }));
  document.getElementById("openUploadBtn").addEventListener("click", () => document.getElementById("uploadPanel").classList.toggle("hidden"));
  document.getElementById("loadUploadBtn").addEventListener("click", loadUploadedFiles);
  ["bondSearch", "bondSignalFilter"].forEach(id => document.getElementById(id).addEventListener("input", renderBondTable));
  document.getElementById("issuerSearch").addEventListener("input", renderIssuerTable);

  loadFromOutputFolder().catch(err => {
    console.warn("Automatic CSV load failed", err);
    setStatus("Use manual upload or localhost");
    document.getElementById("uploadPanel").classList.remove("hidden");
  });
}

init();
