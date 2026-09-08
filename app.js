"use strict";

/* ==========================================================================
   IndexedDB helpers — this is where the local SQLite database bytes and
   nothing else gets persisted between visits. All local to this browser.
   ========================================================================== */

const IDB_NAME = "IndicatorExtractorDB";
const IDB_STORE = "files";

function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbGet(key) {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function idbSet(key, value) {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

/* ==========================================================================
   Local config (indicators, cloud settings) — plain localStorage, since
   this is a real standalone site, not a Claude artifact sandbox.
   ========================================================================== */

const INDICATORS_KEY = "indicator_extractor_indicators";
const CLOUD_KEY = "indicator_extractor_cloud_config";

function loadIndicators() {
  const raw = localStorage.getItem(INDICATORS_KEY);
  if (raw) return JSON.parse(raw);
  return [
    { name: "Completion Rate", synonyms: ["Completion %", "Completion Percentage"] },
    { name: "Budget Utilized", synonyms: ["Budget Used", "Budget Spent"] },
  ];
}
function saveIndicators(list) {
  localStorage.setItem(INDICATORS_KEY, JSON.stringify(list));
}

function loadCloudConfig() {
  const raw = localStorage.getItem(CLOUD_KEY);
  return raw ? JSON.parse(raw) : {
    clientId: "", tenantId: "", shareLink: "",
    worksheet: "Sheet1", tableName: "IndicatorData",
    driveId: "", itemId: "",
  };
}
function saveCloudConfig(cfg) {
  localStorage.setItem(CLOUD_KEY, JSON.stringify(cfg));
}

/* ==========================================================================
   Small utilities
   ========================================================================== */

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function toCsv(headers, rows) {
  const esc = v => `"${String(v ?? "").replace(/"/g, '""')}"`;
  return [headers.map(esc).join(","), ...rows.map(r => r.map(esc).join(","))].join("\r\n");
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function downloadCsv(filename, headers, rows) {
  downloadBlob(filename, new Blob([toCsv(headers, rows)], { type: "text/csv" }));
}

/* ==========================================================================
   Pyodide boot
   ========================================================================== */

let pyodide = null;
const py = {};

async function persistDbToIndexedDb() {
  const bytes = pyodide.FS.readFile("/data/indicator_data.db");
  await idbSet("indicator_data.db", bytes);
}

async function boot() {
  const bootStatus = document.getElementById("boot-status");
  try {
    pyodide = await loadPyodide();

    bootStatus.textContent = "Installing the document-reading library…";
    await pyodide.loadPackage(["lxml", "micropip", "sqlite3"]);
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("python-docx");

    bootStatus.textContent = "Loading the extraction engine…";
    const src = await (await fetch("extractor.py")).text();
    pyodide.runPython(src);

    pyodide.FS.mkdirTree("/data");
    pyodide.FS.mkdirTree("/uploads");

    bootStatus.textContent = "Restoring your saved data…";
    const saved = await idbGet("indicator_data.db");
    if (saved) {
      pyodide.FS.writeFile("/data/indicator_data.db", new Uint8Array(saved));
    }

    py.init_db = pyodide.globals.get("init_db");
    py.extract = pyodide.globals.get("extract_from_docx_json");
    py.extract_lines = pyodide.globals.get("extract_from_lines_json");
    py.save_docx_images = pyodide.globals.get("save_docx_images_json");
    py.insert_rows = pyodide.globals.get("insert_rows");
    py.fetch_all = pyodide.globals.get("fetch_all_rows_json");
    py.fetch_unpushed = pyodide.globals.get("fetch_unpushed_rows_json");
    py.mark_pushed = pyodide.globals.get("mark_pushed");
    py.distinct = pyodide.globals.get("distinct_values_json");
    py.row_count = pyodide.globals.get("row_count");
    py.build_pivot = pyodide.globals.get("build_pivot_json");
    py.clear_all = pyodide.globals.get("clear_all_data");
    py.suggested_period = pyodide.globals.get("suggested_period_value");

    py.init_db();

    bootStatus.style.display = "none";
    document.getElementById("tabs").style.display = "flex";
    document.getElementById("app").style.display = "block";

    initUI();
  } catch (err) {
    bootStatus.textContent = "Failed to load: " + err.message +
      " — check your internet connection (this page needs it once, to load the Python engine) and reload.";
    console.error(err);
  }
}

/* ==========================================================================
   UI wiring
   ========================================================================== */

const PIVOT_FIELD_KEYS = [
  "Indicator", "Document", "Population Type", "Gender",
  "Household/Individual", "Reporting Period", "Source",
];

function initUI() {
  wireTabs();
  wireIndicatorsTab();
  wireExtractTab();
  wireDataTab();
  wireBuildTab();
  wireCloudTab();
  wireBackupButtons();
}

function wireTabs() {
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
      if (btn.dataset.tab === "data") refreshDataTab();
    });
  });
}

/* ---- Indicators tab ---- */

function renderIndicatorsTable() {
  const list = loadIndicators();
  const tbody = document.querySelector("#indicatorsTable tbody");
  tbody.innerHTML = list.map((ind, i) => `
    <tr>
      <td>${escapeHtml(ind.name)}</td>
      <td>${escapeHtml((ind.synonyms || []).join(", "))}</td>
      <td><button class="secondary small" data-remove="${i}">Remove</button></td>
    </tr>
  `).join("");
  tbody.querySelectorAll("[data-remove]").forEach(btn => {
    btn.addEventListener("click", () => {
      const list2 = loadIndicators();
      list2.splice(Number(btn.dataset.remove), 1);
      saveIndicators(list2);
      renderIndicatorsTable();
    });
  });
}

function wireIndicatorsTab() {
  renderIndicatorsTable();
  document.getElementById("addIndicatorBtn").addEventListener("click", () => {
    const name = document.getElementById("indName").value.trim();
    if (!name) { alert("Enter an indicator name."); return; }
    const synonyms = document.getElementById("indSynonyms").value
      .split(",").map(s => s.trim()).filter(Boolean);
    const list = loadIndicators();
    list.push({ name, synonyms });
    saveIndicators(list);
    document.getElementById("indName").value = "";
    document.getElementById("indSynonyms").value = "";
    renderIndicatorsTable();
  });
}

/* ---- Extract tab ---- */

/* ---- OCR (Tesseract.js — runs fully in-browser, no server involved) ---- */

const IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif"];

let tesseractWorker = null;
async function getTesseractWorker() {
  if (!tesseractWorker) {
    tesseractWorker = await Tesseract.createWorker("eng");
  }
  return tesseractWorker;
}

async function ocrImageToLines(fileOrBlob) {
  const worker = await getTesseractWorker();
  const { data: { text } } = await worker.recognize(fileOrBlob);
  return text.split(/\r?\n/);
}

function guessImageMime(path) {
  const ext = path.split(".").pop().toLowerCase();
  const map = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", bmp: "image/bmp", gif: "image/gif", tif: "image/tiff", tiff: "image/tiff" };
  return map[ext] || "image/png";
}

function getFileExt(filename) {
  const parts = filename.split(".");
  return parts.length > 1 ? "." + parts.pop().toLowerCase() : "";
}

function wireExtractTab() {
  const fileInput = document.getElementById("fileInput");
  const fileList = document.getElementById("fileList");
  const periodType = document.getElementById("periodType");
  const periodValue = document.getElementById("periodValue");

  fileInput.addEventListener("change", () => {
    fileList.textContent = fileInput.files.length
      ? `${fileInput.files.length} file(s) selected: ` + Array.from(fileInput.files).map(f => f.name).join(", ")
      : "";
  });

  periodType.addEventListener("change", () => {
    if (periodType.value === "All") {
      periodValue.value = "";
      periodValue.disabled = true;
    } else {
      periodValue.disabled = false;
    }
  });

  document.getElementById("useCurrentPeriod").addEventListener("click", () => {
    if (periodType.value === "All") return;
    periodValue.value = py.suggested_period(periodType.value);
  });

  document.getElementById("processBtn").addEventListener("click", async () => {
    const statusEl = document.getElementById("processStatus");
    const files = fileInput.files;
    if (!files.length) { alert("Add at least one .docx or image file first."); return; }
    const indicators = loadIndicators();
    if (!indicators.length) { alert("Add at least one indicator in the Indicators tab first."); return; }

    const ocrEnabled = document.getElementById("ocrImagesCheckbox").checked;
    const pType = periodType.value;
    const pValue = pType === "All" ? "" : periodValue.value.trim();

    statusEl.textContent = "Processing…";
    let allResults = [];
    try {
      for (const file of files) {
        const ext = getFileExt(file.name);
        const isImage = IMAGE_EXTENSIONS.includes(ext);

        if (ext === ".docx") {
          const buf = await file.arrayBuffer();
          const path = "/uploads/" + file.name;
          pyodide.FS.writeFile(path, new Uint8Array(buf));

          statusEl.textContent = `Reading ${file.name}…`;
          const results = JSON.parse(py.extract(path, JSON.stringify(indicators)));
          results.forEach(r => { r.period_type = pType; r.period_value = pValue; });
          allResults = allResults.concat(results);

          if (ocrEnabled) {
            const imagePaths = JSON.parse(py.save_docx_images(path));
            for (const imgPath of imagePaths) {
              try {
                statusEl.textContent = `OCR'ing an image embedded in ${file.name}…`;
                const bytes = pyodide.FS.readFile(imgPath);
                const blob = new Blob([bytes], { type: guessImageMime(imgPath) });
                const lines = await ocrImageToLines(blob);
                const imgResults = JSON.parse(py.extract_lines(
                  JSON.stringify(lines), JSON.stringify(indicators), file.name, "image in document"
                ));
                imgResults.forEach(r => { r.period_type = pType; r.period_value = pValue; });
                allResults = allResults.concat(imgResults);
              } catch (imgErr) {
                console.warn(`Skipping an embedded image in ${file.name} (OCR error):`, imgErr);
              }
            }
          }
        } else if (isImage) {
          if (!ocrEnabled) {
            statusEl.textContent = `Skipping ${file.name} — check "OCR images" to process image files.`;
            continue;
          }
          statusEl.textContent = `OCR'ing ${file.name}…`;
          const lines = await ocrImageToLines(file);
          const results = JSON.parse(py.extract_lines(
            JSON.stringify(lines), JSON.stringify(indicators), file.name, "image"
          ));
          results.forEach(r => { r.period_type = pType; r.period_value = pValue; });
          allResults = allResults.concat(results);
        } else {
          console.warn("Unsupported file type, skipping:", file.name);
        }
      }

      if (allResults.length) {
        py.insert_rows(JSON.stringify(allResults));
        await persistDbToIndexedDb();
      }

      renderResultsTable(allResults);
      statusEl.textContent = `Done. ${allResults.length} indicator value(s) found this run. ` +
        `Total dataset: ${py.row_count()} row(s).`;
    } catch (err) {
      statusEl.textContent = "Error: " + err.message;
      console.error(err);
    }
  });
}

function renderResultsTable(rows) {
  const tbody = document.querySelector("#resultsTable tbody");
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${escapeHtml(r.document)}</td>
      <td>${escapeHtml(r.indicator)}</td>
      <td>${escapeHtml(r.value)}</td>
      <td>${escapeHtml(r.population_type)}</td>
      <td>${escapeHtml(r.gender)}</td>
      <td>${escapeHtml(r.unit)}</td>
      <td>${escapeHtml(periodDisplay(r.period_type, r.period_value))}</td>
      <td>${escapeHtml(r.source)}</td>
    </tr>
  `).join("");
}

function periodDisplay(periodType, periodValue) {
  if (!periodType || periodType === "All") return "All";
  return periodValue ? `${periodValue} (${periodType})` : `(${periodType})`;
}

/* ---- Data tab ---- */

function wireDataTab() {
  document.getElementById("refreshDataBtn").addEventListener("click", refreshDataTab);
  document.getElementById("dataFilterIndicator").addEventListener("change", refreshDataTab);
  document.getElementById("exportAllCsvBtn").addEventListener("click", () => {
    const rows = JSON.parse(py.fetch_all(JSON.stringify(currentDataFilters())));
    const headers = ["document", "indicator", "value", "population_type", "gender", "unit",
      "source", "period_type", "period_value", "extracted_at"];
    downloadCsv("indicator_data_export.csv", headers,
      rows.map(r => headers.map(h => r[h])));
  });
  document.getElementById("resetDataBtn").addEventListener("click", async () => {
    if (!confirm("This permanently deletes all extracted data stored in this browser. Continue?")) return;
    py.clear_all();
    await persistDbToIndexedDb();
    refreshDataTab();
  });
  refreshDataTab();
}

function currentDataFilters() {
  const val = document.getElementById("dataFilterIndicator").value;
  return val && val !== "(All)" ? { indicator: val } : {};
}

function refreshDataTab() {
  const indicatorSelect = document.getElementById("dataFilterIndicator");
  const currentVal = indicatorSelect.value || "(All)";
  const distinctIndicators = JSON.parse(py.distinct("indicator"));
  indicatorSelect.innerHTML = ["(All)", ...distinctIndicators]
    .map(v => `<option ${v === currentVal ? "selected" : ""}>${escapeHtml(v)}</option>`).join("");

  const rows = JSON.parse(py.fetch_all(JSON.stringify(currentDataFilters())));
  document.getElementById("dataCount").textContent = `${rows.length} row(s) shown, ${py.row_count()} total in dataset.`;

  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${escapeHtml(r.document)}</td>
      <td>${escapeHtml(r.indicator)}</td>
      <td>${escapeHtml(r.value)}</td>
      <td>${escapeHtml(r.population_type)}</td>
      <td>${escapeHtml(r.gender)}</td>
      <td>${escapeHtml(r.unit)}</td>
      <td>${escapeHtml(periodDisplay(r.period_type, r.period_value))}</td>
      <td>${escapeHtml(r.source)}</td>
      <td>${escapeHtml(r.extracted_at)}</td>
    </tr>
  `).join("");
}

/* ---- Build Table tab ---- */

let lastPivot = null;

function wireBuildTab() {
  const rowSel = document.getElementById("pivotRowField");
  const colSel = document.getElementById("pivotColField");
  rowSel.innerHTML = PIVOT_FIELD_KEYS.map(k => `<option>${k}</option>`).join("");
  colSel.innerHTML = `<option>(None)</option>` + PIVOT_FIELD_KEYS.map(k => `<option>${k}</option>`).join("");

  document.getElementById("buildPivotBtn").addEventListener("click", () => {
    const rows = JSON.parse(py.fetch_all("{}"));
    const result = JSON.parse(py.build_pivot(
      JSON.stringify(rows), rowSel.value, colSel.value,
      document.getElementById("pivotAgg").value
    ));
    lastPivot = result;
    renderPivotTable(result.headers, result.rows);
  });

  document.getElementById("exportPivotBtn").addEventListener("click", () => {
    if (!lastPivot) { alert("Build a table first."); return; }
    downloadCsv("indicator_summary_table.csv", lastPivot.headers, lastPivot.rows);
  });
}

function renderPivotTable(headers, rows) {
  document.querySelector("#pivotTable thead tr").innerHTML =
    headers.map(h => `<th>${escapeHtml(h)}</th>`).join("");
  document.querySelector("#pivotTable tbody").innerHTML =
    rows.map(r => `<tr>${r.map(c => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("");
}

/* ---- Cloud Sync tab (Google Drive folder + Sheets, via Google Identity Services) ---- */

const DRIVE = "https://www.googleapis.com/drive/v3";
const SHEETS = "https://sheets.googleapis.com/v4";
const GOOGLE_SCOPES = [
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/spreadsheets",
].join(" ");

const CLOUD_COLUMNS = ["document", "indicator", "value", "population_type", "gender", "unit",
  "source", "period_type", "period_value", "extracted_at"];
const CLOUD_HEADERS = ["Document", "Indicator", "Value", "Population Type", "Gender",
  "Household/Individual", "Source", "Reporting Period Type", "Reporting Period", "Extracted At"];

const GOOGLE_CONFIG_KEY = "indicator_extractor_google_config";
const DRIVE_SYNC_STATE_KEY = "indicator_extractor_drive_sync_state";

function loadGoogleConfig() {
  const raw = localStorage.getItem(GOOGLE_CONFIG_KEY);
  return raw ? JSON.parse(raw) : { clientId: "", folderLink: "", sheetName: "IndicatorData" };
}
function saveGoogleConfig(cfg) {
  localStorage.setItem(GOOGLE_CONFIG_KEY, JSON.stringify(cfg));
}

// Tracks which files (by id -> last-seen modifiedTime) have already been
// processed per folder, so re-running Sync only pulls new/changed files
// instead of reprocessing everything every time.
function loadDriveSyncState() {
  const raw = localStorage.getItem(DRIVE_SYNC_STATE_KEY);
  return raw ? JSON.parse(raw) : {};
}
function saveDriveSyncState(state) {
  localStorage.setItem(DRIVE_SYNC_STATE_KEY, JSON.stringify(state));
}

function extractFolderId(input) {
  const trimmed = (input || "").trim();
  const m = trimmed.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(trimmed)) return trimmed; // raw ID pasted directly
  return null;
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ---- Google sign-in (token-based, via Google Identity Services) ---- */

let googleTokenClient = null;
let googleTokenClientId = null;
let googleAccessToken = null;
let googleTokenExpiry = 0;

function getGoogleToken(clientId) {
  return new Promise((resolve, reject) => {
    if (googleAccessToken && Date.now() < googleTokenExpiry) {
      resolve(googleAccessToken);
      return;
    }
    if (!googleTokenClient || googleTokenClientId !== clientId) {
      googleTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: GOOGLE_SCOPES,
        callback: () => {}, // replaced per-call below
      });
      googleTokenClientId = clientId;
    }
    googleTokenClient.callback = (resp) => {
      if (resp.error) { reject(new Error(resp.error)); return; }
      googleAccessToken = resp.access_token;
      googleTokenExpiry = Date.now() + (resp.expires_in * 1000) - 60000;
      resolve(googleAccessToken);
    };
    googleTokenClient.error_callback = (err) => reject(new Error(err.message || "Sign-in was cancelled or failed."));
    googleTokenClient.requestAccessToken({ prompt: googleAccessToken ? "" : "consent" });
  });
}

/* ---- Drive / Sheets REST calls ---- */

async function driveFetch(token, url, options = {}) {
  return fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
}

async function listDocxInFolder(token, folderId) {
  const q = encodeURIComponent(
    `'${folderId}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document' and trashed=false`
  );
  const resp = await driveFetch(token, `${DRIVE}/files?q=${q}&fields=files(id,name,modifiedTime)&pageSize=1000`);
  if (!resp.ok) throw new Error(`Couldn't list the folder's contents (${resp.status}): ${await resp.text()}`);
  const data = await resp.json();
  return data.files || [];
}

async function downloadFileBytes(token, fileId) {
  const resp = await fetch(`${DRIVE}/files/${fileId}?alt=media`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error(`Couldn't download a file (${resp.status}): ${await resp.text()}`);
  return new Uint8Array(await resp.arrayBuffer());
}

async function findOrCreateSheet(token, folderId, sheetName) {
  const safeName = sheetName.replace(/'/g, "\\'");
  const q = encodeURIComponent(
    `'${folderId}' in parents and name='${safeName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`
  );
  let resp = await driveFetch(token, `${DRIVE}/files?q=${q}&fields=files(id,name)`);
  if (!resp.ok) throw new Error(`Couldn't look for the results Sheet (${resp.status}): ${await resp.text()}`);
  let data = await resp.json();
  if (data.files && data.files.length) return data.files[0].id;

  resp = await driveFetch(token, `${DRIVE}/files`, {
    method: "POST",
    body: JSON.stringify({
      name: sheetName,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [folderId],
    }),
  });
  if (!resp.ok) throw new Error(`Couldn't create the results Sheet (${resp.status}): ${await resp.text()}`);
  data = await resp.json();
  return data.id;
}

async function ensureSheetHeader(token, spreadsheetId, headers) {
  const range = "A1:" + colLetter(headers.length) + "1";
  let resp = await fetch(`${SHEETS}/spreadsheets/${spreadsheetId}/values/${range}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (resp.ok) {
    const data = await resp.json();
    if (data.values && data.values.length) return; // header row already present
  }
  resp = await fetch(`${SHEETS}/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=RAW`, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ values: [headers] }),
  });
  if (!resp.ok) throw new Error(`Couldn't write the header row (${resp.status}): ${await resp.text()}`);
}

async function appendSheetRows(token, spreadsheetId, rowsValues) {
  const CHUNK = 200;
  for (let i = 0; i < rowsValues.length; i += CHUNK) {
    const chunk = rowsValues.slice(i, i + CHUNK);
    const resp = await fetch(
      `${SHEETS}/spreadsheets/${spreadsheetId}/values/A1:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ values: chunk }),
      }
    );
    if (!resp.ok) throw new Error(`Couldn't append rows (${resp.status}): ${await resp.text()}`);
  }
}

/* ---- Combined sync: pull new/changed docs from the folder, extract,
   then push new rows to the results Sheet in that same folder ---- */

function wireCloudTab() {
  const cfg = loadGoogleConfig();
  document.getElementById("googleClientId").value = cfg.clientId;
  document.getElementById("googleFolderLink").value = cfg.folderLink;
  document.getElementById("googleSheetName").value = cfg.sheetName || "IndicatorData";

  const statusEl = document.getElementById("cloudStatus");

  function readForm() {
    return {
      clientId: document.getElementById("googleClientId").value.trim(),
      folderLink: document.getElementById("googleFolderLink").value.trim(),
      sheetName: document.getElementById("googleSheetName").value.trim() || "IndicatorData",
    };
  }

  document.getElementById("saveGoogleSettingsBtn").addEventListener("click", () => {
    saveGoogleConfig(readForm());
    statusEl.textContent = "Settings saved.";
  });

  document.getElementById("googleSignInBtn").addEventListener("click", async () => {
    try {
      const cfg2 = readForm();
      if (!cfg2.clientId) { statusEl.textContent = "Fill in the Client ID first."; return; }
      statusEl.textContent = "Opening sign-in…";
      await getGoogleToken(cfg2.clientId);
      statusEl.textContent = "Signed in.";
    } catch (err) {
      statusEl.textContent = "Sign-in error: " + err.message;
      console.error(err);
    }
  });

  document.getElementById("googleSyncBtn").addEventListener("click", async () => {
    try {
      const cfg2 = readForm();
      if (!cfg2.clientId || !cfg2.folderLink) {
        statusEl.textContent = "Fill in the Client ID and the Drive folder link, then Save Settings first.";
        return;
      }
      saveGoogleConfig(cfg2);

      const folderId = extractFolderId(cfg2.folderLink);
      if (!folderId) {
        statusEl.textContent = "Couldn't read a folder ID from that link — paste the full Drive folder URL.";
        return;
      }

      statusEl.textContent = "Signing in…";
      const token = await getGoogleToken(cfg2.clientId);

      statusEl.textContent = "Checking the folder for documents…";
      const files = await listDocxInFolder(token, folderId);

      const syncState = loadDriveSyncState();
      const stateForFolder = syncState[folderId] || {};
      const newOrChanged = files.filter(f => stateForFolder[f.id] !== f.modifiedTime);

      const indicators = loadIndicators();
      const pType = document.getElementById("periodType").value;
      const pValue = pType === "All" ? "" : document.getElementById("periodValue").value.trim();

      let allResults = [];
      for (const f of newOrChanged) {
        statusEl.textContent = `Downloading and processing ${f.name}…`;
        const bytes = await downloadFileBytes(token, f.id);
        const path = "/uploads/" + f.name;
        pyodide.FS.writeFile(path, bytes);
        const results = JSON.parse(py.extract(path, JSON.stringify(indicators)));
        results.forEach(r => { r.period_type = pType; r.period_value = pValue; });
        allResults = allResults.concat(results);
        stateForFolder[f.id] = f.modifiedTime;
      }

      if (allResults.length) {
        py.insert_rows(JSON.stringify(allResults));
        await persistDbToIndexedDb();
        renderResultsTable(allResults);
      }
      syncState[folderId] = stateForFolder;
      saveDriveSyncState(syncState);

      statusEl.textContent = "Checking the results Sheet…";
      const spreadsheetId = await findOrCreateSheet(token, folderId, cfg2.sheetName);
      await ensureSheetHeader(token, spreadsheetId, CLOUD_HEADERS);

      const unpushed = JSON.parse(py.fetch_unpushed());
      if (!unpushed.length) {
        statusEl.textContent = newOrChanged.length
          ? `Processed ${newOrChanged.length} file(s), found ${allResults.length} value(s). Sheet already up to date.`
          : "No new or changed files in the folder. Sheet already up to date.";
        return;
      }

      statusEl.textContent = `Pushing ${unpushed.length} row(s) to the Sheet…`;
      const values = unpushed.map(r => CLOUD_COLUMNS.map(c => r[c] ?? ""));
      await appendSheetRows(token, spreadsheetId, values);
      py.mark_pushed(JSON.stringify(unpushed.map(r => r.id)));
      await persistDbToIndexedDb();

      statusEl.textContent = `Synced: processed ${newOrChanged.length} file(s), pushed ${unpushed.length} row(s) to the Sheet.`;
    } catch (err) {
      statusEl.textContent = "Error: " + err.message;
      console.error(err);
    }
  });
}


/* ---- Backup export/import (footer) ---- */

function wireBackupButtons() {
  document.getElementById("exportDbBtn").addEventListener("click", () => {
    const bytes = pyodide.FS.readFile("/data/indicator_data.db");
    downloadBlob("indicator_data_backup.db", new Blob([bytes], { type: "application/octet-stream" }));
  });

  const importInput = document.getElementById("importDbInput");
  document.getElementById("importDbBtn").addEventListener("click", () => importInput.click());
  importInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!confirm("This replaces your current in-browser dataset with the backup file. Continue?")) return;
    const buf = await file.arrayBuffer();
    pyodide.FS.writeFile("/data/indicator_data.db", new Uint8Array(buf));
    await persistDbToIndexedDb();
    refreshDataTab();
    alert("Backup restored.");
    importInput.value = "";
  });
}

boot();
