import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import JSZip from "jszip";

// TODO: Update this to your deployed backend URL if different.
// const API_BASE_URL = "http://localhost:3000";
const API_BASE_URL = "https://api.cbtechmp.com";
const TOKEN_STORAGE_KEY = "penExcelAuthToken";
const USER_STORAGE_KEY = "penExcelAuthUser";

const state = {
  runSummary: null,
  targetTabId: null,
  token: null,
  user: null
};

const WEB_COLUMNS = [
  "SN",
  "PEN ID",
  "AWC Child ID",
  "Name",
  "Gender",
  "DOB",
  "Age",
  "Class",
  "Father's Name",
  "Mother's Name",
  "ABHA Number",
  "Mobile Number",
  "Screening Actions",
  "Weight (kg) *",
  "Height/Length (cm) *",
  "Blood Pressure",
  "Vision - Left Eye",
  "Vision - Right Eye",
  "Birth Defect Found",
  "Referal",
  "Select Defect type",
  "Defect Name",
  "Defect Other",
  "Identification Code"
];

/** Columns scraped from the table for "Download Detailed Excel". Only this list is filled; template keeps all other columns as-is. */
const DOWNLOAD_TABLE_COLUMNS = [
  "SN",
  "PEN ID",
  "AWC Child ID",
  "Name",
  "Gender",
  "DOB",
  "Age",
  "Class",
  "Father's Name",
  "Mother's Name",
  "Mobile Number"
];

/** Aliases for matching template column headers to our canonical names. */
const HEADER_ALIASES = {
  "SN": ["S.No.", "S No", "Serial No", "Serial Number", "Sl No", "Sl.No.", "Sr No"],
  "PEN ID": ["PEN ID / AWC Child ID", "PEN ID/AWC Child ID", "PEN ID or AWC Child ID", "PEN"],
  "AWC Child ID": [
    "AWC Child ID",
    "AWCChildID",
    "AWC ChildID",
    "AWC Child Id",
    "AWC ID",
    "AWC",
    "Child AWC ID"
  ],
  "Name": ["Student Name", "Child Name", "Student's Name"],
  "Gender": ["Sex"],
  "DOB": ["Date of Birth", "Date Of Birth", "D.O.B."],
  "Age": ["Age (Years)", "Age (years)"],
  "Class": ["Class/Grade", "Grade"],
  "Father's Name": ["Father Name", "Father's name"],
  "Mother's Name": ["Mother Name", "Mother's name"],
  "Mobile Number": ["Mobile", "Mobile No", "Phone", "Contact Number", "Phone Number"]
};

const signinSection = document.getElementById("signinSection");
const mainSection = document.getElementById("mainSection");
const signinStaffIdInput = document.getElementById("signinStaffId");
const signinEmailInput = document.getElementById("signinEmail");
const signinPasswordInput = document.getElementById("signinPassword");
const signinButton = document.getElementById("signinButton");
const signinStatusText = document.getElementById("signinStatusText");

const excelInput = document.getElementById("excelFile");
const runButton = document.getElementById("runButton");
const downloadDetailedButton = document.getElementById("downloadDetailedButton");
const statusText = document.getElementById("statusText");
const previewContainer = document.getElementById("previewContainer");
const closeButton = document.getElementById("closeButton");
const logoutButton = document.getElementById("logoutButton");
const entryCountBar = document.getElementById("entryCountBar");
const entryCountValueEl = document.getElementById("entryCountValue");

const queryParams = new URLSearchParams(window.location.search);
const queryTabId = Number(queryParams.get("tabId"));
if (Number.isInteger(queryTabId) && queryTabId > 0) {
  state.targetTabId = queryTabId;
}

function setSigninStatus(message, kind = "info") {
  if (!signinStatusText) return;
  signinStatusText.textContent = message;
  signinStatusText.className = `status-text ${kind}`;
}

function showSignin() {
  if (signinSection) signinSection.style.display = "block";
  if (mainSection) mainSection.style.display = "none";
  if (logoutButton) logoutButton.style.display = "none";
  if (entryCountBar) entryCountBar.style.display = "none";
}

function updateEntryCountDisplay() {
  if (entryCountValueEl) {
    const n = state.user?.entryCount;
    entryCountValueEl.textContent = typeof n === "number" ? String(n) : "—";
  }
}

function showMain() {
  if (signinSection) signinSection.style.display = "none";
  if (mainSection) mainSection.style.display = "block";
  if (logoutButton) logoutButton.style.display = "";
  if (entryCountBar) entryCountBar.style.display = "";
  updateEntryCountDisplay();
}

function loadStoredToken() {
  if (!chrome.storage || !chrome.storage.local) {
    // Storage permission not available – fall back to signin every time.
    showSignin();
    return;
  }

  chrome.storage.local.get([TOKEN_STORAGE_KEY, USER_STORAGE_KEY], (result) => {
    const token = result?.[TOKEN_STORAGE_KEY];
    if (typeof token === "string" && token) {
      state.token = token;
      const user = result?.[USER_STORAGE_KEY];
      if (user && typeof user === "object") state.user = user;
      showMain();
    } else {
      showSignin();
    }
  });
}

function loadStoredUser() {
  if (!chrome.storage || !chrome.storage.local) {
    return;
  }
  chrome.storage.local.get([USER_STORAGE_KEY], (result) => {
    const user = result?.[USER_STORAGE_KEY];
    if (user && typeof user === "object") {
      state.user = user;
    }
  });
}

/**
 * Fetches current user from API (GET auth/me), stores in state and chrome.storage.
 * Requires state.token to be set. Returns the user object or throws on failure.
 */
async function fetchUser() {
  if (!state.token) {
    throw new Error("Not signed in.");
  }
  const res = await fetch(`${API_BASE_URL}/auth/me`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${state.token}`
    }
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    state.user = null;
    if (chrome.storage?.local) {
      chrome.storage.local.remove(USER_STORAGE_KEY, () => {});
    }
    throw new Error(data?.message || "Failed to fetch user.");
  }
  const user = data?.user ?? data;
  if (!user || typeof user !== "object") {
    throw new Error("Invalid user response.");
  }
  state.user = user;
  if (chrome.storage?.local) {
    chrome.storage.local.set({ [USER_STORAGE_KEY]: user }, () => {});
  }
  return user;
}

async function signinUser() {
  if (!signinStaffIdInput || !signinPasswordInput) {
    setSigninStatus("Signin form not ready.", "error");
    return;
  }

  const staffID = signinStaffIdInput.value.trim();
  const password = signinPasswordInput.value.trim();

  if (!staffID || !password) {
    setSigninStatus("All fields are required to sign in.", "error");
    return;
  }

  try {
    setSigninStatus("Signing you in…", "info");
    signinButton.disabled = true;

    const res = await fetch(`${API_BASE_URL}/auth/login`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        staffID,
        password
      })
    });

    const data = await res.json().catch(() => ({}));

    if (!res.ok || !data?.success || !data?.token) {
      const msg = data?.message || "Sign in failed. Please check your details and try again.";
      setSigninStatus(msg, "error");
      signinButton.disabled = false;
      return;
    }

    state.token = data.token;
    if (chrome.storage?.local) {
      chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: state.token }, () => {});
    }

    await fetchUser();

    setSigninStatus("Sign in successful. You can now use PEN Excel Updater.", "ok");
    showMain();
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    setSigninStatus(msg || "Sign in failed. Please try again.", "error");
    signinButton.disabled = false;
  }
}

function setStatus(message, kind = "info") {
  statusText.textContent = message;
  statusText.className = `status-text ${kind}`;
}

function normalizePenValue(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/\s+/g, "")
    .trim();
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

/** For matching template column headers to scraped headers: normalize for comparison. */
function normalizeHeaderForMatch(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/\*$/g, "")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\s/g, "");
}

function renderSummary(summary) {
  if (!summary) {
    previewContainer.textContent = "No run summary yet.";
    return;
  }

  const rows = [
    ["Total Excel Rows", summary.total],
    ["Attempted (with PEN ID / AWC Child ID)", summary.attempted],
    ["Success", summary.success],
    ["Failed", summary.failed]
  ];

  const table = document.createElement("table");
  table.className = "preview-table";

  const thead = document.createElement("thead");
  thead.innerHTML = "<tr><th>Field</th><th>Value</th></tr>";
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const [key, val] of rows) {
    const tr = document.createElement("tr");
    const tdKey = document.createElement("td");
    const tdVal = document.createElement("td");
    tdKey.textContent = key;
    tdVal.textContent = String(val ?? "");
    tr.append(tdKey, tdVal);
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  previewContainer.innerHTML = "";
  previewContainer.appendChild(table);

  if (Array.isArray(summary.failures) && summary.failures.length) {
    const details = document.createElement("div");
    details.className = "failures";

    const title = document.createElement("h3");
      title.textContent = "Failed PEN IDs / AWC Child IDs";
    details.appendChild(title);

    const list = document.createElement("ul");
    for (const item of summary.failures.slice(0, 20)) {
      const li = document.createElement("li");
      li.textContent = `${item.penId}: ${item.reason}`;
      list.appendChild(li);
    }
    details.appendChild(list);

    if (summary.failures.length > 20) {
      const more = document.createElement("p");
      more.textContent = `...and ${summary.failures.length - 20} more failures.`;
      details.appendChild(more);
    }
    previewContainer.appendChild(details);
  }
}

function isExtensionPage(url) {
  return String(url || "").startsWith("chrome-extension://");
}

async function getTargetTab() {
  if (typeof state.targetTabId === "number") {
    try {
      const tab = await chrome.tabs.get(state.targetTabId);
      if (tab?.id && !isExtensionPage(tab.url)) return tab;
    } catch (_error) {
      state.targetTabId = null;
    }
  }

  const activeTabs = await chrome.tabs.query({ active: true });
  const bestMatch = activeTabs.find((tab) => typeof tab.id === "number" && !isExtensionPage(tab.url));
  if (bestMatch) {
    state.targetTabId = bestMatch.id;
    return bestMatch;
  }

  throw new Error("No target webpage tab found. Open a website tab, then click the extension icon from that tab.");
}

function isSupportedPage(url) {
  if (!url) return false;
  return /^(https?:\/\/|file:\/\/)/i.test(url);
}

function isMissingReceiverError(error) {
  const msg = String(error?.message || error || "").toLowerCase();
  return msg.includes("receiving end does not exist") || msg.includes("could not establish connection");
}

async function sendMessageToTab(tabId, payload) {
  return chrome.tabs.sendMessage(tabId, payload);
}

async function injectContentScript(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["build/content.js"]
  });
}

function beautifyPageError(error, url) {
  const raw = String(error?.message || error || "");
  const lower = raw.toLowerCase();

  if (lower.includes("cannot access contents of url")) {
    if ((url || "").startsWith("file://")) {
      return "This is a local file page. Enable extension access to file URLs in Chrome extension details, then retry.";
    }
    return "This page blocks extension access. Try on a normal website tab (http/https).";
  }
  if (isMissingReceiverError(error)) {
    return "Could not connect to page script. Refresh the page and try again.";
  }
  return raw || "Could not process the page.";
}

async function sendAutoFillRequest(tabId, rows, options = {}) {
  return sendMessageToTab(tabId, {
    type: "AUTO_FILL_EXCEL_ROWS",
    rows,
    options
  });
}

async function sendExtractTableRequest(tabId) {
  return sendMessageToTab(tabId, {
    type: "EXTRACT_STUDENT_TABLE",
    options: {
      allPages: true
    }
  });
}

async function runAutoFillOnPage(rows) {
  // Safety: require a token before running main actions.
  if (!state.token) {
    showSignin();
    throw new Error("Please sign in first to use PEN Excel Updater.");
  }

  const tab = await getTargetTab();
  const tabId = tab.id;
  const url = tab.url || "";

  if (!isSupportedPage(url)) {
    throw new Error("Open a normal website tab (http/https or file://), then try again.");
  }

  try {
    const response = await sendAutoFillRequest(tabId, rows, { stopAfterFirst: false, autoSubmit: true });
    if (!response?.ok) throw new Error(response?.error || "Could not auto-fill the page.");
    return response.summary;
  } catch (firstError) {
    if (!isMissingReceiverError(firstError)) {
      throw new Error(beautifyPageError(firstError, url));
    }

    try {
      await injectContentScript(tabId);
      const secondResponse = await sendAutoFillRequest(tabId, rows, { stopAfterFirst: false, autoSubmit: true });
      if (!secondResponse?.ok) throw new Error(secondResponse?.error || "Could not auto-fill the page.");
      return secondResponse.summary;
    } catch (secondError) {
      throw new Error(beautifyPageError(secondError, url));
    }
  }
}

function timestampForFileName() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `${y}${m}${d}_${hh}${mm}${ss}`;
}

function downloadRowsAsExcel(headers, rows) {
  if (!Array.isArray(headers) || !Array.isArray(rows)) {
    throw new Error("Download failed: headers and rows must be arrays.");
  }
  const worksheetRows = rows.map((row) => {
    const out = {};
    for (const header of headers) out[header] = row[header] ?? "";
    return out;
  });
  const worksheet = XLSX.utils.json_to_sheet(worksheetRows, { header: headers });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Students");
  const filename = `student_table_${timestampForFileName()}.xlsx`;

  if (typeof chrome !== "undefined" && chrome.downloads && typeof chrome.downloads.download === "function") {
    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({ url, filename, saveAs: true }, () => {
      setTimeout(() => URL.revokeObjectURL(url), 10000);
    });
  } else {
    XLSX.writeFile(workbook, filename);
  }
}

/**
 * Filter table export to only the columns we want for Download Detailed Excel.
 * Maps table headers to canonical names so template columns line up.
 * @param {string[]} headers Table headers from the page.
 * @param {Object[]} rows Table rows (objects keyed by header).
 * @returns {{ headers: string[], rows: Object[] }}
 */
function filterTableDataForDownload(headers, rows) {
  const canonicalToTableHeader = new Map();
  DOWNLOAD_TABLE_COLUMNS.forEach((canonical) => {
    const norm = normalizeHeaderForMatch(canonical);
    const aliases = HEADER_ALIASES[canonical] || [];
    const found = headers.find((h) => {
      const normH = normalizeHeaderForMatch(h);
      return normH === norm || aliases.some((a) => normalizeHeaderForMatch(a) === normH);
    });
    if (found != null) canonicalToTableHeader.set(canonical, found);
  });

  // Extra-robust: detect AWC column even if header text is slightly different.
  // (Some pages use "AWC Child ID", others "AWC ID", etc.)
  if (!canonicalToTableHeader.has("AWC Child ID")) {
    const awcHeader = headers.find((h) => {
      const nh = normalizeHeader(String(h || ""));
      return (
        nh === "awcchildid" ||
        nh === "awcid" ||
        nh === "awc" ||
        (nh.includes("awc") && nh.includes("id"))
      );
    });
    if (awcHeader != null) canonicalToTableHeader.set("AWC Child ID", awcHeader);
  }

  // PEN ID and AWC Child ID represent the same identifier for our workflows.
  // If the table doesn't have a PEN ID column but does have an AWC Child ID column,
  // treat AWC as PEN so the template's PEN column gets filled.
  if (!canonicalToTableHeader.has("PEN ID") && canonicalToTableHeader.has("AWC Child ID")) {
    canonicalToTableHeader.set("PEN ID", canonicalToTableHeader.get("AWC Child ID"));
  }

  const getRowValueByHeader = (row, headerText) => {
    if (!row || !headerText) return "";
    const v = row[headerText];
    return v !== undefined && v !== null ? v : "";
  };

  const outHeaders = DOWNLOAD_TABLE_COLUMNS.filter((c) => canonicalToTableHeader.has(c));

  // Critical: ensure PEN ID is always included if we can derive it from AWC.
  // Otherwise the template writer never attempts to write the PEN ID cell.
  if (!outHeaders.includes("PEN ID") && canonicalToTableHeader.has("AWC Child ID")) {
    outHeaders.splice(1, 0, "PEN ID"); // right after SN, to keep a sensible order
  }

  const outRows = rows.map((row) => {
    const out = {};

    // Fill all requested columns normally...
    outHeaders.forEach((h) => {
      const tableKey = canonicalToTableHeader.get(h);
      out[h] = tableKey != null ? getRowValueByHeader(row, tableKey) : "";
    });

    // ...but ALWAYS treat AWC Child ID as PEN ID (same meaning).
    // Even if the table contains a PEN column, it may be blank while AWC has the real ID.
    const penHeader = canonicalToTableHeader.get("PEN ID") || "";
    const awcHeader = canonicalToTableHeader.get("AWC Child ID") || "";
    const rawPen = getRowValueByHeader(row, penHeader);
    const rawAwc = getRowValueByHeader(row, awcHeader);
    const penText = String(rawPen ?? "").trim();
    const awcText = String(rawAwc ?? "").trim();
    if (!penText && awcText) out["PEN ID"] = rawAwc;
    // Also keep PEN if present but ensure it is copied consistently for downstream.
    if (!out["PEN ID"] && penText) out["PEN ID"] = rawPen;

    return out;
  });
  return { headers: outHeaders, rows: outRows };
}

/**
 * Load the bundled Template.xlsx from the extension.
 * @returns {Promise<ArrayBuffer|null>} Template buffer or null if not available.
 */
async function getTemplateBuffer() {
  if (typeof chrome !== "undefined" && chrome.runtime?.getURL) {
    const url = chrome.runtime.getURL("templates/Template.xlsx");
    const res = await fetch(url);
    if (res.ok) return res.arrayBuffer();
  }
  return null;
}

/** Get a string from an ExcelJS cell (handles value, richText, number). */
function getCellText(cell) {
  if (!cell) return "";
  const v = cell.value;
  if (v == null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (v instanceof Date) return v.toISOString ? v.toISOString() : String(v);
  if (typeof v === "object" && v.richText && Array.isArray(v.richText)) {
    return v.richText.map((t) => (typeof t === "string" ? t : t.text || "")).join("").trim();
  }
  if (typeof v === "object" && typeof v.text === "string") return v.text.trim();
  return String(v).trim();
}

/** Find header row (1-based). Scans row 1..5 and picks row with most matches to our columns. */
function findTemplateHeaderRow(sheet) {
  const MAX_HEADER_COLS = 120;
  const SCAN_ROWS = 5;
  let best = { headerRow: 1, score: 0 };

  for (let rowNum = 1; rowNum <= SCAN_ROWS; rowNum++) {
    let score = 0;
    for (const canonical of DOWNLOAD_TABLE_COLUMNS) {
      const normC = normalizeHeaderForMatch(canonical);
      const aliases = HEADER_ALIASES[canonical] || [];
      let match = false;
      for (let col = 1; col <= MAX_HEADER_COLS; col++) {
        const th = getCellText(sheet.getCell(rowNum, col));
        if (!th) continue;
        const normTh = normalizeHeaderForMatch(th);
        if (normTh === normC || aliases.some((a) => normalizeHeaderForMatch(a) === normTh)) {
          match = true;
          break;
        }
      }
      if (match) score++;
    }
    if (score > best.score) best = { headerRow: rowNum, score };
  }

  return best;
}

function findTemplateColumnIndex(sheet, headerRow, canonical) {
  const MAX_HEADER_COLS = 120;
  const normC = normalizeHeaderForMatch(canonical);
  const aliases = HEADER_ALIASES[canonical] || [];
  for (let col = 1; col <= MAX_HEADER_COLS; col++) {
    const th = getCellText(sheet.getCell(headerRow, col));
    if (!th) continue;
    const normTh = normalizeHeaderForMatch(th);
    if (normTh === normC || aliases.some((a) => normalizeHeaderForMatch(a) === normTh)) return col;
  }
  return null;
}

function findTemplateColumnsByHeader(sheet, headerRow, headerText) {
  const MAX_HEADER_COLS = 120;
  const wanted = normalizeHeaderForMatch(headerText);
  const cols = [];
  for (let col = 1; col <= MAX_HEADER_COLS; col++) {
    const th = getCellText(sheet.getCell(headerRow, col));
    if (!th) continue;
    if (normalizeHeaderForMatch(th) === wanted) cols.push(col);
  }
  return cols;
}

function cloneDataValidation(dv) {
  if (!dv) return null;
  return {
    type: dv.type,
    allowBlank: dv.allowBlank,
    operator: dv.operator,
    showErrorMessage: dv.showErrorMessage,
    showInputMessage: dv.showInputMessage,
    promptTitle: dv.promptTitle,
    prompt: dv.prompt,
    errorTitle: dv.errorTitle,
    error: dv.error,
    errorStyle: dv.errorStyle,
    formulae: Array.isArray(dv.formulae) ? dv.formulae.slice() : dv.formulae
  };
}

function adjustValidationFormulaeForRow(formulae, rowNumber) {
  if (!Array.isArray(formulae)) return formulae;
  return formulae.map((f) => {
    const s = String(f ?? "");
    // Common pattern in your template: INDIRECT(AM3) -> INDIRECT(AM{row})
    return s.replace(/INDIRECT\(([A-Z]{1,3})(\d+)\)/gi, (_m, col, _oldRow) => `INDIRECT(${col}${rowNumber})`);
  });
}

function colLetter(index1Based) {
  let letter = "";
  let n = index1Based - 1;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

function escapeXml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getFirstWorksheetXmlPathFromZip(zip) {
  const workbookXml = zip.file("xl/workbook.xml");
  const relsXml = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relsXml) return "xl/worksheets/sheet1.xml";
  // We'll parse in the caller (async) and fall back to sheet1.xml on any issue.
  return null;
}

async function resolveFirstWorksheetXmlPath(zip) {
  const direct = getFirstWorksheetXmlPathFromZip(zip);
  if (direct) return direct;

  try {
    const [workbookStr, relsStr] = await Promise.all([
      zip.file("xl/workbook.xml").async("string"),
      zip.file("xl/_rels/workbook.xml.rels").async("string")
    ]);
    const parser = new DOMParser();
    const wbDoc = parser.parseFromString(workbookStr, "application/xml");
    const relDoc = parser.parseFromString(relsStr, "application/xml");
    const sheetEl =
      wbDoc.querySelector("workbook > sheets > sheet") ||
      wbDoc.getElementsByTagName("sheet")?.[0] ||
      null;
    const rid = sheetEl?.getAttribute("r:id") || sheetEl?.getAttribute("id");
    if (!rid) return "xl/worksheets/sheet1.xml";

    const relEl = Array.from(relDoc.getElementsByTagName("Relationship")).find((el) => el.getAttribute("Id") === rid);
    const target = relEl?.getAttribute("Target") || "";
    if (!target) return "xl/worksheets/sheet1.xml";

    const normalized = target.replace(/^\.?\//, "");
    return normalized.startsWith("xl/") ? normalized : `xl/${normalized}`;
  } catch (_e) {
    return "xl/worksheets/sheet1.xml";
  }
}

function setCellInlineString(sheetDoc, rowNumber, colNumber1Based, value) {
  const ns = sheetDoc.documentElement?.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const createEl = (name) => sheetDoc.createElementNS(ns, name);
  const col = colLetter(colNumber1Based);
  const cellRef = `${col}${rowNumber}`;

  // Ensure row exists
  const sheetData = sheetDoc.getElementsByTagNameNS(ns, "sheetData")?.[0] || sheetDoc.getElementsByTagName("sheetData")?.[0];
  if (!sheetData) return;

  const existingRows = Array.from(sheetData.getElementsByTagNameNS(ns, "row"));
  let rowEl = existingRows.find((r) => Number(r.getAttribute("r") || 0) === rowNumber) || null;
  if (!rowEl) {
    rowEl = createEl("row");
    rowEl.setAttribute("r", String(rowNumber));
    // Insert in order if possible
    const next = existingRows.find((r) => Number(r.getAttribute("r") || 0) > rowNumber);
    if (next) sheetData.insertBefore(rowEl, next);
    else sheetData.appendChild(rowEl);
  }

  // Find or create cell
  const existingCells = Array.from(rowEl.getElementsByTagNameNS(ns, "c"));
  let cEl = existingCells.find((c) => c.getAttribute("r") === cellRef) || null;
  if (!cEl) {
    cEl = createEl("c");
    cEl.setAttribute("r", cellRef);
    // Insert in column order
    const next = existingCells.find((c) => {
      const r = c.getAttribute("r") || "";
      const m = r.match(/^([A-Z]+)(\d+)$/);
      if (!m) return false;
      // Compare by column number using a rough conversion
      const letters = m[1];
      let num = 0;
      for (let i = 0; i < letters.length; i++) num = num * 26 + (letters.charCodeAt(i) - 64);
      return num > colNumber1Based;
    });
    if (next) rowEl.insertBefore(cEl, next);
    else rowEl.appendChild(cEl);
  }

  // Set as inline string
  cEl.setAttribute("t", "inlineStr");
  // Remove existing children
  while (cEl.firstChild) cEl.removeChild(cEl.firstChild);
  const isEl = createEl("is");
  const tEl = createEl("t");
  // Preserve spaces if present
  tEl.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
  tEl.textContent = String(value ?? "");
  isEl.appendChild(tEl);
  cEl.appendChild(isEl);
}

function shiftFormulaForRow(formula, rowDelta) {
  const s = String(formula ?? "");
  if (!rowDelta) return s;
  // Shift only relative row references. Absolute rows ($12) remain fixed.
  return s.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (_m, colAbs, col, rowAbs, rowNumText) => {
    if (rowAbs === "$") return `${colAbs}${col}${rowAbs}${rowNumText}`;
    const rowNum = Number(rowNumText);
    if (!Number.isFinite(rowNum)) return `${colAbs}${col}${rowAbs}${rowNumText}`;
    return `${colAbs}${col}${rowAbs}${rowNum + rowDelta}`;
  });
}

function ensureCellNode(sheetDoc, rowNumber, colLetters) {
  const ns = sheetDoc.documentElement?.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const createEl = (name) => sheetDoc.createElementNS(ns, name);
  const ref = `${colLetters}${rowNumber}`;
  const sheetData = sheetDoc.getElementsByTagNameNS(ns, "sheetData")?.[0] || sheetDoc.getElementsByTagName("sheetData")?.[0];
  if (!sheetData) return null;

  const allRows = Array.from(sheetData.getElementsByTagNameNS(ns, "row"));
  let rowEl = allRows.find((r) => Number(r.getAttribute("r") || 0) === rowNumber) || null;
  if (!rowEl) {
    rowEl = createEl("row");
    rowEl.setAttribute("r", String(rowNumber));
    const nextRow = allRows.find((r) => Number(r.getAttribute("r") || 0) > rowNumber);
    if (nextRow) sheetData.insertBefore(rowEl, nextRow);
    else sheetData.appendChild(rowEl);
  }

  const existingCells = Array.from(rowEl.getElementsByTagNameNS(ns, "c"));
  let cEl = existingCells.find((c) => c.getAttribute("r") === ref) || null;
  if (!cEl) {
    cEl = createEl("c");
    cEl.setAttribute("r", ref);
    rowEl.appendChild(cEl);
  }
  return cEl;
}

function ensureFormulaRows(sheetDoc, sourceRow, targetEndRow) {
  if (!sourceRow || targetEndRow <= sourceRow) return;
  const ns = sheetDoc.documentElement?.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const rowNodes = Array.from(sheetDoc.getElementsByTagNameNS(ns, "row"));
  const sourceRowNode = rowNodes.find((r) => Number(r.getAttribute("r") || 0) === sourceRow);
  if (!sourceRowNode) return;

  const formulaCells = Array.from(sourceRowNode.getElementsByTagNameNS(ns, "c"))
    .map((cell) => {
      const ref = cell.getAttribute("r") || "";
      const m = ref.match(/^([A-Z]{1,3})(\d+)$/);
      if (!m) return null;
      const colLetters = m[1];
      const fNode = cell.getElementsByTagNameNS(ns, "f")?.[0] || null;
      if (!fNode || !fNode.textContent) return null;
      const styleId = cell.getAttribute("s");
      return { colLetters, formula: fNode.textContent, styleId };
    })
    .filter(Boolean);

  if (!formulaCells.length) return;

  for (let rowNum = sourceRow + 1; rowNum <= targetEndRow; rowNum++) {
    const rowDelta = rowNum - sourceRow;
    for (const fc of formulaCells) {
      const existing = ensureCellNode(sheetDoc, rowNum, fc.colLetters);
      if (!existing) continue;
      const fExisting = existing.getElementsByTagNameNS(ns, "f")?.[0] || existing.getElementsByTagName("f")?.[0] || null;
      if (fExisting && String(fExisting.textContent || "").trim() !== "") continue;

      if (fc.styleId && !existing.getAttribute("s")) existing.setAttribute("s", fc.styleId);
      while (existing.firstChild) existing.removeChild(existing.firstChild);
      const fNode = sheetDoc.createElementNS(ns, "f");
      fNode.textContent = shiftFormulaForRow(fc.formula, rowDelta);
      existing.appendChild(fNode);
    }
  }
}

function expandDataValidationsToRows(sheetDoc, sourceRow, targetEndRow) {
  if (!sourceRow || targetEndRow <= sourceRow) return;
  const ns = sheetDoc.documentElement?.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const dvNodes = Array.from(sheetDoc.getElementsByTagNameNS(ns, "dataValidation"));
  dvNodes.forEach((dv) => {
    const sqref = dv.getAttribute("sqref");
    if (!sqref) return;
    const refs = sqref.split(/\s+/).filter(Boolean);
    const out = [];
    refs.forEach((token) => {
      const range = token.match(/^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/);
      if (range) {
        const r1 = Number(range[2]);
        const r2 = Number(range[4]);
        if (r1 <= sourceRow && sourceRow <= r2) {
          out.push(`${range[1]}${sourceRow}:${range[3]}${targetEndRow}`);
        } else {
          out.push(token);
        }
        return;
      }
      const single = token.match(/^([A-Z]{1,3})(\d+)$/);
      if (single) {
        const r = Number(single[2]);
        if (r === sourceRow) out.push(`${single[1]}${sourceRow}:${single[1]}${targetEndRow}`);
        else out.push(token);
        return;
      }
      out.push(token);
    });
    dv.setAttribute("sqref", Array.from(new Set(out)).join(" "));
  });
}

/**
 * Fill template workbook with scraped data, preserving template dropdowns/validations.
 * Maps scraped headers to template columns by normalized header match + aliases; columns not in template are skipped.
 */
async function downloadWithTemplate(templateBuffer, headers, rows) {
  // IMPORTANT: ExcelJS rewriting can break complex named ranges / validations in some templates.
  // We only use ExcelJS here to locate header row and column indices, then we patch the XLSX XML directly
  // to preserve all original dropdowns, defined names, and validation data.

  const probeWb = new ExcelJS.Workbook();
  await probeWb.xlsx.load(templateBuffer);
  if (!probeWb.worksheets || probeWb.worksheets.length === 0) {
    throw new Error("Template has no worksheets or could not be read.");
  }
  const probeSheet = probeWb.worksheets[0];
  const { headerRow } = findTemplateHeaderRow(probeSheet);

  const headerToColIndex = new Map();
  headers.forEach((canonical) => {
    const col = findTemplateColumnIndex(probeSheet, headerRow, canonical);
    if (col != null) headerToColIndex.set(canonical, col);
  });
  if (headerToColIndex.size === 0) {
    throw new Error(
      "Template has no columns matching SN, PEN ID, Name, etc. Ensure the template header row contains at least one of these column names."
    );
  }

  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetPath = await resolveFirstWorksheetXmlPath(zip);
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`Template worksheet XML not found at ${sheetPath}.`);

  const xmlStr = await sheetFile.async("string");
  const parser = new DOMParser();
  const sheetDoc = parser.parseFromString(xmlStr, "application/xml");
  const firstDataRow = headerRow + 1;
  // Keep formulas + dropdowns resilient against row deletions by pre-extending them.
  const protectedEndRow = Math.max(1000, firstDataRow + rows.length + 200);
  ensureFormulaRows(sheetDoc, firstDataRow, protectedEndRow);
  expandDataValidationsToRows(sheetDoc, firstDataRow, protectedEndRow);

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const excelRow = headerRow + 1 + r;
    headers.forEach((header) => {
      const colIndex = headerToColIndex.get(header);
      if (colIndex == null) return;
      const val = row[header];
      const hasValue =
        val !== undefined &&
        val !== null &&
        (typeof val === "number" || String(val).trim() !== "");
      if (!hasValue) return;
      // Write as inline string to avoid touching sharedStrings.xml
      setCellInlineString(sheetDoc, excelRow, colIndex, typeof val === "number" ? String(val) : String(val).trim());
    });
  }

  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(sheetDoc);
  zip.file(sheetPath, newXml);
  const out = await zip.generateAsync({ type: "arraybuffer" });
  return out;
}

function downloadBlobAsFile(blob, filename) {
  if (typeof chrome !== "undefined" && chrome.downloads?.download) {
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({ url, filename, saveAs: true }, () => {
      setTimeout(() => URL.revokeObjectURL(url), 10000);
    });
  } else {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }
}

function renderTableExportSummary(data) {
  const rows = [
    ["Detected Table Index", data.tableIndex],
    ["Columns", data.headers.length],
    ["Rows Exported", data.rows.length]
  ];

  const table = document.createElement("table");
  table.className = "preview-table";
  const thead = document.createElement("thead");
  thead.innerHTML = "<tr><th>Metric</th><th>Value</th></tr>";
  table.appendChild(thead);
  const tbody = document.createElement("tbody");

  for (const [key, val] of rows) {
    const tr = document.createElement("tr");
    const tdKey = document.createElement("td");
    const tdVal = document.createElement("td");
    tdKey.textContent = key;
    tdVal.textContent = String(val ?? "");
    tr.append(tdKey, tdVal);
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  const headersInfo = document.createElement("p");
  headersInfo.className = "hint";
  headersInfo.textContent = `Headers: ${data.headers.join(", ")}`;

  previewContainer.innerHTML = "";
  previewContainer.appendChild(table);
  previewContainer.appendChild(headersInfo);
}

function renderDetailedExportSummary(data) {
  const rows = [
    ["Rows Exported", data.rows.length],
    ["Columns", data.headers.length],
    ["Failed Rows", Array.isArray(data.failures) ? data.failures.length : 0]
  ];

  const table = document.createElement("table");
  table.className = "preview-table";
  const thead = document.createElement("thead");
  thead.innerHTML = "<tr><th>Metric</th><th>Value</th></tr>";
  table.appendChild(thead);
  const tbody = document.createElement("tbody");

  for (const [key, val] of rows) {
    const tr = document.createElement("tr");
    const tdKey = document.createElement("td");
    const tdVal = document.createElement("td");
    tdKey.textContent = key;
    tdVal.textContent = String(val ?? "");
    tr.append(tdKey, tdVal);
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  previewContainer.innerHTML = "";
  previewContainer.appendChild(table);

  if (Array.isArray(data.failures) && data.failures.length) {
    const details = document.createElement("div");
    details.className = "failures";

    const title = document.createElement("h3");
      title.textContent = "Failed PEN IDs / AWC Child IDs";
    details.appendChild(title);

    const list = document.createElement("ul");
    for (const item of data.failures.slice(0, 20)) {
      const li = document.createElement("li");
      li.textContent = `${item.penId}: ${item.reason}`;
      list.appendChild(li);
    }
    details.appendChild(list);
    previewContainer.appendChild(details);
  }
}

async function runTableExportFromPage() {
  if (!state.token) {
    showSignin();
    throw new Error("Please sign in first to use PEN Excel Updater.");
  }
  const tab = await getTargetTab();
  const tabId = tab.id;
  const url = tab.url || "";

  if (!isSupportedPage(url)) {
    throw new Error("Open a normal website tab (http/https or file://), then try again.");
  }

  try {
    const response = await sendExtractTableRequest(tabId);
    if (!response?.ok) throw new Error(response?.error || "Could not extract student table.");
    return response.data;
  } catch (firstError) {
    if (!isMissingReceiverError(firstError)) {
      throw new Error(beautifyPageError(firstError, url));
    }

    await injectContentScript(tabId);
    const secondResponse = await sendExtractTableRequest(tabId);
    if (!secondResponse?.ok) throw new Error(secondResponse?.error || "Could not extract student table.");
    return secondResponse.data;
  }
}

function findPenColumnIndex(headers) {
  const candidates = new Set([
    "penid",
    "pen",
    "penidentifier",
    "awcchildid",
    "awcchild",
    "awcid",
    "awc"
  ]);
  for (let i = 0; i < headers.length; i += 1) {
    const normalized = normalizeHeader(headers[i]);
    if (candidates.has(normalized)) return i;
  }
  return -1;
}

function findHeaderRowIndex(aoa) {
  const scanLimit = Math.min(40, aoa.length);
  const idColumnNames = new Set(["penid", "pen", "awcchildid", "awcchild", "awcid", "awc"]);
  for (let r = 0; r < scanLimit; r += 1) {
    const row = aoa[r] || [];
    const hasPenOrAwcColumn = row.some((cell) => idColumnNames.has(normalizeHeader(cell)));
    if (hasPenOrAwcColumn) return r;
  }
  return -1;
}

function toCanonicalColumnName(rawName) {
  const normalizedRaw = normalizeHeader(rawName);
  const matched = WEB_COLUMNS.find((col) => normalizeHeader(col) === normalizedRaw);
  return matched || String(rawName || "").replace(/\u00A0/g, " ").trim();
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.readAsArrayBuffer(file);
  });
}

function getDirectCellValue(sheet, rowIdx, colIdx) {
  const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const cell = sheet[addr];
  if (!cell) return "";
  if (cell.w != null && String(cell.w).trim() !== "") return String(cell.w).trim();
  if (cell.v != null) return String(cell.v).trim();
  if (cell.h != null) return String(cell.h).replace(/<[^>]*>/g, "").trim();
  return "";
}

async function parseExcelRows(file) {
  const fileBuffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(fileBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Workbook has no sheets.");

  const sheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!aoa.length) throw new Error("Sheet is empty.");

  const headerRowIndex = findHeaderRowIndex(aoa);
  if (headerRowIndex < 0) {
    throw new Error("Could not detect a header row containing PEN ID or AWC Child ID.");
  }

  const headers = [...(aoa[headerRowIndex] || [])].map((h) => toCanonicalColumnName(h));
  const penColIndex = findPenColumnIndex(headers);
  if (penColIndex < 0) {
    throw new Error("Could not find PEN ID / AWC Child ID column. Use: PEN ID, AWC Child ID, PEN_ID, pen_id, or PEN.");
  }

  const rows = [];

  for (let r = headerRowIndex + 1; r < aoa.length; r += 1) {
    const row = aoa[r] || [];
    const penId = normalizePenValue(row[penColIndex]);
    if (!penId) continue;

    const record = {};
    for (let c = 0; c < headers.length; c += 1) {
      const colName = headers[c];
      if (!colName) continue;
      let val = String(row[c] ?? "").trim();
      if (!val) {
        val = getDirectCellValue(sheet, r, c);
      }
      if (!val && record[colName]) continue;
      record[colName] = val;
    }

    // Ensure downstream always has a PEN ID field even if the sheet uses AWC Child ID.
    if (
      (!record["PEN ID"] || String(record["PEN ID"]).trim() === "") &&
      record["AWC Child ID"] &&
      String(record["AWC Child ID"]).trim() !== ""
    ) {
      record["PEN ID"] = record["AWC Child ID"];
    }
    rows.push(record);
  }

  if (!rows.length) throw new Error("No valid data rows with PEN ID / AWC Child ID were found in Excel.");
  return rows;
}

runButton.addEventListener("click", async () => {
  try {
    await fetchUser();
    if (state.user?.permission === 0) {
      setStatus("Your permission is not allowed yet", "error");
      return;
    }
    const file = excelInput.files?.[0];
    if (!file) {
      setStatus("Please select an Excel/CSV file first.", "error");
      return;
    }

    runButton.disabled = true;
    setStatus("Reading Excel rows...");

    const rows = await parseExcelRows(file);
    const limit = state.user?.entryCount;
    if (typeof limit === "number" && rows.length > limit) {
      setStatus(
        `The number of rows in this file exceeds the allowed limit. You can import up to ${limit} rows, but the file currently contains ${rows.length} rows. Please reduce the number of rows and try again.`,
        "error"
      );
      return;
    }
    setStatus(`Starting one-person auto-fill + submit for ${rows.length} Excel rows...`);

    const summary = await runAutoFillOnPage(rows);
    state.runSummary = summary;
    renderSummary(summary);

    if (summary.success > 0 && summary.failed === 0) {
      setStatus(`Completed: Screening Details filled for PEN ID / AWC Child ID. Stopped on details page (no submit).`, "ok");
    } else {
      setStatus(
        `Completed with issues: ${summary.success} success, ${summary.failed} failed. Check failed list below.`,
        "error"
      );
    }

    if (summary.success > 0) {
      await fetch(`${API_BASE_URL}/subscription/entry-count/update`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${state.token}`
        },
        body: JSON.stringify({
          entryCount: summary.success
        })
      });
      await fetchUser();
      updateEntryCountDisplay();
    }
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error), "error");
  } finally {
    runButton.disabled = false;
  }
});


if (downloadDetailedButton) {
  downloadDetailedButton.addEventListener("click", async () => {
    if (state.user?.permission === 0) {
      setStatus("Your permission is not allowed yet", "error");
      return;
    }
    try {
      downloadDetailedButton.disabled = true;
      setStatus("Reading table from page (SN, PEN ID, Name, Gender, DOB, Age, Class, Father's Name, Mother's Name, Mobile Number)...");
      const data = await runTableExportFromPage();
      const rawHeaders = data && data.headers;
      const rawRows = data && data.rows;
      if (!Array.isArray(rawHeaders) || !Array.isArray(rawRows)) {
        throw new Error("Export did not return valid data (missing headers or rows).");
      }
      if (rawRows.length === 0) {
        setStatus("No rows to export. Check the page and try again.", "error");
        return;
      }

      const { headers, rows } = filterTableDataForDownload(rawHeaders, rawRows);
      if (!headers.length) {
        setStatus("Could not find any of the required columns (SN, PEN ID, Name, etc.) in the table.", "error");
        return;
      }

      const templateBuffer = await getTemplateBuffer();
      const filename = `student_table_${timestampForFileName()}.xlsx`;

      if (!templateBuffer) {
        throw new Error(
          'Template.xlsx not found in the extension. Place your real template at "Excel project/templates/Template.xlsx", then rebuild and reload the extension.'
        );
      }

      setStatus("Filling template with table data (all pages, other columns preserved)...");
      const buffer = await downloadWithTemplate(templateBuffer, headers, rows);
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      downloadBlobAsFile(blob, filename);
      setStatus(`Export complete: ${rows.length} row(s) filled into template. Download started.`, "ok");

      renderDetailedExportSummary({ headers, rows, failures: [] });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), "error");
    } finally {
      downloadDetailedButton.disabled = false;
    }
  });
}

function clearAllStorage() {
  state.token = null;
  state.user = null;
  if (chrome.storage?.local) {
    chrome.storage.local.clear(() => {});
  }
}

if (closeButton) {
  closeButton.addEventListener("click", () => {
    clearAllStorage();
    window.close();
  });
}

// When user closes popup via OS X (or any other close), clear all storage
window.addEventListener("beforeunload", () => {
  clearAllStorage();
});

// Wire up signin button and load any stored token when popup opens.
if (signinButton) {
  signinButton.addEventListener("click", () => {
    void signinUser();
  });
}

loadStoredToken();
loadStoredUser();
