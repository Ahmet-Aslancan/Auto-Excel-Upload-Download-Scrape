import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const ROOT = process.cwd();
const INPUT_FILE = path.join(ROOT, "Project sheet School.xlsx");
const OUTPUT_FILE = path.join(ROOT, "student-portal.html");

const COLUMNS = [
  "SN",
  "PEN ID",
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
  "Defect Other",
  "Identification Code"
];

function normalize(value) {
  return String(value || "")
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isRowEmpty(row) {
  return !(row || []).some((cell) => String(cell || "").trim() !== "");
}

function detectHeaderRow(rows) {
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const hasPenId = row.some((cell) => normalize(cell) === "penid" || normalize(cell) === "pen");
    const hasSn = row.some((cell) => normalize(cell) === "sn");
    if (hasPenId && hasSn) return i;
  }
  return -1;
}

function findInstituteName(rows, headerRowIndex) {
  for (let i = headerRowIndex - 1; i >= 0; i -= 1) {
    const row = rows[i] || [];
    const values = row.map((c) => String(c || "").trim()).filter(Boolean);
    if (!values.length) continue;

    const first = values[0];
    if (/institute\s*name|institute\s*nale/i.test(first) && values[1]) return values[1];
    return first;
  }
  return "Student Table";
}

function extractStudents(rows) {
  const headerRowIndex = detectHeaderRow(rows);
  if (headerRowIndex < 0) throw new Error("Could not detect the table header row (SN + PEN ID).");

  const sourceHeaders = (rows[headerRowIndex] || []).map((cell) =>
    String(cell || "").replace(/\u00A0/g, " ").trim()
  );
  const sourceIndexes = COLUMNS.map((target) =>
    sourceHeaders.findIndex((source) => normalize(source) === normalize(target))
  );
  const penCol = COLUMNS.findIndex((c) => normalize(c) === "penid");

  const students = [];
  for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
    const row = rows[r] || [];
    if (isRowEmpty(row)) continue;

    const item = {};
    for (let i = 0; i < COLUMNS.length; i += 1) {
      const src = sourceIndexes[i];
      item[COLUMNS[i]] = src >= 0 ? String(row[src] ?? "").trim() : "";
    }

    if (!String(item[COLUMNS[penCol]] || "").trim()) continue;
    students.push(item);
  }

  return {
    instituteName: findInstituteName(rows, headerRowIndex),
    students
  };
}

function renderHtml({ instituteName, students }) {
  const seedJson = JSON.stringify(students).replaceAll("<", "\\u003c");
  const columnsJson = JSON.stringify(COLUMNS).replaceAll("<", "\\u003c");

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Student Auto-Fill Portal</title>
    <style>
      :root {
        --bg: #f3f6ff;
        --card: #ffffff;
        --text: #151a32;
        --muted: #667089;
        --brand: #5e37e8;
        --brand-2: #864fff;
        --accent: #0faec2;
        --ok: #0b8f63;
        --danger: #d93036;
        --border: #dde3f3;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: Inter, Segoe UI, Arial, sans-serif;
        color: var(--text);
        background: linear-gradient(160deg, #edf2ff 0%, #f8f4ff 100%);
      }
      .container {
        max-width: 1420px;
        margin: 0 auto;
        padding: 20px;
      }
      .hero {
        border-radius: 16px;
        background: linear-gradient(120deg, var(--brand), var(--brand-2) 72%, var(--accent));
        color: #fff;
        padding: 18px;
        margin-bottom: 14px;
      }
      .hero h1 { margin: 0 0 8px; font-size: 24px; }
      .hero p { margin: 0; opacity: 0.95; }
      .badge {
        display: inline-block;
        margin-top: 10px;
        border: 1px solid rgba(255,255,255,0.35);
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 12px;
        background: rgba(255,255,255,0.15);
      }
      .toolbar {
        display: flex;
        gap: 10px;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 10px 12px;
        margin-bottom: 10px;
      }
      .toolbar-left {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      .toolbar input {
        min-width: 300px;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 8px 10px;
      }
      .btn {
        border: none;
        border-radius: 9px;
        padding: 8px 12px;
        font-size: 13px;
        font-weight: 600;
        cursor: pointer;
      }
      .btn-outline {
        color: #4d5a7b;
        border: 1px solid var(--border);
        background: #fff;
      }
      .table-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 10px;
      }
      .table-wrap {
        border: 1px solid var(--border);
        border-radius: 10px;
        overflow: auto;
        max-height: calc(100vh - 260px);
      }
      table {
        width: 100%;
        min-width: 2100px;
        border-collapse: collapse;
      }
      th, td {
        border-bottom: 1px solid var(--border);
        border-right: 1px solid var(--border);
        padding: 8px 10px;
        white-space: nowrap;
        font-size: 12px;
        text-align: left;
      }
      th {
        position: sticky;
        top: 0;
        z-index: 1;
        font-weight: 700;
        background: #f5f8ff;
      }
      tbody tr:hover {
        background: #f8f5ff;
      }
      tbody tr {
        cursor: pointer;
      }
      th:first-child, td:first-child {
        position: sticky;
        left: 0;
        z-index: 2;
        background: #fff;
      }
      thead th:first-child {
        z-index: 3;
        background: #f5f8ff;
      }
      .empty {
        color: var(--muted);
        text-align: center;
      }
      .hint {
        font-size: 12px;
        color: var(--muted);
      }
      .hidden {
        display: none !important;
      }
      .editor-wrap {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 14px;
      }
      .editor-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
        margin-bottom: 12px;
      }
      .btn-primary {
        color: #fff;
        background: linear-gradient(120deg, var(--brand), var(--brand-2));
      }
      .btn-back {
        background: #fff;
        border: 1px solid var(--border);
      }
      .grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 10px;
      }
      .field {
        display: flex;
        flex-direction: column;
        gap: 6px;
      }
      .field label {
        font-size: 12px;
        font-weight: 600;
        color: #3f4b6b;
      }
      .field input {
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 8px 10px;
      }
      .readonly {
        background: #f7f8fc;
      }
      .status {
        margin-top: 10px;
        min-height: 18px;
        font-size: 12px;
      }
      .status.ok { color: var(--ok); }
      .status.error { color: var(--danger); }
    </style>
  </head>
  <body>
    <main class="container">
      <section class="hero">
        <h1>Student Data Editor</h1>
        <p>Double-click any student row to open one edit window. Save changes and return to full table view.</p>
        <span class="badge">${escapeHtml(instituteName)} | Seed rows: ${students.length}</span>
      </section>

      <section id="tableView">
        <div class="toolbar">
          <div class="toolbar-left">
            <input id="searchInput" placeholder="Search by PEN ID / Name / Class" />
            <button id="resetSeedBtn" class="btn btn-outline" type="button">Reset DB to Excel Seed</button>
          </div>
          <span class="hint">Tip: Double-click a row to edit selected student.</span>
        </div>
        <div class="table-card">
          <div class="table-wrap">
            <table id="studentTable">
              <thead><tr id="headerRow"></tr></thead>
              <tbody id="tableBody"></tbody>
            </table>
          </div>
        </div>
      </section>

      <section id="editorView" class="hidden">
        <div class="editor-wrap">
          <div class="editor-top">
            <div>
              <h2 style="margin:0">Edit Student</h2>
              <div class="hint">Update selected student data and submit to local database.</div>
            </div>
            <div style="display:flex; gap:8px">
              <button id="backBtn" type="button" class="btn btn-back">Back to Table</button>
              <button id="submitBtn" type="button" class="btn btn-primary">Submit Changes</button>
            </div>
          </div>
          <form id="editForm" class="grid"></form>
          <div id="editorStatus" class="status"></div>
        </div>
      </section>
    </main>

    <script>
      const DB_KEY = "studentPortalDbV1";
      const COLUMNS = ${columnsJson};
      const PEN_ID_KEY = "PEN ID";
      const seedStudents = ${seedJson};

      let students = [];
      let filtered = [];
      let selectedPenId = "";

      const tableView = document.getElementById("tableView");
      const editorView = document.getElementById("editorView");
      const tableBody = document.getElementById("tableBody");
      const headerRow = document.getElementById("headerRow");
      const searchInput = document.getElementById("searchInput");
      const resetSeedBtn = document.getElementById("resetSeedBtn");
      const editForm = document.getElementById("editForm");
      const editorStatus = document.getElementById("editorStatus");
      const submitBtn = document.getElementById("submitBtn");
      const backBtn = document.getElementById("backBtn");

      function readDb() {
        try {
          const raw = localStorage.getItem(DB_KEY);
          if (!raw) return null;
          const parsed = JSON.parse(raw);
          return Array.isArray(parsed) ? parsed : null;
        } catch (_err) {
          return null;
        }
      }

      function writeDb(rows) {
        localStorage.setItem(DB_KEY, JSON.stringify(rows));
      }

      function resetDbToSeed() {
        students = structuredClone(seedStudents);
        writeDb(students);
      }

      function loadStudents() {
        const dbRows = readDb();
        if (dbRows && dbRows.length) {
          students = dbRows;
        } else {
          resetDbToSeed();
        }
        filtered = [...students];
      }

      function renderHeader() {
        headerRow.innerHTML = COLUMNS.map((col) => "<th>" + col + "</th>").join("");
      }

      function renderTable(rows) {
        if (!rows.length) {
          tableBody.innerHTML = "<tr><td class='empty' colspan='" + COLUMNS.length + "'>No matching records</td></tr>";
          return;
        }

        const html = rows
          .map((row) => {
            const penId = row[PEN_ID_KEY] || "";
            const tds = COLUMNS.map((col) => "<td>" + (row[col] || "") + "</td>").join("");
            return "<tr data-pen-id='" + penId + "'>" + tds + "</tr>";
          })
          .join("");
        tableBody.innerHTML = html;
      }

      function applySearch(query) {
        const q = String(query || "").toLowerCase().trim();
        if (!q) {
          filtered = [...students];
        } else {
          filtered = students.filter((row) => {
            return (
              String(row["PEN ID"] || "").toLowerCase().includes(q) ||
              String(row["Name"] || "").toLowerCase().includes(q) ||
              String(row["Class"] || "").toLowerCase().includes(q)
            );
          });
        }
        renderTable(filtered);
      }

      function openEditorByPenId(penId) {
        const student = students.find((row) => String(row[PEN_ID_KEY]) === String(penId));
        if (!student) return;

        selectedPenId = String(penId);
        editorStatus.textContent = "";
        editorStatus.className = "status";

        editForm.innerHTML = COLUMNS.map((col) => {
          const readOnly = col === "SN" || col === "PEN ID";
          const value = String(student[col] || "").replaceAll('"', "&quot;");
          return (
            "<div class='field'>" +
            "<label>" + col + "</label>" +
            "<input data-col='" + col + "' value='" + value + "' " + (readOnly ? "readonly class='readonly'" : "") + " />" +
            "</div>"
          );
        }).join("");

        tableView.classList.add("hidden");
        editorView.classList.remove("hidden");
      }

      function backToTable() {
        editorView.classList.add("hidden");
        tableView.classList.remove("hidden");
        applySearch(searchInput.value);
      }

      async function saveStudentToDb(updatedStudent) {
        await new Promise((resolve) => setTimeout(resolve, 120));
        const idx = students.findIndex((row) => String(row[PEN_ID_KEY]) === String(updatedStudent[PEN_ID_KEY]));
        if (idx < 0) throw new Error("Student not found in database.");

        students[idx] = updatedStudent;
        writeDb(students);
      }

      function collectFormData() {
        const obj = {};
        const inputs = editForm.querySelectorAll("input[data-col]");
        inputs.forEach((input) => {
          const col = input.getAttribute("data-col");
          obj[col] = input.value.trim();
        });
        return obj;
      }

      searchInput.addEventListener("input", () => applySearch(searchInput.value));

      tableBody.addEventListener("dblclick", (event) => {
        const tr = event.target.closest("tr[data-pen-id]");
        if (!tr) return;
        openEditorByPenId(tr.getAttribute("data-pen-id"));
      });

      resetSeedBtn.addEventListener("click", () => {
        resetDbToSeed();
        applySearch(searchInput.value);
      });

      backBtn.addEventListener("click", backToTable);

      submitBtn.addEventListener("click", async () => {
        try {
          const updated = collectFormData();
          if (!updated[PEN_ID_KEY]) {
            throw new Error("PEN ID is required.");
          }
          submitBtn.disabled = true;
          await saveStudentToDb(updated);
          editorStatus.textContent = "Saved successfully to database.";
          editorStatus.className = "status ok";
          setTimeout(backToTable, 280);
        } catch (err) {
          editorStatus.textContent = err instanceof Error ? err.message : String(err);
          editorStatus.className = "status error";
        } finally {
          submitBtn.disabled = false;
        }
      });

      loadStudents();
      renderHeader();
      applySearch("");
    </script>
  </body>
</html>`;
}

if (!fs.existsSync(INPUT_FILE)) {
  throw new Error("Missing input file: " + INPUT_FILE);
}

const workbook = XLSX.readFile(INPUT_FILE);
const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });
const parsed = extractStudents(rows);
const html = renderHtml(parsed);
fs.writeFileSync(OUTPUT_FILE, html, "utf8");

console.log("Generated " + OUTPUT_FILE + " with " + parsed.students.length + " student rows.");
