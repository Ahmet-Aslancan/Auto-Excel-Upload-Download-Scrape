/**
 * Generates templates/Template.xlsx with header row and dropdown data validations.
 * Run: node scripts/generate-template.mjs
 * The generated file is used by the extension for "Download Detailed Excel" (template-based).
 */
import ExcelJS from "exceljs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");
const outPath = path.join(root, "templates", "Template.xlsx");

// Headers that match the export (main table + details). Order preserved for template.
const TEMPLATE_HEADERS = [
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
  "Weight (kg)*",
  "Height/Length (cm)*",
  "BMI (calculated)",
  "BMI Classification",
  "BMI Result",
  "Blood Pressure*",
  "Blood Pressure Classification",
  "Vision - Left Eye*",
  "Vision - Right Eye *",
  "Hb Count",
  "Birth Defect Found",
  "Referal",
  "Select Defect type",
  "Defect Name",
  "Defect Other",
  "Identification Code",
  "2nd defect Available Yes/No",
  "Select 2nd Defect if Available",
  "2nd Defect Name"
];

// Columns that get a dropdown (list validation). Key = exact header; value = options array.
const DROPDOWN_OPTIONS = {
  "Birth Defect Found": ["Yes", "No"],
  Referal: ["DEIC", "Non-DEIC", "Other"],
  "Select Defect type": [
    "Defects at Birth",
    "Defects_at_birth",
    "Defects After Birth",
    "Defects_after_birth",
    "Other"
  ],
  "2nd defect Available Yes/No": ["Yes", "No"],
  "Select 2nd Defect if Available": [
    "Defects at Birth",
    "Defects_at_birth",
    "Defects After Birth",
    "Defects_after_birth",
    "Other"
  ]
};

function colLetter(index) {
  let letter = "";
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Students", { views: [{ state: "frozen", ySplit: 1 }] });

  // Header row (row 1)
  TEMPLATE_HEADERS.forEach((header, colIndex) => {
    const cell = sheet.getCell(1, colIndex + 1);
    cell.value = header;
    cell.font = { bold: true };
  });

  // Add data validations for dropdown columns (apply to many rows so user can select)
  // Increase this to cover typical exports across many pages.
  const dataRowCount = 10000;
  for (let c = 0; c < TEMPLATE_HEADERS.length; c++) {
    const header = TEMPLATE_HEADERS[c];
    const options = DROPDOWN_OPTIONS[header];
    if (!options || !options.length) continue;
    const range = `${colLetter(c)}2:${colLetter(c)}${dataRowCount + 1}`;
    const formula = '"' + options.map((o) => String(o).replace(/"/g, '""')).join(",") + '"';
    sheet.dataValidations.add(range, {
      type: "list",
      allowBlank: true,
      formulae: [formula]
    });
  }

  // Set column widths for readability
  sheet.columns.forEach((col, i) => {
    col.width = Math.min( Math.max(TEMPLATE_HEADERS[i]?.length || 10, 12), 40 );
  });

  const fs = await import("node:fs");
  const dir = path.dirname(outPath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  await workbook.xlsx.writeFile(outPath);
  console.log("Written:", outPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
