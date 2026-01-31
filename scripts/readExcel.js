const xlsx = require("xlsx");
const path = require("path");

const HEADER_ROW_INDEX = 4; // Excel row 5 (0-based)

function cleanKey(k) {
  return String(k || "").trim().replace(/\s+/g, " ");
}

function loadTestCases() {
  const filePath = path.join(__dirname, "..", "data", "IT23240506.xlsx");

  const wb = xlsx.readFile(filePath);

  
  const sheetName = wb.SheetNames.find(
    (name) => name.trim() === "Test cases"
  );

  if (!sheetName) {
    throw new Error(
      `Sheet "Test cases" not found. Available sheets: ${wb.SheetNames.join(", ")}`
    );
  }

  const sheet = wb.Sheets[sheetName];

  let rows = xlsx.utils.sheet_to_json(sheet, {
    range: HEADER_ROW_INDEX,
    defval: "",
  });

  // Normalize header keys
  rows = rows.map((row) => {
    const normalized = {};
    for (const key of Object.keys(row)) {
      normalized[cleanKey(key)] = row[key];
    }
    return normalized;
  });

  // Return only valid test cases
  return rows
    .filter((r) => cleanKey(r["TC ID"]) !== "")
    .map((r) => ({
      tcId: cleanKey(r["TC ID"]),
      name: cleanKey(r["Test case name"]),
      type: cleanKey(r["Input length type"]),
      input: String(r["Input"] || "").trim(),
      expected: String(r["Expected output"] || "").trim(),
      status: String(r["Status"] || "").trim(),
    }));
}

module.exports = { loadTestCases };
