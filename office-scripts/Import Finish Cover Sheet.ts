/**
 * Import Finish Cover Sheet
 *
 * Reads the office-prepared MASTER FINISH COVER SHEET (table FINISHCOVER, job date in B4)
 * and returns the pre-fill rows for Power Automate to write into the SCHEDULED JOBS list.
 * Comment columns are intentionally ignored; installers enter those in the app.
 *
 * Fails with a clear message (so the flow can email the sender) when:
 *   - the FINISHCOVER table is missing
 *   - B4 does not hold a real job date
 *   - no job rows are filled in
 *
 * Written without arrow functions, escaped quotes or regular expressions so it survives
 * being pasted into the Excel script editor.
 */

interface ScheduledRow {
  jobNumber: string;
  supervisor: string;
  leadInstaller: string;
  helperInstaller: string;
  builder: string;
  subdivision: string;
  lot: string;
  phase: string;
  address: string;
  importKey: string;
}

interface ImportResult {
  jobDate: string; // yyyy-mm-dd
  rows: ScheduledRow[];
  warnings: string[];
}

function main(workbook: ExcelScript.Workbook): ImportResult {
  const table = workbook.getTable("FINISHCOVER");
  if (!table) {
    throw new Error("Table FINISHCOVER was not found. Use the MASTER FINISH COVER SHEET template.");
  }

  const jobDate = readJobDate(table.getWorksheet().getRange("B4"));

  // Look columns up by header name so column order changes in the template don't break the import.
  const headerTexts = table.getHeaderRowRange().getTexts()[0];
  const headers: string[] = [];
  for (let h = 0; h < headerTexts.length; h++) {
    headers.push(headerTexts[h].trim().toLowerCase());
  }

  const cJob = requiredColumn(headers, "Job Number");
  const cSupervisor = requiredColumn(headers, "Supervisor");
  const cLead = requiredColumn(headers, "Lead Installer");
  const cHelper = requiredColumn(headers, "Helper Installer");
  const cBuilder = requiredColumn(headers, "Builder");
  const cSubdivision = requiredColumn(headers, "Subdivision");
  const cLot = requiredColumn(headers, "Lot #");
  const cPhase = requiredColumn(headers, "Phase");
  // Optional column: older cover sheets don't have it.
  const cAddress = headers.indexOf("address");

  const body = table.getRangeBetweenHeaderAndTotal();
  if (!body) {
    throw new Error("Table FINISHCOVER has no rows.");
  }
  // getTexts() returns what is displayed, so lot numbers like 110 stay text and cached lookup values are used.
  const texts = body.getTexts();

  const rows: ScheduledRow[] = [];
  const warnings: string[] = [];
  const seen: string[] = [];

  for (let i = 0; i < texts.length; i++) {
    const r = texts[i];
    const rawJob = clean(r[cJob]).toUpperCase();
    if (rawJob === "") {
      continue; // unused template row
    }
    // Same rule as the app: add N only when it is missing.
    const jobNumber = rawJob.startsWith("N") ? rawJob : "N" + rawJob;
    const sheetRow = i + 7; // table header is row 6

    if (seen.indexOf(jobNumber) >= 0) {
      warnings.push("Row " + sheetRow + ": " + jobNumber + " is listed more than once; only the first row was imported.");
      continue;
    }
    seen.push(jobNumber);

    let address = "";
    if (cAddress >= 0) {
      address = lookupValue(r[cAddress]);
    }

    const row: ScheduledRow = {
      jobNumber: jobNumber,
      supervisor: lookupValue(r[cSupervisor]),
      leadInstaller: clean(r[cLead]),
      helperInstaller: clean(r[cHelper]),
      builder: lookupValue(r[cBuilder]),
      subdivision: lookupValue(r[cSubdivision]),
      lot: lookupValue(r[cLot]),
      phase: clean(r[cPhase]).toUpperCase(),
      address: address,
      importKey: jobNumber + "|" + jobDate
    };

    const missing: string[] = [];
    if (row.supervisor === "") missing.push("Supervisor");
    if (row.builder === "") missing.push("Builder");
    if (row.subdivision === "") missing.push("Subdivision");
    if (row.lot === "") missing.push("Lot #");
    if (row.leadInstaller === "") missing.push("Lead Installer");
    if (missing.length > 0) {
      warnings.push("Row " + sheetRow + " (" + jobNumber + "): missing " + missing.join(", ") + ". Installer will need to fill these in.");
    }

    rows.push(row);
  }

  if (rows.length === 0) {
    throw new Error("No job numbers were found in table FINISHCOVER.");
  }

  return { jobDate: jobDate, rows: rows, warnings: warnings };
}

/** Index of a required column; stops with a clear message if it is missing. */
function requiredColumn(headers: string[], name: string): number {
  const i = headers.indexOf(name.toLowerCase());
  if (i < 0) {
    throw new Error("Column '" + name + "' was not found in table FINISHCOVER.");
  }
  return i;
}

/** Collapses spaces, tabs and line breaks from wrapped cells into single spaces. */
function clean(value: string): string {
  if (value === undefined || value === null) {
    return "";
  }
  const text = String(value);
  let out = "";
  let lastWasSpace = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    const isSpace = text.charCodeAt(i) <= 32; // space, tab, line break
    if (isSpace) {
      if (!lastWasSpace) {
        out = out + " ";
      }
      lastWasSpace = true;
    } else {
      out = out + ch;
      lastWasSpace = false;
    }
  }
  return out.trim();
}

/** Production-schedule XLOOKUP cells show "Not found" or "#REF!"-style errors when the link fails; treat those as blank. */
function lookupValue(value: string): string {
  const v = clean(value);
  if (v === "" || v.toLowerCase() === "not found" || v.startsWith("#")) {
    return "";
  }
  return v;
}

/** True when every character of the text is a digit 0-9 (and it isn't empty). */
function isDigits(text: string): boolean {
  if (text.length === 0) {
    return false;
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (ch < "0" || ch > "9") {
      return false;
    }
  }
  return true;
}

/** Accepts a real Excel date or m/d/yyyy text in B4 and returns yyyy-mm-dd. */
function readJobDate(cell: ExcelScript.Range): string {
  const raw = cell.getValue();
  const badDate = "Job Date in B4 (" + cell.getText() + ") is not a valid date. Use m/d/yyyy.";
  let y: number;
  let m: number;
  let d: number;

  if (typeof raw === "number") {
    // Excel serial date -> calendar date (UTC avoids time-zone shifts).
    const dt = new Date(Math.round((raw - 25569) * 86400 * 1000));
    y = dt.getUTCFullYear();
    m = dt.getUTCMonth() + 1;
    d = dt.getUTCDate();
  } else {
    const parts = String(raw).trim().split("/");
    if (parts.length !== 3 || !isDigits(parts[0]) || !isDigits(parts[1]) || !isDigits(parts[2]) || parts[2].length > 4) {
      throw new Error(badDate);
    }
    m = Number(parts[0]);
    d = Number(parts[1]);
    y = Number(parts[2]);
    if (y < 100) {
      y += 2000;
    }
  }

  const check = new Date(Date.UTC(y, m - 1, d));
  const currentYear = new Date().getFullYear();
  if (
    check.getUTCFullYear() !== y || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d ||
    y < currentYear - 1 || y > currentYear + 1
  ) {
    throw new Error(badDate);
  }

  const mm = m < 10 ? "0" + m : "" + m;
  const dd = d < 10 ? "0" + d : "" + d;
  return y + "-" + mm + "-" + dd;
}
