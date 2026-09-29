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
  const headers = (table.getHeaderRowRange().getTexts()[0]).map(h => h.trim().toLowerCase());
  const col = (name: string): number => {
    const i = headers.indexOf(name.toLowerCase());
    if (i < 0) {
      throw new Error("Column \"" + name + "\" was not found in table FINISHCOVER.");
    }
    return i;
  };
  const c = {
    job: col("Job Number"),
    supervisor: col("Supervisor"),
    lead: col("Lead Installer"),
    helper: col("Helper Installer"),
    builder: col("Builder"),
    subdivision: col("Subdivision"),
    lot: col("Lot #"),
    phase: col("Phase")
  };

  const body = table.getRangeBetweenHeaderAndTotal();
  if (!body) {
    throw new Error("Table FINISHCOVER has no rows.");
  }
  // getTexts() returns what is displayed, so lot numbers like 110 stay text and cached lookup values are used.
  const texts = body.getTexts();

  const rows: ScheduledRow[] = [];
  const warnings: string[] = [];
  const seen: string[] = [];

  texts.forEach((r, i) => {
    const rawJob = clean(r[c.job]).toUpperCase();
    if (rawJob === "") {
      return; // unused template row
    }
    // Same rule as the app: add N only when it is missing.
    const jobNumber = rawJob.startsWith("N") ? rawJob : "N" + rawJob;
    const sheetRow = i + 7; // table header is row 6

    if (seen.indexOf(jobNumber) >= 0) {
      warnings.push("Row " + sheetRow + ": " + jobNumber + " is listed more than once; only the first row was imported.");
      return;
    }
    seen.push(jobNumber);

    const row: ScheduledRow = {
      jobNumber: jobNumber,
      supervisor: lookupValue(r[c.supervisor]),
      leadInstaller: clean(r[c.lead]),
      helperInstaller: clean(r[c.helper]),
      builder: lookupValue(r[c.builder]),
      subdivision: lookupValue(r[c.subdivision]),
      lot: lookupValue(r[c.lot]),
      phase: clean(r[c.phase]).toUpperCase(),
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
  });

  if (rows.length === 0) {
    throw new Error("No job numbers were found in table FINISHCOVER.");
  }

  return { jobDate: jobDate, rows: rows, warnings: warnings };
}

/** Collapses whitespace/line breaks from wrapped cells. */
function clean(value: string): string {
  return (value === undefined || value === null ? "" : String(value)).replace(/\s+/g, " ").trim();
}

/** Production-schedule XLOOKUP cells show "Not found" or "#REF!"-style errors when the link fails; treat those as blank. */
function lookupValue(value: string): string {
  const v = clean(value);
  if (v === "" || v.toLowerCase() === "not found" || v.startsWith("#")) {
    return "";
  }
  return v;
}

/** Accepts a real Excel date or m/d/yyyy text in B4 and returns yyyy-mm-dd. */
function readJobDate(cell: ExcelScript.Range): string {
  const raw = cell.getValue();
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
    const digitsOnly = parts.every(p => p.length > 0 && !isNaN(Number(p)) && p.indexOf(".") < 0);
    if (parts.length !== 3 || !digitsOnly || parts[2].length > 4) {
      throw new Error("Job Date in B4 (\"" + cell.getText() + "\") is not a valid date. Use m/d/yyyy.");
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
    throw new Error("Job Date in B4 (\"" + cell.getText() + "\") is not a valid date. Use m/d/yyyy.");
  }

  return y + "-" + (m < 10 ? "0" : "") + m + "-" + (d < 10 ? "0" : "") + d;
}
