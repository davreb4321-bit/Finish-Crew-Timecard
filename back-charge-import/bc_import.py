#!/usr/bin/env python3
"""
Back Charge Import
==================

Replaces IMPORTBC2.PY. Reads every accounting back charge PDF in the loading
folder and writes BC_Import_Ready.xlsx with one row per back charge, already in
the Back Charge Data column order (A-W).

What it fills in:
  * From the accounting page: PO#, posting date, job #, amount, INV #, check #,
    deposit (deduction) date, notes.
  * From the builder's page that follows: lot, activity, builder description and
    the builder's back charge date (used as the incident date).
  * From PRODUCTION DATA (read once, no XLOOKUP formulas): builder, address,
    subdivision, lot and supervisor for the job #.
  * Suggestions for the drop-down columns (Claim Type, Building Phase,
    Responsible Dept, Incident Type, Type of Property Damaged, Responsible
    Person) based on how similar notes were categorized in the Registry.
    Every suggested cell is highlighted yellow so it can be confirmed.

Back charges already in the Registry are left out and listed on the Review tab.

Usage
-----
    pip install pandas pyxlsb openpyxl pdfplumber
    python bc_import.py
    python bc_import.py --folder "D:/some/other/folder" --workbook "D:/tracking.xlsb"

Then, in the tracking workbook, click "Import Back Charges" (macro in
BackChargeImport.bas), confirm the yellow cells, and continue as before.
"""

import argparse
import math
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pyxlsb import open_workbook

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
LOADING_FOLDER = r"P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER"
TRACKING_WORKBOOK = os.path.join(LOADING_FOLDER, "2026 BACK CHARGE TRACKING.xlsb")
OUTPUT_FILE_NAME = "BC_Import_Ready.xlsx"

DEFAULT_RESPONDENT = "David Martin"   # column V
SUGGEST_NEIGHBOURS = 7                # how many similar past notes vote on each suggestion

# Back Charge Data columns A-W, in order (row 9 headings)
COLUMNS = [
    "Builder Name", "Phone Number", "Job #", "Job Address", "Subdivision", "Lot Number",
    "BC Amount", "Incident Date", "Deduction Date", "Posting Date", "Supervisor",
    "Claim Type", "Building Phase", "PO/Back Charge #", "Responsible Dept", "INV #",
    "Description", "Check #", "Incident Type", "Type of Property Damaged",
    "Responsible Person", "Respondent", "Customer Name",
]
DATE_COLUMNS = {"Incident Date", "Deduction Date", "Posting Date"}

# Drop-down column -> (Registry heading it is learned from, Drop-Down Data heading)
SUGGESTED = {
    "Claim Type": ("Claim Type", "Claim Type"),
    "Building Phase": ("Phase of Incident", "Building Phase"),
    "Responsible Dept": ("Responsible Department", "Responsible Department"),
    "Incident Type": ("Incident Type", "Incident Type"),
    "Type of Property Damaged": ("Type of Property Damaged", "Type of Property Damaged"),
}

GUESS_FILL = PatternFill("solid", fgColor="FFF2A8")     # yellow: please confirm
MISSING_FILL = PatternFill("solid", fgColor="F8CBAD")   # orange: could not be found


# --------------------------------------------------------------------------
# Reading the tracking workbook (read-only)
# --------------------------------------------------------------------------
def read_sheet(path, sheet, header_row):
    """header_row is the Excel row number holding the headings."""
    df = pd.read_excel(path, sheet_name=sheet, engine="pyxlsb", header=header_row - 1)
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    return df


def load_reference(path):
    print(f"Reading tracking workbook: {path}")
    prod = read_sheet(path, "PRODUCTION DATA", 7)
    prod["Job#"] = prod["Job#"].astype(str).str.strip().str.upper()
    prod = prod.drop_duplicates("Job#", keep="last").set_index("Job#")
    print(f"  PRODUCTION DATA: {len(prod):,} jobs")

    reg = read_sheet(path, "Registry", 9)
    reg = reg[reg["Description of Incident"].notna()].copy()
    print(f"  Registry: {len(reg):,} past back charges to learn from")

    # Drop-Down Data: list headings are on row 2, values below
    dd = pd.read_excel(path, sheet_name="Drop-Down Data", engine="pyxlsb", header=None)
    lists, builder_ids = {}, {}
    for c in dd.columns:
        head = str(dd.iat[1, c]).strip() if pd.notna(dd.iat[1, c]) else ""
        vals = [str(v).strip() for v in dd[c].iloc[2:] if pd.notna(v) and str(v).strip()]
        if head:
            lists[head] = vals
        if head == "Builder ID Number":
            ids = dd[c + 1].iloc[2:]
            for name, cid in zip(dd[c].iloc[2:], ids):
                if pd.notna(name) and pd.notna(cid):
                    builder_ids[str(cid).strip().upper()] = str(name).strip()
    return prod, reg, lists, builder_ids


# --------------------------------------------------------------------------
# PDF text
# --------------------------------------------------------------------------
def pdf_text(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text(layout=True) or "" for page in pdf.pages)
    if text.strip():
        return text
    return ocr_text(path)


def ocr_text(path):
    """Fallback for scanned PDFs. Needs pytesseract + Tesseract installed."""
    try:
        import pytesseract
    except ImportError:
        print("    No text layer and pytesseract is not installed; skipping.")
        return ""
    for exe in (r"C:\Users\DMartin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(exe):
            pytesseract.pytesseract.tesseract_cmd = exe
            break
    print("    No text layer, running OCR...")
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join(pytesseract.image_to_string(p.to_image(resolution=300).original)
                             for p in pdf.pages)
    except Exception as e:                         # Tesseract missing or failed
        print(f"    OCR failed: {e}")
        return ""


def one_line(s):
    return re.sub(r"\s+", " ", s or "").strip()


def find(pattern, text, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return one_line(m.group(1)) if m else None


def to_date(s):
    if not s:
        return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def to_amount(s):
    if not s:
        return None
    try:
        return float(s.replace("$", "").replace(",", "").strip("() "))
    except ValueError:
        return None


def parse_entries(text):
    """One entry per accounting page; the builder's page(s) ride along until the next one."""
    chunks = re.split(r"(?=BACKCHARGE\s*/\s*PO#)", text, flags=re.IGNORECASE)
    out = []
    for chunk in chunks:
        if not re.search(r"BACKCHARGE\s*/\s*PO#", chunk, re.IGNORECASE):
            continue
        # Accounting page ends at APPROVAL; everything after is the builder's paperwork
        split = re.split(r"\n\s*APPROVAL\s*\n", chunk, maxsplit=1, flags=re.IGNORECASE)
        acct, builder = split[0], (split[1] if len(split) > 1 else "")

        notes = re.search(r"NOTES:\s*(.*?)(?=ACTION\s+TO\s+BE\s+TAKEN|CHARGE\s+TO\s+SUB|$)",
                          acct, re.IGNORECASE | re.DOTALL)
        e = {
            "po": find(r"BACKCHARGE\s*/\s*PO#\s*(\d{4,10})", acct),
            "posting": to_date(find(r"\bDATE\s+(\d{1,2}/\d{1,2}/\d{2,4})", acct)),
            "division": find(r"DIVISION\s+([A-Z]+)", acct),
            "builder_pdf": find(r"BUILDER\s+NAME\s+(.+)", acct),
            "customer": (find(r"CUSTOMER\s*#\s*([A-Z0-9]+)", acct) or "").upper() or None,
            "address_pdf": find(r"JOB\s+ADDRESS\s+(.+)", acct),
            "amount": to_amount(find(r"AMOUNT\s+OF\s+BACK\s+CHARGE:?\s*(\(?\$?\s*[\d,]+\.?\d{0,2}\)?)", acct)),
            "job": (find(r"JOB\s*#\s*(N\d{5,9})", acct) or "").upper() or None,
            "inv": find(r"INV\s*#\s*(\d{5,12})", acct),
            "check": find(r"CHECK\s*#\s*([A-Z0-9-]+)", acct),
            "deposit": to_date(find(r"DEPOSIT\s+DATE:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", acct)),
            "notes": one_line(notes.group(1)) if notes else "",
            # Builder page (layout differs by builder; each field is optional)
            "builder_date": to_date(find(r"\bDate:\s*(\d{1,2}/\d{1,2}/\d{2,4})", builder)),
            "activity": find(r"Activity:\s*(.+)", builder),
            "builder_desc": find(r"Description:\s*(.+)", builder),
            "lot_pdf": find(r"Lot\s*#?\s*:?\s*([0-9]+[A-Z]?)\b", builder),
        }
        if e["builder_pdf"]:
            e["builder_pdf"] = re.split(r"\s{2,}", e["builder_pdf"])[0]
        if e["address_pdf"]:
            e["address_pdf"] = re.split(r"\s{2,}", e["address_pdf"])[0]
        out.append(e)
    return out


# --------------------------------------------------------------------------
# Suggestions learned from the Registry
# --------------------------------------------------------------------------
TOKEN = re.compile(r"[a-z0-9%/]+")
STOP = set("the a an and or of to in for at on is was be by bc it as from with that this due after "
           "had has have need needed were are back charge po".split())


def tokens(text):
    words = [w for w in TOKEN.findall(str(text).lower()) if w not in STOP and len(w) > 1]
    return words + [a + "_" + b for a, b in zip(words, words[1:])]


class Suggester:
    """Similar-note voting (TF-IDF cosine, k nearest Registry rows)."""

    def __init__(self, reg, lists):
        self.reg = reg.reset_index(drop=True)
        docs = [tokens(d) for d in self.reg["Description of Incident"]]
        df = Counter(t for d in docs for t in set(d))
        n = len(docs)
        self.idf = {t: math.log((n + 1) / (c + 1)) + 1 for t, c in df.items()}
        self.vecs = [self._vec(d) for d in docs]
        self.lists = lists

    def _vec(self, toks):
        tf = Counter(toks)
        v = {t: (1 + math.log(c)) * self.idf.get(t, 1.0) for t, c in tf.items()}
        norm = math.sqrt(sum(x * x for x in v.values())) or 1.0
        return {t: x / norm for t, x in v.items()}

    def neighbours(self, text):
        q = self._vec(tokens(text))
        sims = [(sum(w * v.get(t, 0.0) for t, w in q.items()), i) for i, v in enumerate(self.vecs)]
        sims.sort(reverse=True)
        return [(s, i) for s, i in sims[:SUGGEST_NEIGHBOURS] if s > 0.05]

    def _match_list(self, value, list_name):
        """Map a Registry value onto the exact drop-down spelling ('ROUGH' -> 'Rough')."""
        allowed = self.lists.get(list_name, [])
        for a in allowed:
            if a.strip().upper() == str(value).strip().upper():
                return a
        return None

    def suggest(self, text):
        nb = self.neighbours(text)
        result = {}
        for col, (reg_col, list_name) in SUGGESTED.items():
            votes = defaultdict(float)
            for s, i in nb:
                v = self.reg.at[i, reg_col]
                if pd.isna(v):
                    continue
                m = self._match_list(v, list_name)
                if m:
                    votes[m] += s
            if votes:
                ranked = sorted(votes.items(), key=lambda kv: -kv[1])
                total = sum(votes.values())
                result[col] = (ranked[0][0], [(k, round(100 * w / total)) for k, w in ranked[:3]])
        return result, nb


# --------------------------------------------------------------------------
# Build rows
# --------------------------------------------------------------------------
def build_rows(entries, prod, reg, lists, builder_ids):
    existing = set(reg["PO#/Backcharge #"].dropna().astype(str).str.strip().str.upper())
    person_by_job = (reg[reg["Responsible Person"].notna()]
                     .assign(job=lambda d: d["Job Number"].astype(str).str.strip().str.upper())
                     .groupby("job")["Responsible Person"].agg(lambda s: s.mode().iat[0]))
    sug = Suggester(reg, lists)

    rows, review, seen = [], [], set()
    for e in entries:
        po = f"BC-{e['po']}" if e["po"] else None
        if po and (po.upper() in existing or po.upper() in seen):
            review.append({"PO/Back Charge #": po, "Source PDF": e["file"], "Status": "Skipped",
                           "Notes": "Already in the Registry" if po.upper() in existing
                           else "Appears twice in the PDFs"})
            continue
        seen.add((po or "").upper())

        row = {c: None for c in COLUMNS}
        guessed, missing, notes = set(), set(), []
        job = e["job"]
        p = prod.loc[job] if job and job in prod.index else None
        if job and p is None:
            notes.append(f"Job {job} not found in PRODUCTION DATA")

        def pval(col):
            v = None if p is None else p.get(col)
            if v is None or pd.isna(v):
                return None
            if isinstance(v, float) and v.is_integer():    # lot 24.0 -> "24"
                v = int(v)
            return str(v).strip()

        row["Builder Name"] = pval("Builder") or builder_ids.get(e["customer"] or "") or e["builder_pdf"]
        row["Job #"] = job
        row["Job Address"] = pval("Address") or e["address_pdf"]
        row["Subdivision"] = pval("Subdivision")
        row["Lot Number"] = pval("Lot") or e["lot_pdf"]
        row["Supervisor"] = pval("Supervisor")
        row["BC Amount"] = e["amount"]
        row["Incident Date"] = e["builder_date"]
        if e["builder_date"]:
            guessed.add("Incident Date")
        row["Deduction Date"] = e["deposit"]
        row["Posting Date"] = e["posting"]
        row["PO/Back Charge #"] = po
        row["INV #"] = e["inv"]
        row["Check #"] = e["check"]
        row["Respondent"] = DEFAULT_RESPONDENT

        desc = e["notes"]
        bdesc = e["builder_desc"]
        if bdesc and one_line(bdesc).lower().rstrip(". ") not in desc.lower():
            desc = f"{desc} | Builder: {bdesc}" if desc else bdesc
        row["Description"] = desc or None

        learn_text = " ".join(filter(None, [e["notes"], bdesc, e["activity"]]))
        picks, nb = sug.suggest(learn_text) if learn_text else ({}, [])
        alt_text = []
        for col, (value, ranked) in picks.items():
            row[col] = value
            guessed.add(col)
            if len(ranked) > 1:
                alt_text.append(f"{col}: " + ", ".join(f"{k} {w}%" for k, w in ranked))
        if job and job in person_by_job.index:
            row["Responsible Person"] = str(person_by_job[job]).strip()
            guessed.add("Responsible Person")
            notes.append("Responsible Person taken from an earlier back charge on this job")

        for col in ("Builder Name", "Job #", "Job Address", "BC Amount", "Deduction Date",
                    "PO/Back Charge #", "Description"):
            if row[col] in (None, ""):
                missing.add(col)
        for col in SUGGESTED:
            if row[col] is None:
                missing.add(col)

        similar = [one_line(str(reg.reset_index(drop=True).at[i, "Description of Incident"]))[:70]
                   for _, i in nb[:2]]
        rows.append((row, guessed, missing))
        review.append({
            "PO/Back Charge #": po, "Source PDF": e["file"], "Status": "Ready to import",
            "Activity (builder page)": e["activity"],
            "Close calls": "; ".join(alt_text),
            "Most similar past notes": " || ".join(similar),
            "Notes": "; ".join(notes),
        })
    return rows, review


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def write_output(rows, review, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Back Charge Data"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for row, guessed, missing in rows:
        ws.append([row[c] for c in COLUMNS])
        r = ws.max_row
        for j, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r, column=j)
            if col in DATE_COLUMNS and cell.value:
                cell.number_format = "m/d/yyyy"
            if col == "BC Amount" and cell.value is not None:
                cell.number_format = "$#,##0.00"
            if col in missing:
                cell.fill = MISSING_FILL
            elif col in guessed:
                cell.fill = GUESS_FILL
    for j, col in enumerate(COLUMNS, start=1):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = \
            60 if col == "Description" else max(12, len(col) + 2)
    ws.freeze_panes = "A2"

    rv = wb.create_sheet("Review")
    keys = ["PO/Back Charge #", "Status", "Source PDF", "Activity (builder page)",
            "Close calls", "Most similar past notes", "Notes"]
    rv.append(keys)
    for c in rv[1]:
        c.font = Font(bold=True)
    for item in review:
        rv.append([item.get(k) for k in keys])
    for j, w in enumerate([16, 16, 34, 28, 60, 80, 50], start=1):
        rv.column_dimensions[rv.cell(row=1, column=j).column_letter].width = w

    key = wb.create_sheet("Key")
    key.append(["Colour", "Meaning"])
    key.append(["Yellow", "Suggested from past back charges or the builder's page. Confirm or change it."])
    key.append(["Orange", "Could not be found. Fill it in."])
    key["A2"].fill, key["A3"].fill = GUESS_FILL, MISSING_FILL
    key.column_dimensions["B"].width = 80
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Read accounting back charge PDFs into Back Charge Data order.")
    ap.add_argument("--folder", default=LOADING_FOLDER, help="Folder holding the accounting PDFs")
    ap.add_argument("--workbook", default=TRACKING_WORKBOOK, help="The back charge tracking workbook")
    ap.add_argument("--out", default=None, help="Output file (default: BC_Import_Ready.xlsx in the folder)")
    args = ap.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"Folder not found: {folder}\nCheck that the P: drive is connected.")
    if not Path(args.workbook).exists():
        sys.exit(f"Tracking workbook not found: {args.workbook}")
    out = Path(args.out) if args.out else folder / OUTPUT_FILE_NAME

    prod, reg, lists, builder_ids = load_reference(args.workbook)

    entries = []
    pdfs = sorted(f for f in folder.iterdir() if f.suffix.lower() == ".pdf")
    print(f"\nFound {len(pdfs)} PDF file(s) in {folder}")
    for f in pdfs:
        found = parse_entries(pdf_text(f))
        for e in found:
            e["file"] = f.name
        print(f"  {f.name}: {len(found)} back charge(s)")
        entries.extend(found)

    rows, review = build_rows(entries, prod, reg, lists, builder_ids)
    try:
        write_output(rows, review, out)
    except PermissionError:
        sys.exit(f"\nCould not save {out}. Close it in Excel and run again.")

    skipped = sum(1 for r in review if r["Status"] == "Skipped")
    guesses = sum(len(g) for _, g, _ in rows)
    gaps = sum(len(m) for _, _, m in rows)
    print(f"\nReady to import: {len(rows)}   Skipped (already in Registry or repeated): {skipped}")
    print(f"Cells to confirm (yellow): {guesses}   Cells to fill (orange): {gaps}")
    print(f"Saved {out}")


if __name__ == "__main__":
    main()
