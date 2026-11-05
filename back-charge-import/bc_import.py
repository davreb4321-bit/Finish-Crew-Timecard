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

Each back charge's pages (accounting sheet + the builder's pages behind it) are
also saved as their own PDF in the "Accounting Sheets" folder inside the
loading folder, including the ones already in the Registry. The PDF macro
files them into the house's G: drive folder.

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
import pypdfium2 as pdfium    # installed with pdfplumber; used to split the PDFs
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pyxlsb import open_workbook

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
LOADING_FOLDER = r"P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER"
TRACKING_WORKBOOK = os.path.join(LOADING_FOLDER, "2026 BACK CHARGE TRACKING.xlsb")
OUTPUT_FILE_NAME = "BC_Import_Ready.xlsx"
SPLIT_FOLDER_NAME = "Accounting Sheets"   # inside the loading folder

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
# Glossary (keep in step with the glossary in BackChargeImport.bas)
# --------------------------------------------------------------------------
# Builders: (name used in the Registry / PRODUCTION DATA, [text that identifies it]).
# A builder matches when its text is CONTAINED anywhere in the name being read, after
# capitals, punctuation and extra spaces are ignored ("M.I. Homes" -> "M I HOMES").
# Longer, more specific entries go first.
BUILDERS = [
    ("PULTE HOMES OF INDIANA", ["PULTE"]),
    ("M.I. HOMES", ["M I HOMES", "MI HOMES", "M I SCHOTTENSTEIN"]),
    ("ARBOR HOMES", ["ARBOR"]),
    ("FISCHER MANAGEMENT", ["FISCHER", "FISHER HOMES", "FISHER MANAGEMENT"]),
    ("EC NEW VISION INDIANA", ["EC NEW VISION", "NEW VISION"]),
    ("LENNAR", ["LENNAR"]),
    ("BEAZER", ["BEAZER"]),
    ("RYAN HOMES", ["RYAN HOMES"]),
    ("DREES PREMIER HOMES", ["DREES"]),
    ("D. R. HORTON", ["HORTON"]),
    ("DAVID WEEKLEY", ["WEEKLEY"]),
    ("ESTRIDGE HOMES", ["ESTRIDGE"]),
    ("SILVERTHORNE HOMES", ["SILVERTHORNE"]),
    ("TAYLOR MORRISON", ["TAYLOR MORRISON"]),
    ("DAVIS BUILDING GROUP LLC", ["DAVIS BUILDING", "DAVIS BLDG"]),   # not "DAVIS": Paul Davis is a restoration company
    ("TWG WHISLETT GROUP", ["WHISLETT", "TWG"]),
    ("WESTPORT", ["WESTPORT"]),
    ("BEACON BUILDERS LLC", ["BEACON BUILDERS"]),
    ("BEDROCK BUILDERS, INC.", ["BEDROCK"]),
    ("Old Town Design Group", ["OLD TOWN"]),
    ("PRIDE ONE", ["PRIDE ONE"]),
    ("JOYNER HOMES LLC", ["JOYNER"]),
    ("ONYX AND EAST LLC", ["ONYX"]),
]

# Street words: every spelling on the left of each list means the word on the right.
STREET_WORDS = {
    "LANE": ["LN", "LA"], "STREET": ["ST", "STR"], "DRIVE": ["DR", "DRV"],
    "COURT": ["CT", "CRT"], "AVENUE": ["AVE", "AV"], "ROAD": ["RD"], "PLACE": ["PL"],
    "CIRCLE": ["CIR", "CIRC"], "BOULEVARD": ["BLVD", "BL"], "TRAIL": ["TRL", "TR"],
    "PARKWAY": ["PKWY", "PKY"], "TERRACE": ["TER", "TERR"], "WAY": ["WY"],
    "POINTE": ["PT", "POINT"], "CROSSING": ["XING", "CRSG"], "COVE": ["CV"],
    "SQUARE": ["SQ"], "HIGHWAY": ["HWY"], "BEND": ["BND"], "GLEN": ["GLN"],
    "HOLLOW": ["HOLW"], "CREEK": ["CRK"], "RIDGE": ["RDG"], "COMMONS": ["CMNS"],
    "MEADOWS": ["MDWS"], "TRACE": ["TRCE"], "RUN": [], "PASS": [], "PIKE": [], "LOOP": [],
    "ROW": [], "WALK": [], "PATH": [], "VIEW": ["VW"], "CHASE": [],
    "NORTH": ["N"], "SOUTH": ["S"], "EAST": ["E"], "WEST": ["W"],
}
DIRECTIONS = {"NORTH", "SOUTH", "EAST", "WEST"}
_STREET_LOOKUP = {v: k for k, vs in STREET_WORDS.items() for v in vs + [k]}


def plain(text):
    """Upper case, punctuation to spaces, single spaces."""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9&]+", " ", str(text or "").upper())).strip()


def builder_name(*texts):
    """First glossary builder CONTAINED in any of the texts -> Registry spelling."""
    for t in texts:
        t = f" {plain(t)} "
        if not t.strip():
            continue
        for name, variants in BUILDERS:
            if any(f" {plain(v)} " in t for v in variants):
                return name
    return None


def address_key(text):
    """'793 Rosebud Ln.', '793 ROSEBUD LANE' and '793 Rosebud' -> '793 ROSEBUD'.
    House number + street name. Street-type words at the end and a direction
    right after the number are dropped, so a missing or abbreviated 'Lane' still
    matches. At least one street word is always kept."""
    words = [_STREET_LOOKUP.get(w, w) for w in plain(text).split()]
    if len(words) < 2 or not words[0][0].isdigit():
        return None
    num, street = words[0], words[1:]
    if len(street) > 1 and street[0] in DIRECTIONS:
        street = street[1:]
    while len(street) > 1 and street[-1] in STREET_WORDS:
        street = street[:-1]
    return f"{num} {' '.join(street)}"


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
    prod["_addr"] = prod["Address"].map(address_key)
    prod["_builder"] = prod["Builder"].map(lambda b: builder_name(b) or plain(b))
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
def pdf_pages(path):
    """Text of each page (OCR when the PDF has no text layer)."""
    with pdfplumber.open(path) as pdf:
        pages = [page.extract_text(layout=True) or "" for page in pdf.pages]
    if any(p.strip() for p in pages):
        return pages
    return ocr_pages(path)


def ocr_pages(path):
    """Fallback for scanned PDFs. Needs pytesseract + Tesseract installed."""
    try:
        import pytesseract
    except ImportError:
        print("    No text layer and pytesseract is not installed; skipping.")
        return []
    for exe in (r"C:\Users\DMartin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(exe):
            pytesseract.pytesseract.tesseract_cmd = exe
            break
    print("    No text layer, running OCR...")
    try:
        with pdfplumber.open(path) as pdf:
            return [pytesseract.image_to_string(p.to_image(resolution=300).original)
                    for p in pdf.pages]
    except Exception as e:                         # Tesseract missing or failed
        print(f"    OCR failed: {e}")
        return []


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


ACCOUNTING_START = re.compile(r"BACKCHARGE\s*/\s*PO#", re.IGNORECASE)


def parse_entries(pages):
    """One entry per accounting sheet. A sheet starts on a page with BACKCHARGE/PO#; the
    builder's page(s) after it belong to it until the next one starts."""
    groups = []
    for i, page in enumerate(pages):
        if ACCOUNTING_START.search(page):
            groups.append([i])
        elif groups:
            groups[-1].append(i)
    out = []
    for g in groups:
        e = parse_chunk("\n".join(pages[i] for i in g))
        e["pages"] = g
        out.append(e)
    return out


def po_number(acct):
    """Number after BACKCHARGE/PO#, without the BC- prefix (added later):
    1048000 -> 1048000, INV-15299503 / INV 15299503 / BC-INV-15299503 -> INV-15299503,
    1234567-01 -> 1234567-01."""
    m = re.search(r"BACKCHARGE\s*/\s*PO#\s*(?:BC\s*-\s*)?(INV\s*-?\s*)?(\d{4,10}(?:-\d{1,3})?)\b",
                  acct, re.IGNORECASE)
    if not m:
        return None
    return ("INV-" if m.group(1) else "") + m.group(2)


def parse_chunk(chunk):
    # Accounting page ends at APPROVAL; everything after is the builder's paperwork
    split = re.split(r"\n\s*APPROVAL\s*\n", chunk, maxsplit=1, flags=re.IGNORECASE)
    acct, builder = split[0], (split[1] if len(split) > 1 else "")

    notes = re.search(r"NOTES:\s*(.*?)(?=ACTION\s+TO\s+BE\s+TAKEN|CHARGE\s+TO\s+SUB|$)",
                      acct, re.IGNORECASE | re.DOTALL)
    e = {
        "po": po_number(acct),
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
    return e


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
def find_house(e, prod, builder_ids):
    """PRODUCTION DATA row for the back charge: by job #, else by address (+ builder).
    Returns (row or None, job #, builder from the PDF, notes, job found by address?)."""
    notes = []
    job = e["job"]
    pdf_builder = builder_name(builder_ids.get(e["customer"] or ""), e["builder_pdf"], e["file"])
    p = prod.loc[job] if job and job in prod.index else None
    by_address = False
    if p is None:
        key = address_key(e["address_pdf"])
        hits = prod[prod["_addr"] == key] if key else prod.iloc[0:0]
        if pdf_builder and len(hits) > 1:
            hits = hits[hits["_builder"] == pdf_builder]
        if len(hits) == 1:
            p = hits.iloc[0]
            notes.append(f"Job {job or '(none on PDF)'} not in PRODUCTION DATA; "
                         f"matched by address to {hits.index[0]}")
            job, by_address = hits.index[0], True
        else:
            notes.append(f"Job {job or '(none on PDF)'} not found in PRODUCTION DATA"
                         + (f"; {len(hits)} jobs share this address" if len(hits) > 1 else ""))
    return p, job, pdf_builder, notes, by_address


def prod_value(p, col):
    v = None if p is None else p.get(col)
    if v is None or pd.isna(v):
        return None
    if isinstance(v, float) and v.is_integer():    # lot 24.0 -> "24"
        v = int(v)
    return str(v).strip()


def house_record(e, prod, builder_ids, status):
    """One row of the Accounting Sheets tab: where the split PDF should be filed."""
    p, job, pdf_builder, _, _ = find_house(e, prod, builder_ids)
    po = f"BC-{e['po']}" if e["po"] else None
    return {
        "PO/Back Charge #": po, "Status": status, "Split PDF": e.get("split_pdf"),
        "Builder Name": prod_value(p, "Builder") or pdf_builder or e["builder_pdf"],
        "Job #": job, "Job Address": prod_value(p, "Address") or e["address_pdf"],
        "Subdivision": prod_value(p, "Subdivision"),
        "Lot Number": prod_value(p, "Lot") or e["lot_pdf"],
        "Deduction Date": e["deposit"],
    }


def build_rows(entries, prod, reg, lists, builder_ids):
    existing = set(reg["PO#/Backcharge #"].dropna().astype(str).str.strip().str.upper())
    person_by_job = (reg[reg["Responsible Person"].notna()]
                     .assign(job=lambda d: d["Job Number"].astype(str).str.strip().str.upper())
                     .groupby("job")["Responsible Person"].agg(lambda s: s.mode().iat[0]))
    sug = Suggester(reg, lists)

    rows, review, sheets, seen = [], [], [], set()
    for e in entries:
        po = f"BC-{e['po']}" if e["po"] else None
        if po and (po.upper() in existing or po.upper() in seen):
            reason = "Already in the Registry" if po.upper() in existing else "Appears twice in the PDFs"
            review.append({"PO/Back Charge #": po, "Source PDF": e["file"], "Status": "Skipped",
                           "Notes": reason})
            if po.upper() in existing:     # still needs its accounting sheet filed
                sheets.append(house_record(e, prod, builder_ids, reason))
            continue
        seen.add((po or "").upper())
        sheets.append(house_record(e, prod, builder_ids, "Imported"))

        row = {c: None for c in COLUMNS}
        guessed, missing = set(), set()
        p, job, pdf_builder, notes, by_address = find_house(e, prod, builder_ids)
        if by_address:
            guessed.add("Job #")

        def pval(col):
            return prod_value(p, col)

        row["Builder Name"] = pval("Builder") or pdf_builder or e["builder_pdf"]
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
    return rows, review, sheets


def make_split_folder(folder):
    """The Accounting Sheets folder. If Windows won't create it (a file with that name, or a
    just-deleted folder still held open on the network drive), use a dated one instead."""
    out_folder = folder / SPLIT_FOLDER_NAME
    try:
        out_folder.mkdir(exist_ok=True)
        return out_folder
    except (FileExistsError, PermissionError) as err:
        alt = folder / f"{SPLIT_FOLDER_NAME} {datetime.now():%Y-%m-%d %H%M}"
        print(f"\nCould not use {out_folder} ({err.strerror}).")
        print(f"Saving the accounting sheets in {alt} instead.")
        alt.mkdir(exist_ok=True)
        return alt


def split_sheets(entries, folder, out_folder):
    """Save each back charge's pages as its own PDF. Sets e["split_pdf"]."""
    by_file = defaultdict(list)
    for e in entries:
        by_file[e["file"]].append(e)
    made = 0
    for fname, es in by_file.items():
        src = pdfium.PdfDocument(str(folder / fname))
        try:
            for e in es:
                po = f"BC-{e['po']}" if e["po"] else Path(fname).stem
                addr = plain(e["address_pdf"]) or "NO ADDRESS"
                name = re.sub(r'[\\/:*?"<>|]', "-", f"{po} - {addr} - ACCOUNTING SHEET.pdf")
                path = out_folder / name
                doc = pdfium.PdfDocument.new()
                doc.import_pages(src, e["pages"])
                doc.save(str(path))
                doc.close()
                e["split_pdf"] = str(path.resolve())
                made += 1
        finally:
            src.close()
    return made


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def write_output(rows, review, sheets, path):
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

    acc = wb.create_sheet("Accounting Sheets")
    cols = ["PO/Back Charge #", "Status", "Split PDF", "Builder Name", "Job #", "Job Address",
            "Subdivision", "Lot Number", "Deduction Date"]
    acc.append(cols)
    for c in acc[1]:
        c.font = Font(bold=True)
    for item in sheets:
        acc.append([item.get(k) for k in cols])
        if item.get("Deduction Date"):
            acc.cell(row=acc.max_row, column=9).number_format = "m/d/yyyy"
    for j, w in enumerate([16, 24, 90, 24, 12, 26, 20, 10, 14], start=1):
        acc.column_dimensions[acc.cell(row=1, column=j).column_letter].width = w

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
        found = parse_entries(pdf_pages(f))
        for e in found:
            e["file"] = f.name
        print(f"  {f.name}: {len(found)} back charge(s)")
        entries.extend(found)

    split_folder = make_split_folder(folder)
    split_count = split_sheets(entries, folder, split_folder)
    rows, review, sheets = build_rows(entries, prod, reg, lists, builder_ids)
    try:
        write_output(rows, review, sheets, out)
    except PermissionError:
        sys.exit(f"\nCould not save {out}. Close it in Excel and run again.")

    skipped = sum(1 for r in review if r["Status"] == "Skipped")
    guesses = sum(len(g) for _, g, _ in rows)
    gaps = sum(len(m) for _, _, m in rows)
    print(f"\nReady to import: {len(rows)}   Skipped (already in Registry or repeated): {skipped}")
    print(f"Cells to confirm (yellow): {guesses}   Cells to fill (orange): {gaps}")
    print(f"Accounting sheets split into their own PDFs: {split_count} "
          f"(in {split_folder})")
    print(f"Saved {out}")


if __name__ == "__main__":
    main()
