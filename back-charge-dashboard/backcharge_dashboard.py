#!/usr/bin/env python3
"""
Back Charge Dashboard
=====================

Reads the "Registry" sheet of the back-charge tracking workbook (.xlsb or .xlsx),
sorts every "Description of Incident" note (column K) into a root-cause theme,
and writes:

  * back_charge_dashboard.html  - a self-contained interactive dashboard
                                  (open it in any browser, filters at the top)
  * back_charge_analysis.xlsx   - the categorized registry plus summary tabs,
                                  so you can review / audit how each note was sorted

Usage
-----
    pip install pandas pyxlsb openpyxl
    python backcharge_dashboard.py                       (uses DEFAULT_WORKBOOK below)
    python backcharge_dashboard.py "other_workbook.xlsb"
    python backcharge_dashboard.py --out "C:/Reports"

Tuning
------
The THEMES list below drives the note classification. Each theme has a list of
regular expressions; a note goes to the FIRST theme whose pattern matches. To
fix a mis-sorted note, add a keyword to the right theme (or move a theme higher).
Check the "Registry (Categorized)" tab of the Excel output to see the result.
"""

import argparse
import html
import json
import re
from datetime import date
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
DEFAULT_WORKBOOK = r"P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER\2026 BACK CHARGE TRACKING.xlsb"
SHEET_NAME = "Registry"
HEADER_ROW = 9          # Excel row number that holds the column headings

COL = {                 # headings as they appear in row 9
    "id": "PO#/Backcharge #",
    "respondent": "Respondent",
    "address": "Job Address",
    "incident_date": "Backcharge Incident Date",
    "claim_type": "Claim Type",
    "dept": "Responsible Department",
    "incident_type": "Incident Type",
    "property": "Type of Property Damaged",
    "desc": "Description of Incident",
    "phase": "Phase of Incident",
    "person": "Responsible Person",
    "invoice_date": "Backcharge Invoice Date",
    "cost": "Invoiced Cost",
    "job": "Job Number",
    "payee": "Payee",
    "reason": "Non-Payment Reason",
}

# --------------------------------------------------------------------------
# Root-cause themes (first match wins - order matters)
# Patterns are matched against: Incident Type + " | " + Description (case-insensitive)
# --------------------------------------------------------------------------
THEMES = [
    ("Not installer fault",
     "Homeowner maintenance, plumber, manufacturer, weather, builder/billing errors",
     [r"h/o maintenance", r"dirty filter", r"plumber'?s fault", r"sump pump", r"water heater leaking",
      r"floor drain advised", r"not airtron install", r"blown over", r"overpaid", r"cm error",
      r"builder error", r"no airtron responsibility", r"due to humidifier age", r"design issue"]),
    ("Equipment failure / manufacturer defect",
     "Failed equipment (compressor, coil pan, controls) and the damage it caused",
     [r"^manufacturer defect", r"malfunction", r"compressor", r"not shutting off"]),
    ("Humidifier / SharkBite leak",
     "Humidifier fittings or drain hoses leaking",
     [r"humidifier", r"shark ?bite"]),
    ("Flue / exhaust pipe leak",
     "Unglued or loose PVC joints, elbows, concentric kits, pipe boots",
     [r"^flue leak", r"^(?!j-block).*flue", r"exhaust (pvc )?pipe", r"not glued", r"wasn.?t glue", r"glued",
      r"concentric", r"pipe boot", r"^(?!range vent|ductwork).*\belbow\b", r"reducer",
      r"^(?!drain).*hvac (boot )?leak", r"hvac pipe leak",
      r"vent stack\.$", r"roof leak"]),
    ("Condensate drain / pan leak",
     "Condensate lines, pans, traps and float switches",
     [r"condensate", r"drain", r"\bpan\b", r"p-?trap", r"j-?trap", r"float switch", r"^drain 90%",
      r"^drain ac"]),
    ("Line set (kinks, brazes, leaks)",
     "Kinked, leaking, badly brazed, missing or mis-located line sets",
     [r"line ?set", r"braze", r"copper line", r"set line"]),
    ("Duct leakage / failed duct or blower test",
     "Duct blaster, TDL, blower door and leakage retests",
     [r"duct blast", r"duct blaser", r"duct leakage", r"duct test", r"blower", r"\btdl\b", r"retest",
      r"^failed duct test", r"too much duct", r"smoke bomb", r"\bpeg\b"]),
    ("Range / microwave hood vent",
     "Hood vent stack missing, short, too low, off-center or cut-out not done",
     [r"range", r"\bhood", r"microwave", r"wall ?stack", r"vent stack", r"\bcollar\b", r"canopy",
      r"\brv\b", r"micro vent", r"chef hood", r"exhaust vent too low", r"\bstack\b", r"\brbr\b"]),
    ("Exterior vent misaligned (siding / roof)",
     "Crooked or mis-located wall/roof terminations damaging siding, soffit, brick or roofing",
     [r"^j-block", r"siding", r"soffit", r"brick", r"roof (boot|vent|labor)", r"box vent",
      r"exhaust penetration", r"flashing", r"termination"]),
    ("Ductwork missed / disconnected at rough",
     "Runs not installed, not connected, crushed, torn or not per print",
     [r"miss(ed|ing) (heat run|supply|a supply)", r"not (be )?installed", r"never installed",
      r"not connected", r"disconnect", r"not attached", r"open\b.*ports", r"not secure",
      r"not extend", r"did not install", r"did not get installed", r"not per print", r"per print",
      r"crushed", r"tear in duct", r"water in duct", r"unfinished work", r"not sealed",
      r"kinked (flex|duct)", r"re-?stretched", r"debris blocking", r"not tied in", r"missing supply",
      r"kitchen return", r"registers not installed", r"insulation not properly", r"flex insulation",
      r"duct leak", r"leaking (hvac )?duct", r"armaflex", r"reattach", r"resecure", r"hanging",
      r"coming undone", r"falling down", r"tighten", r"improper dryer vent connection"]),
    ("Framing over-cut / structural damage",
     "Cut trusses, joists, top plates, studs, blocking, subfloor or sheathing",
     [r"over ?cut", r"over notched", r"truss", r"i joist", r"top plate", r"cut.{0,15}blocking", r"blocking.{0,25}cut",
      r"subfloor", r"sheathing", r"osb", r"drilled holes", r"double (up|all|out)", r"framer",
      r"framing", r"\bbow", r"board and batten", r"cutting out 2x6"]),
    ("Fixture in wrong location / not centered",
     "Boots, registers, returns, dryer vents, thermostats placed wrong and moved",
     [r"wrong (location|spot)", r"incorrect location", r"too close", r"not centered", r"off.?center",
      r"not even", r"in the way", r"in way of", r"\bmoved", r"\bmove\b", r"reloca", r"to close to",
      r"installed wrong", r"unneeded", r"bulk ?head", r"middle of the room",
      r"under a cabinet", r"not fitting", r"set incorrectly", r"raised", r"too low", r"too large",
      r"dryer vent", r"thermostat", r"t-stat"]),
    ("Failed inspection (other / undefined)",
     "Failed rough/final inspections without a specific item, safety findings at inspection",
     [r"fail", r"inspection", r"could not inspect", r"exhaust fans"]),
    ("Crew damage & jobsite conduct",
     "Physical damage by crews, clean-up, safety fines, meter pits, deliveries",
     [r"clean ?up", r"cleaning", r"not cleaning", r"trash", r"debris", r"tripped", r"dropped",
      r"stepp?ed", r"stepping", r"backed into", r"broken", r"window", r"spray paint", r"glue on",
      r"pvc glue", r"mastic", r"meter pit", r"ran over", r"run over", r"silt fence", r"hard hat",
      r"safety", r"parked", r"hit the wall", r"screw", r"gauge", r"damaged", r"cracking",
      r"radon", r"metal tape"]),
    ("Late / back-order / return trips",
     "Schedule slips, back-ordered equipment, permits, late crews",
     [r"\blate\b", r"back ?order", r"permit", r"delay", r"trip charge", r"return trip",
      r"not set", r"securing"]),
    ("HVAC repair access cuts",
     "Drywall cut so HVAC could reach or fix something; the root cause is not described",
     [r"cut", r"access", r"hvac (repairs|eval|issues)", r"service", r"repair(s|ed)? (on|to|from)",
      r"leak", r"mini.?split", r"furnace"]),
    ("Touch-up / no cause given",
     "Generic drywall or paint touch-ups and notes too vague to identify a cause",
     [r".*"]),
]

# Phrases that signal weak documentation (flagged separately from the theme)
VAGUE_PATTERNS = [r"undefined", r"not defined", r"no description", r"no photo", r"no pictures",
                  r"no backup", r"no other documentation", r"not enough info", r"no documentation",
                  r"unidentified", r"unsure", r"uncertain"]

# Things the incident cost had to fix (multi-label)
CONSEQUENCES = [
    ("Drywall", [r"drywal", r"drywawll", r"ceiling", r"bulk ?head", r"patch", r"hole", r"sheets of"]),
    ("Paint / touch-up", [r"paint", r"touch ?up", r"texture", r"caulk"]),
    ("Water damage / mold", [r"leak", r"water", r"mold", r"remediation", r"microbial", r"moisture",
                             r"condensation", r"wet"]),
    ("Flooring / trim", [r"floor", r"carpet", r"lvp", r"trim", r"crown", r"tile", r"backsplash",
                         r"back-splash", r"door jamb"]),
    ("Siding / roof / exterior", [r"siding", r"roof", r"soffit", r"brick", r"shingle", r"gutter",
                                  r"flashing"]),
    ("Re-inspection / retest fee", [r"reinspect", r"re-inspect", r"retest", r"re-test", r"fail"]),
    ("Other trade return trip", [r"trip", r"return", r"rbr", r"framer", r"electrician", r"painter",
                                 r"plumb", r"cabinet", r"roofer"]),
]

# Themes counted as water leaks in the findings and the "Water-leak share" tile
LEAK_THEMES = ["Flue / exhaust pipe leak", "Condensate drain / pan leak", "Humidifier / SharkBite leak"]

PAYEE_ALIASES = {
    "PULTE HOMES OF INDIANAPOLIS": "PULTE HOMES OF INDIANA",
    "SILVERTHORNE": "SILVERTHORNE HOMES",
}


# --------------------------------------------------------------------------
# Loading & cleaning
# --------------------------------------------------------------------------
def excel_date(v):
    """Excel serial number -> date (ignores obviously bad serials)."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        if isinstance(v, (pd.Timestamp, date)):
            return pd.Timestamp(v).date()
        return None
    if pd.isna(f) or f < 40000 or f > 60000:   # blank, or before 2009 / after 2064 -> typo
        return None
    return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(f))).date()


def clean_text(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def up(v):
    return clean_text(v).upper()


def classify_theme(text):
    t = text.lower()
    for name, _, pats in THEMES:
        if any(re.search(p, t) for p in pats):
            return name
    return "Touch-up / no cause given"


def consequences(text):
    t = text.lower()
    found = [name for name, pats in CONSEQUENCES if any(re.search(p, t) for p in pats)]
    return found or ["Not stated"]


def recovery_bucket(reason):
    r = up(reason)
    if not r:
        return "No reason recorded"
    if r == "EMPLOYEE":
        return "Employee (no sub deduction)"
    if "PENDING" in r or "DISPUTED" in r:
        return "Pending / disputed"
    if any(k in r for k in ("NOT ENOUGH", "UNKNOWN", "NO PICTURES", "UNCERTAIN", "NOT ENOUGH INFO")):
        return "Insufficient evidence"
    if any(k in r for k in ("NO LONGER WITH", "TOO OLD", "NOT SIGNIFICANT", "LATE TO THE JOB")):
        return "Not pursued (aged / left / small)"
    if r.startswith("BC ") or "BACK CHARGE" in r or "BACK CHARGED" in r or "BC IS GOING" in r:
        return "Passed to named installer"
    if any(k in r for k in ("NOT THE FAULT", "NOT FAULT", "DESIGN", "ESTIMATING", "EXISTING",
                            "COLLECTOR", "ALL TRADES", "REVERSED", "BEING WORKED OUT")):
        return "Not installer fault"
    return "Other"


def load_registry(path):
    path = Path(path)
    engine = "pyxlsb" if path.suffix.lower() == ".xlsb" else None
    raw = pd.read_excel(path, sheet_name=SHEET_NAME, header=HEADER_ROW - 1, engine=engine)
    raw.columns = [clean_text(c) for c in raw.columns]
    missing = [v for v in COL.values() if v not in raw.columns]
    if missing:
        raise SystemExit(f"Missing expected columns in row {HEADER_ROW}: {missing}")
    raw = raw[raw[COL["desc"]].notna() | raw[COL["id"]].notna()].copy()

    df = pd.DataFrame()
    df["id"] = raw[COL["id"]].map(clean_text)
    df["respondent"] = raw[COL["respondent"]].map(clean_text)
    df["address"] = raw[COL["address"]].map(up)
    df["job"] = raw[COL["job"]].map(up)
    df["incident_date"] = raw[COL["incident_date"]].map(excel_date)
    df["invoice_date"] = raw[COL["invoice_date"]].map(excel_date)
    df["claim_type"] = raw[COL["claim_type"]].map(lambda v: clean_text(v).title() or "(Blank)")
    df["dept"] = raw[COL["dept"]].map(lambda v: up(v) or "(NOT RECORDED)")
    df["incident_type"] = raw[COL["incident_type"]].map(lambda v: clean_text(v) or "(Not recorded)")
    df["property"] = raw[COL["property"]].map(
        lambda v: (clean_text(v).replace("Insullation", "Insulation").replace("Carpert", "Carpet"))
        or "(Not recorded)")
    df["desc"] = raw[COL["desc"]].map(clean_text)
    df["phase"] = raw[COL["phase"]].map(
        lambda v: up(v) if isinstance(v, str) and up(v) else "(NOT RECORDED)")
    df["person"] = raw[COL["person"]].map(lambda v: up(v) or "(NOT RECORDED)")
    df["cost"] = pd.to_numeric(raw[COL["cost"]], errors="coerce").fillna(0.0).round(2)
    df["payee"] = raw[COL["payee"]].map(lambda v: PAYEE_ALIASES.get(up(v), up(v)) or "(NOT RECORDED)")
    df["reason"] = raw[COL["reason"]].map(clean_text)

    text = df["incident_type"] + " | " + df["desc"]
    df["theme"] = text.map(classify_theme)
    df["consequences"] = text.map(consequences)
    df["vague"] = text.str.lower().map(lambda t: any(re.search(p, t) for p in VAGUE_PATTERNS)) | \
        (df["desc"].str.split().str.len() <= 3)
    df["recovery"] = df["reason"].map(recovery_bucket)
    df["month"] = df["incident_date"].map(lambda d: d.strftime("%Y-%m") if d else "")
    df["days_to_invoice"] = [
        (b - a).days if a and b and (b - a).days >= 0 else None
        for a, b in zip(df["incident_date"], df["invoice_date"])
    ]
    return df.reset_index(drop=True)


# --------------------------------------------------------------------------
# Summaries (used for the Excel tabs and the written findings)
# --------------------------------------------------------------------------
def summarize(df, key):
    g = df.groupby(key).agg(Count=("cost", "size"), Total_Cost=("cost", "sum"),
                            Avg_Cost=("cost", "mean"))
    g["Pct_Count"] = g["Count"] / g["Count"].sum()
    g["Pct_Cost"] = g["Total_Cost"] / g["Total_Cost"].sum()
    return g.sort_values("Total_Cost", ascending=False).round(2)


def money(v):
    return f"${v:,.0f}"


def findings(df):
    """Plain-language findings computed from the full data set."""
    out = []
    n, total = len(df), df["cost"].sum()
    th = summarize(df, "theme")

    leak_themes = LEAK_THEMES
    leaks = df[df["theme"].isin(leak_themes)]
    out.append(
        f"Water leaks from installed equipment (flue/exhaust pipes, condensate drains, humidifiers) are "
        f"{len(leaks)} of {n} back charges ({len(leaks) / n:.0%}) but {money(leaks['cost'].sum())} of "
        f"{money(total)} ({leaks['cost'].sum() / total:.0%}). They cost {money(leaks['cost'].mean())} "
        f"each on average versus {money(df['cost'].mean())} overall, because a leak turns into "
        f"drywall, flooring and mold remediation.")

    rough = df[df["theme"].isin(["Range / microwave hood vent", "Ductwork missed / disconnected at rough",
                                 "Fixture in wrong location / not centered",
                                 "Exterior vent misaligned (siding / roof)",
                                 "Framing over-cut / structural damage",
                                 "Duct leakage / failed duct or blower test",
                                 "Failed inspection (other / undefined)"])]
    out.append(
        f"Rough-in workmanship and layout (missed or disconnected ducts, vents in the wrong place, "
        f"hood stacks, crooked terminations, failed duct/inspection tests, over-cut framing) is the "
        f"largest group by count: {len(rough)} incidents ({len(rough) / n:.0%}), "
        f"{money(rough['cost'].sum())}. These are mostly lower-dollar but highly repeatable, which "
        f"makes them the best target for a rough-in checklist.")

    top = th.sort_values("Count", ascending=False).head(3)
    out.append("Most frequent themes from the notes: " + "; ".join(
        f"{k} ({int(r.Count)}, {money(r.Total_Cost)})" for k, r in top.iterrows()) + ".")

    hood = df[df["theme"] == "Range / microwave hood vent"]
    if len(hood):
        out.append(
            f"Range/microwave hood vents alone produced {len(hood)} back charges "
            f"({money(hood['cost'].sum())}). Notes repeat the same few causes: wall stack not installed "
            f"or too short, vent too low, collar off-center, or the cut-out not made at rough.")

    known = df[df["person"] != "(NOT RECORDED)"]
    ps = summarize(known, "person").sort_values("Count", ascending=False)
    if len(ps):
        p1 = ps.index[0]
        p1_top = known[known["person"] == p1]["theme"].value_counts().head(2)
        out.append(
            f"{p1.title()} is named on {int(ps.iloc[0].Count)} back charges "
            f"({ps.iloc[0].Count / len(known):.0%} of those with a responsible person; "
            f"{money(ps.iloc[0].Total_Cost)}), mostly " +
            " and ".join(f"{t.lower()} ({c})" for t, c in p1_top.items()) +
            f". The next highest is {ps.index[1].title()} with {int(ps.iloc[1].Count)}.")
        p1_ev = (known[known["person"] == p1]["recovery"] == "Insufficient evidence").sum()
        out.append(
            f"Of {p1.title()}'s {int(ps.iloc[0].Count)} back charges, {p1_ev} are marked 'not enough evidence "
            f"to determine fault', so the person is named but the cost is not being recovered. The heat map "
            f"shows which mistakes each person repeats, which is the starting point for targeted coaching.")

    blank = (df["person"] == "(NOT RECORDED)").sum()
    ev = df[df["recovery"] == "Insufficient evidence"]
    out.append(
        f"Accountability gaps: {blank} back charges ({blank / n:.0%}) have no responsible person, and "
        f"{len(ev)} ({money(ev['cost'].sum())}) were not passed on because there was not enough "
        f"evidence to determine fault. Requiring photos and the installer name at the time of the "
        f"report would make more of these recoverable.")

    pay = summarize(df, "payee").sort_values("Count", ascending=False)
    out.append(
        f"By builder, {pay.index[0].title()} ({int(pay.iloc[0].Count)}, {money(pay.iloc[0].Total_Cost)}) "
        f"and {pay.index[1].title()} ({int(pay.iloc[1].Count)}, {money(pay.iloc[1].Total_Cost)}) account "
        f"for {(pay.iloc[0].Count + pay.iloc[1].Count) / n:.0%} of all back charges.")

    multi = df[df["address"] != ""].groupby("address").size()
    multi = multi[multi >= 2]
    out.append(
        f"{len(multi)} job addresses have two or more back charges ({int(multi.sum())} incidents), "
        f"which usually means one rough-in mistake triggered several trades' repair invoices.")
    return out


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def write_excel(df, path):
    out = df.copy()
    out["consequences"] = out["consequences"].map(", ".join)
    out = out.rename(columns={
        "id": "PO#/Backcharge #", "respondent": "Respondent", "address": "Job Address", "job": "Job Number",
        "incident_date": "Incident Date", "invoice_date": "Invoice Date", "claim_type": "Claim Type",
        "dept": "Responsible Department", "incident_type": "Incident Type",
        "property": "Type of Property Damaged", "desc": "Description of Incident",
        "phase": "Phase of Incident", "person": "Responsible Person", "cost": "Invoiced Cost",
        "payee": "Payee (cleaned)", "reason": "Non-Payment Reason", "theme": "Root-Cause Theme (from notes)",
        "consequences": "What Had To Be Fixed", "vague": "Vague Note?", "recovery": "Recovery Status",
        "month": "Incident Month", "days_to_invoice": "Days Incident->Invoice"})
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame({"Finding": findings(df)}).to_excel(xw, sheet_name="Findings", index=False)
        out.to_excel(xw, sheet_name="Registry (Categorized)", index=False)
        summarize(df, "theme").to_excel(xw, sheet_name="By Theme")
        summarize(df, "person").to_excel(xw, sheet_name="By Responsible Person")
        pd.crosstab(df["person"], df["theme"], margins=True, margins_name="Total") \
            .sort_values("Total", ascending=False).to_excel(xw, sheet_name="Person x Theme")
        summarize(df, "payee").to_excel(xw, sheet_name="By Builder")
        summarize(df, "dept").to_excel(xw, sheet_name="By Department")
        summarize(df, "phase").to_excel(xw, sheet_name="By Phase")
        summarize(df, "incident_type").to_excel(xw, sheet_name="By Incident Type")
        summarize(df, "recovery").to_excel(xw, sheet_name="Recovery Status")
        summarize(df[df["month"] != ""], "month").sort_index().to_excel(xw, sheet_name="By Month")
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 70)


def write_html(df, path, source_name):
    rows = []
    for r in df.itertuples():
        rows.append({
            "id": r.id, "addr": r.address, "job": r.job,
            "date": r.incident_date.isoformat() if r.incident_date else "",
            "month": r.month, "claim": r.claim_type, "dept": r.dept, "itype": r.incident_type,
            "prop": r.property, "desc": r.desc, "phase": r.phase, "person": r.person,
            "cost": float(r.cost), "payee": r.payee, "reason": r.reason, "theme": r.theme,
            "cons": r.consequences, "vague": bool(r.vague), "rec": r.recovery,
            "resp": r.respondent,
            "lag": None if r.days_to_invoice is None or pd.isna(r.days_to_invoice) else int(r.days_to_invoice),
        })
    payload = {
        "rows": rows,
        "themes": [{"name": n, "desc": d} for n, d, _ in THEMES],
        "findings": findings(df),
        "leakThemes": LEAK_THEMES,
        "source": source_name,
        "generated": date.today().strftime("%B %d, %Y").replace(" 0", " "),
    }
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    page = HTML_TEMPLATE.replace("__DATA__", data_json).replace("__SOURCE__", html.escape(source_name))
    Path(path).write_text(page, encoding="utf-8")


HTML_TEMPLATE = r"""<title>Back Charge Root Causes</title>
<meta name="claude:viewport-fit-sides" content="cover">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Layout: field-report sheet - a filter strip, a findings brief, then chart panels on a 12-col grid */
:root {
  --bg: #f4f5f7; --panel: #ffffff; --ink: #14171c; --ink-2: #4a5160; --ink-3: #7a8190;
  --rule: #dde0e6; --grid: #eceef2; --accent: #2a78d6; --accent-soft: #e3eefb;
  --warn: #b5530f; --chip: #eef0f4;
  --seq-0: #f3f7fd; --seq-1: #cde2fb; --seq-2: #9ec5f4; --seq-3: #6da7ec; --seq-4: #3987e5; --seq-5: #256abf; --seq-6: #184f95;
  --display: "Barlow Condensed", "Arial Narrow", "Roboto Condensed", sans-serif;
  --body: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif;
  --mono: "IBM Plex Mono", ui-monospace, Consolas, monospace;
}
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
  --bg: #121418; --panel: #1b1e24; --ink: #eef0f3; --ink-2: #b7bdc8; --ink-3: #8a91a0;
  --rule: #2c313a; --grid: #252a32; --accent: #3987e5; --accent-soft: #1d2d44; --warn: #f0954f; --chip: #252a32;
  --seq-0: #1d2330; --seq-1: #1a3354; --seq-2: #184f95; --seq-3: #256abf; --seq-4: #3987e5; --seq-5: #6da7ec; --seq-6: #b7d3f6;
  color-scheme: dark; } }
:root[data-theme="dark"] {
  --bg: #121418; --panel: #1b1e24; --ink: #eef0f3; --ink-2: #b7bdc8; --ink-3: #8a91a0;
  --rule: #2c313a; --grid: #252a32; --accent: #3987e5; --accent-soft: #1d2d44; --warn: #f0954f; --chip: #252a32;
  --seq-0: #1d2330; --seq-1: #1a3354; --seq-2: #184f95; --seq-3: #256abf; --seq-4: #3987e5; --seq-5: #6da7ec; --seq-6: #b7d3f6;
  color-scheme: dark; }
* { box-sizing: border-box; }
body { background: var(--bg); color: var(--ink); font: 14px/1.5 var(--body); margin: 0; }
.wrap { max-width: 1280px; margin: 0 auto; padding-block: 28px 56px;
  padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px));
  display: grid; grid-template-columns: minmax(0, 1fr); gap: 20px; }
header { display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between; gap: 12px;
  border-bottom: 3px solid var(--ink); padding-bottom: 12px; }
h1 { font: 700 clamp(30px, 5vw, 44px)/1 var(--display); letter-spacing: .01em; margin: 0; text-transform: uppercase; text-wrap: balance; }
.meta { font: 12px/1.4 var(--mono); color: var(--ink-3); text-align: right; }
h2 { font: 600 20px/1.2 var(--display); text-transform: uppercase; letter-spacing: .04em; margin: 0; }
.sub { color: var(--ink-3); font-size: 12.5px; margin: 2px 0 0; }
.filters { display: flex; flex-wrap: wrap; gap: 10px 14px; align-items: end; }
.filters label { display: grid; gap: 3px; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); }
select, input[type=search] { font: 13px var(--body); color: var(--ink); background: var(--panel); border: 1px solid var(--rule);
  border-radius: 4px; padding: 6px 8px; min-width: 0; max-width: 220px; }
select:focus-visible, input:focus-visible, button:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
button.reset { font: 600 12px var(--body); color: var(--accent); background: none; border: 1px solid var(--rule); border-radius: 4px; padding: 7px 10px; cursor: pointer; }
.filter-note { font-size: 12px; color: var(--ink-3); flex-basis: 100%; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 1px; background: var(--rule); border: 1px solid var(--rule); }
.kpi { background: var(--panel); padding: 14px 16px; display: grid; gap: 2px; align-content: start; }
.kpi .v { font: 600 28px/1.1 var(--mono); font-variant-numeric: tabular-nums; }
.kpi .l { font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); }
.kpi .n { font-size: 12px; color: var(--ink-2); }
.kpi.flag .v { color: var(--warn); }
.grid { display: grid; grid-template-columns: repeat(12, minmax(0, 1fr)); gap: 20px; }
.panel { background: var(--panel); border: 1px solid var(--rule); padding: 16px 18px; min-width: 0; display: grid; gap: 12px; align-content: start; }
.s12 { grid-column: span 12; } .s7 { grid-column: span 7; } .s5 { grid-column: span 5; } .s6 { grid-column: span 6; }
@media (max-width: 900px) { .s7, .s5, .s6 { grid-column: span 12; } }
.phead { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: start; gap: 8px; }
.seg { display: inline-flex; border: 1px solid var(--rule); border-radius: 4px; overflow: hidden; }
.seg button { font: 500 12px var(--body); background: var(--panel); color: var(--ink-2); border: 0; padding: 5px 10px; cursor: pointer; }
.seg button[aria-pressed=true] { background: var(--ink); color: var(--panel); }
.findings ol { margin: 0; padding-left: 20px; display: grid; gap: 8px; max-width: 110ch; }
.findings li { color: var(--ink-2); } .findings li::marker { font: 600 13px var(--mono); color: var(--ink); }
.chart { width: 100%; overflow-x: auto; }
.chart svg { display: block; }
.chart text { fill: var(--ink-2); font: 12px var(--body); }
.chart .val { fill: var(--ink); font: 500 11.5px var(--mono); }
.chart .axis { fill: var(--ink-3); font: 11px var(--mono); }
.bar { fill: var(--accent); } .bar:hover, .bar.hot { fill: var(--seq-5); cursor: pointer; }
.gl { stroke: var(--grid); stroke-width: 1; }
.tip { position: fixed; z-index: 20; pointer-events: none; background: var(--ink); color: var(--panel); font-size: 12px;
  padding: 8px 10px; border-radius: 4px; max-width: 320px; box-shadow: 0 4px 14px rgba(0,0,0,.18); }
.tip b { font: 600 13px var(--mono); display: block; }
table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
th, td { text-align: left; padding: 7px 8px; border-bottom: 1px solid var(--grid); vertical-align: top; }
th { font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-3); font-weight: 600; position: sticky; top: 0; background: var(--panel); }
td.num, th.num { text-align: right; font-family: var(--mono); font-variant-numeric: tabular-nums; white-space: nowrap; }
.tablebox { overflow: auto; max-height: 520px; border: 1px solid var(--rule); }
.chip { display: inline-block; background: var(--chip); color: var(--ink-2); border-radius: 3px; padding: 1px 6px; font-size: 11px; margin: 1px 2px 1px 0; white-space: nowrap; }
.chip.warn { color: var(--warn); }
.theme-desc { color: var(--ink-3); font-size: 11.5px; }
.ex { color: var(--ink-2); font-size: 12px; }
.heat td.c { text-align: center; font: 500 11.5px var(--mono); min-width: 34px; }
.heat th.rot { writing-mode: vertical-rl; transform: rotate(180deg); height: 150px; white-space: nowrap; text-transform: none; font-size: 11.5px; letter-spacing: 0; vertical-align: bottom; }
.legend-scale { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--ink-3); }
.legend-scale span.sw { width: 18px; height: 10px; display: inline-block; }
.empty { color: var(--ink-3); font-style: italic; padding: 12px 0; }
.count-line { font: 12px var(--mono); color: var(--ink-3); }
@media (prefers-reduced-motion: no-preference) { .bar { transition: fill .12s; } }
</style>

<div class="wrap">
  <header>
    <div>
      <h1>Back Charge Root Causes</h1>
      <p class="sub">What the incident notes say is going wrong, who is involved, and what it costs.</p>
    </div>
    <div class="meta" id="meta"></div>
  </header>

  <section class="panel" aria-label="Filters">
    <div class="filters">
      <label>Builder<select id="f-payee"></select></label>
      <label>Responsible dept<select id="f-dept"></select></label>
      <label>Responsible person<select id="f-person"></select></label>
      <label>Theme<select id="f-theme"></select></label>
      <label>Claim type<select id="f-claim"></select></label>
      <label>From<select id="f-from"></select></label>
      <label>To<select id="f-to"></select></label>
      <button class="reset" id="f-reset" type="button">Clear filters</button>
      <div class="filter-note" id="f-note"></div>
    </div>
  </section>

  <section class="kpis" id="kpis" aria-label="Key numbers"></section>

  <section class="panel findings">
    <div><h2>Findings</h2><p class="sub">Calculated from the full registry each time the script runs (not affected by the filters).</p></div>
    <ol id="findings"></ol>
  </section>

  <div class="grid">
    <section class="panel s7">
      <div class="phead">
        <div><h2>Root causes from the notes</h2><p class="sub">Each "Description of Incident" sorted into one theme by keyword rules.</p></div>
        <div class="seg" data-target="theme"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div>
      </div>
      <div class="chart" id="c-theme"></div>
    </section>
    <section class="panel s5">
      <div class="phead"><div><h2>What had to be fixed</h2><p class="sub">Damage named in the notes. One incident can hit several.</p></div>
        <div class="seg" data-target="cons"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-cons"></div>
    </section>

    <section class="panel s12">
      <div><h2>Theme detail</h2><p class="sub">Average cost shows which problems are expensive when they happen. Click a row to filter the dashboard to that theme.</p></div>
      <div class="tablebox" style="max-height:none"><table id="t-theme"></table></div>
    </section>

    <section class="panel s12">
      <div class="phead"><div><h2>Back charges by month of incident</h2><p class="sub">Uses Backcharge Incident Date.</p></div>
        <div class="seg" data-target="month"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-month"></div>
    </section>

    <section class="panel s6">
      <div class="phead"><div><h2>Responsible person</h2><p class="sub">Top 15. Click a bar to filter. "(Not recorded)" shown for scale.</p></div>
        <div class="seg" data-target="person"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-person"></div>
    </section>
    <section class="panel s6">
      <div class="phead"><div><h2>Builder (payee)</h2><p class="sub">Builder names cleaned so case variants count together. Click to filter.</p></div>
        <div class="seg" data-target="payee"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-payee"></div>
    </section>

    <section class="panel s12">
      <div class="phead"><div><h2>Who is repeating which mistake</h2><p class="sub">Responsible person (top 12 named) by root-cause theme, count of back charges.</p></div>
        <div class="legend-scale" id="heat-legend"></div></div>
      <div class="tablebox" style="max-height:none"><table class="heat" id="t-heat"></table></div>
    </section>

    <section class="panel s6">
      <div class="phead"><div><h2>Responsible department</h2><p class="sub">Spelling variants combined.</p></div>
        <div class="seg" data-target="dept"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-dept"></div>
    </section>
    <section class="panel s6">
      <div class="phead"><div><h2>Phase of incident</h2><p class="sub">Dates typed into this column are treated as not recorded.</p></div>
        <div class="seg" data-target="phase"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-phase"></div>
    </section>

    <section class="panel s6">
      <div class="phead"><div><h2>Was the cost passed on?</h2><p class="sub">Grouped from "Non-Payment Reason".</p></div>
        <div class="seg" data-target="rec"><button type="button" aria-pressed="true" data-m="count">Count</button><button type="button" aria-pressed="false" data-m="cost">Cost</button></div></div>
      <div class="chart" id="c-rec"></div>
    </section>
    <section class="panel s6">
      <div><h2>Record quality</h2><p class="sub">Blank or weak fields make back charges hard to recover.</p></div>
      <div class="chart" id="c-quality"></div>
    </section>

    <section class="panel s6">
      <div><h2>Jobs with repeat back charges</h2><p class="sub">Addresses with 2 or more incidents.</p></div>
      <div class="tablebox"><table id="t-repeat"></table></div>
    </section>
    <section class="panel s6">
      <div><h2>Most expensive incidents</h2><p class="sub">Top 12 by invoiced cost.</p></div>
      <div class="tablebox"><table id="t-top"></table></div>
    </section>

    <section class="panel s12">
      <div class="phead"><div><h2>All incidents</h2><p class="count-line" id="all-count"></p></div>
        <label class="filters"><span class="sr" style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--ink-3)">Search notes</span><input type="search" id="q" placeholder="e.g. glue, stack, kinked"></label></div>
      <div class="tablebox"><table id="t-all"></table></div>
    </section>
  </div>
</div>
<div class="tip" id="tip" hidden></div>

<script>
const D = __DATA__;
const $ = s => document.querySelector(s);
const fmt$ = v => (v < 0 ? "-$" : "$") + Math.abs(Math.round(v)).toLocaleString();
const fmtN = v => v.toLocaleString();
const pct = (a, b) => b ? Math.round(100 * a / b) + "%" : "0%";
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const title = s => s.startsWith("(") ? s.charAt(0) + s.slice(1).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()) : s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).replace(/'S\b/g, "'s");
const mode = { theme: "count", cons: "count", month: "count", person: "count", payee: "count", dept: "count", phase: "count", rec: "count" };
const F = { payee: "", dept: "", person: "", theme: "", claim: "", from: "", to: "" };
const months = [...new Set(D.rows.map(r => r.month).filter(Boolean))].sort();
const themeOrder = D.themes.map(t => t.name);

$("#meta").innerHTML = `Source: ${esc(D.source)}<br>Sheet: Registry · Generated ${esc(D.generated)}`;
D.findings.forEach(f => { const li = document.createElement("li"); li.textContent = f; $("#findings").appendChild(li); });

function opts(id, values, label, labeler = title) {
  const el = $(id); el.innerHTML = "";
  el.append(new Option(label, ""));
  values.forEach(v => el.append(new Option(labeler(v), v)));
}
function byCount(key) {
  const m = {}; D.rows.forEach(r => m[r[key]] = (m[r[key]] || 0) + 1);
  return Object.keys(m).sort((a, b) => m[b] - m[a]);
}
opts("#f-payee", byCount("payee"), "All builders");
opts("#f-dept", byCount("dept"), "All departments");
opts("#f-person", byCount("person"), "All people");
opts("#f-theme", themeOrder, "All themes", s => s);
opts("#f-claim", byCount("claim"), "All claim types", s => s);
const mLabel = m => new Date(m + "-15").toLocaleString("en-US", { month: "short", year: "numeric" });
opts("#f-from", months, "Earliest", mLabel);
opts("#f-to", months, "Latest", mLabel);
Object.keys(F).forEach(k => $("#f-" + k).addEventListener("change", e => { F[k] = e.target.value; render(); }));
$("#f-reset").addEventListener("click", () => { Object.keys(F).forEach(k => { F[k] = ""; $("#f-" + k).value = ""; }); $("#q").value = ""; render(); });
function setFilter(k, v) { F[k] = F[k] === v ? "" : v; $("#f-" + k).value = F[k]; render(); window.scrollTo({ top: 0, behavior: "smooth" }); }

function filtered() {
  return D.rows.filter(r => (!F.payee || r.payee === F.payee) && (!F.dept || r.dept === F.dept) &&
    (!F.person || r.person === F.person) && (!F.theme || r.theme === F.theme) && (!F.claim || r.claim === F.claim) &&
    (!F.from || (r.month && r.month >= F.from)) && (!F.to || (r.month && r.month <= F.to)));
}
function group(rows, key, multi = false) {
  const m = new Map();
  rows.forEach(r => (multi ? r[key] : [r[key]]).forEach(k => {
    const g = m.get(k) || { key: k, count: 0, cost: 0, rows: [] };
    g.count++; g.cost += r.cost; g.rows.push(r); m.set(k, g);
  }));
  return [...m.values()];
}

/* tooltip */
const tip = $("#tip");
function showTip(e, html) { tip.innerHTML = html; tip.hidden = false; moveTip(e); }
function moveTip(e) {
  const x = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 8);
  const y = Math.min(e.clientY + 14, window.innerHeight - tip.offsetHeight - 8);
  tip.style.left = x + "px"; tip.style.top = y + "px";
}
function hideTip() { tip.hidden = true; }

/* horizontal bar chart */
function hbar(el, groups, m, opt = {}) {
  const box = $(el);
  if (!groups.length) { box.innerHTML = '<p class="empty">No back charges match these filters.</p>'; return; }
  const data = groups.slice().sort((a, b) => b[m] - a[m]).slice(0, opt.limit || 99);
  const W = Math.max(box.clientWidth || 600, 300), rowH = 26, labelW = Math.min(opt.labelW || 230, W * .45), valW = 92;
  const H = data.length * rowH + 8, plotW = W - labelW - valW;
  const max = Math.max(...data.map(d => d[m]), 1);
  const tot = groups.reduce((s, g) => s + g[m], 0);
  let s = `<svg width="${W}" height="${H}" role="img" aria-label="${esc(opt.aria || "")}">`;
  data.forEach((d, i) => {
    const y = i * rowH + 4, w = Math.max(2, plotW * Math.max(d[m], 0) / max);
    const lbl = (opt.labeler || (x => x))(d.key);
    const short = lbl.length > 34 ? lbl.slice(0, 33) + "…" : lbl;
    const val = m === "cost" ? fmt$(d.cost) : fmtN(d.count);
    s += `<text x="${labelW - 8}" y="${y + 13}" text-anchor="end" dominant-baseline="middle">${esc(short)}</text>`;
    s += `<rect class="bar" data-i="${i}" x="${labelW}" y="${y + 3}" width="${w}" height="${rowH - 8}" rx="3"/>`;
    s += `<text class="val" x="${labelW + w + 6}" y="${y + 13}" dominant-baseline="middle">${val}</text>`;
  });
  s += "</svg>";
  box.innerHTML = s;
  box.querySelectorAll("rect.bar").forEach(rect => {
    const d = data[+rect.dataset.i];
    const lbl = (opt.labeler || (x => x))(d.key);
    const h = `<b>${m === "cost" ? fmt$(d.cost) : fmtN(d.count)}</b>${esc(lbl)}<br>${fmtN(d.count)} back charges · ${fmt$(d.cost)} · avg ${fmt$(d.cost / d.count)}<br>${pct(d[m], tot)} of ${m === "cost" ? "cost" : "count"} shown`;
    rect.addEventListener("pointermove", e => showTip(e, h));
    rect.addEventListener("pointerleave", hideTip);
    if (opt.filterKey) rect.addEventListener("click", () => setFilter(opt.filterKey, d.key));
  });
}

/* monthly columns */
function columns(el, rows, m) {
  const box = $(el);
  const g = new Map(group(rows.filter(r => r.month), "month").map(x => [x.key, x]));
  const keys = months.filter(k => (!F.from || k >= F.from) && (!F.to || k <= F.to));
  if (!keys.length) { box.innerHTML = '<p class="empty">No dated back charges match these filters.</p>'; return; }
  const W = Math.max(box.clientWidth || 800, 300), H = 230, padL = 56, padB = 34, padT = 18;
  const vals = keys.map(k => (g.get(k) || { count: 0, cost: 0 })[m]);
  const rawMax = Math.max(...vals, 1);
  const step = niceStep(rawMax / 4), max = Math.ceil(rawMax / step) * step;
  const bw = (W - padL - 8) / keys.length;
  let s = `<svg width="${W}" height="${H}" role="img" aria-label="Back charges by month">`;
  for (let v = 0; v <= max + 1e-9; v += step) {
    const y = padT + (H - padT - padB) * (1 - v / max);
    s += `<line class="gl" x1="${padL}" x2="${W - 4}" y1="${y}" y2="${y}"/><text class="axis" x="${padL - 6}" y="${y}" text-anchor="end" dominant-baseline="middle">${m === "cost" ? fmt$(v) : v}</text>`;
  }
  keys.forEach((k, i) => {
    const v = vals[i], h = (H - padT - padB) * Math.max(v, 0) / max, x = padL + i * bw + 3, w = Math.max(bw - 6, 2);
    s += `<rect class="bar" data-k="${k}" x="${x}" y="${H - padB - h}" width="${w}" height="${h}" rx="3"/>`;
    if (keys.length <= 18 || i % 2 === 0) s += `<text class="axis" x="${x + w / 2}" y="${H - padB + 16}" text-anchor="middle">${mLabel(k).replace(" 20", " '")}</text>`;
  });
  s += "</svg>";
  box.innerHTML = s;
  box.querySelectorAll("rect.bar").forEach(rect => {
    const d = g.get(rect.dataset.k) || { count: 0, cost: 0 };
    const top = d.rows ? group(d.rows, "theme").sort((a, b) => b.count - a.count).slice(0, 3).map(t => `${esc(t.key)} (${t.count})`).join("<br>") : "";
    const h = `<b>${m === "cost" ? fmt$(d.cost) : fmtN(d.count)}</b>${mLabel(rect.dataset.k)}<br>${fmtN(d.count)} back charges · ${fmt$(d.cost)}${top ? "<br><br>Top themes:<br>" + top : ""}`;
    rect.addEventListener("pointermove", e => showTip(e, h));
    rect.addEventListener("pointerleave", hideTip);
  });
}
function niceStep(x) { const p = Math.pow(10, Math.floor(Math.log10(x || 1))); const n = x / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; }

function kpis(rows) {
  const n = rows.length, cost = rows.reduce((s, r) => s + r.cost, 0);
  const costs = rows.map(r => r.cost).sort((a, b) => a - b);
  const med = n ? costs[Math.floor((n - 1) / 2)] : 0;
  const noPerson = rows.filter(r => r.person === "(NOT RECORDED)").length;
  const ev = rows.filter(r => r.rec === "Insufficient evidence");
  const leakThemes = D.leakThemes;
  const leak = rows.filter(r => leakThemes.includes(r.theme)).reduce((s, r) => s + r.cost, 0);
  const lags = rows.map(r => r.lag).filter(v => v !== null).sort((a, b) => a - b);
  const lagMed = lags.length ? lags[Math.floor((lags.length - 1) / 2)] : null;
  const k = [
    ["Back charges", fmtN(n), `${new Set(rows.map(r => r.addr)).size} job addresses`],
    ["Total invoiced", fmt$(cost), `median ${fmt$(med)} · avg ${fmt$(n ? cost / n : 0)}`],
    ["Water-leak share of cost", pct(leak, cost), `${fmt$(leak)} from leaks`],
    ["No responsible person", pct(noPerson, n), `${noPerson} records blank`, true],
    ["Not enough evidence", fmt$(ev.reduce((s, r) => s + r.cost, 0)), `${ev.length} back charges not passed on`, true],
    ["Median days to invoice", lagMed === null ? "–" : lagMed, "incident date to invoice date"],
  ];
  $("#kpis").innerHTML = k.map(([l, v, nn, flag]) => `<div class="kpi${flag ? " flag" : ""}"><span class="l">${l}</span><span class="v">${v}</span><span class="n">${nn}</span></div>`).join("");
}

function themeTable(rows) {
  const g = new Map(group(rows, "theme").map(x => [x.key, x]));
  const tot = rows.length, totC = rows.reduce((s, r) => s + r.cost, 0);
  const list = D.themes.filter(t => g.has(t.name)).map(t => ({ ...t, ...g.get(t.name) })).sort((a, b) => b.count - a.count);
  let h = `<thead><tr><th>Theme</th><th class="num">Count</th><th class="num">% count</th><th class="num">Cost</th><th class="num">% cost</th><th class="num">Avg</th><th>Top responsible people</th><th>Example notes</th></tr></thead><tbody>`;
  list.forEach(t => {
    const ppl = group(t.rows.filter(r => r.person !== "(NOT RECORDED)"), "person").sort((a, b) => b.count - a.count).slice(0, 3);
    const ex = [...new Set(t.rows.map(r => r.desc))].sort((a, b) => a.length - b.length).filter(d => d.length > 18).slice(0, 2);
    h += `<tr data-theme="${esc(t.name)}" style="cursor:pointer"><td><b>${esc(t.name)}</b><div class="theme-desc">${esc(t.desc)}</div></td>
      <td class="num">${t.count}</td><td class="num">${pct(t.count, tot)}</td><td class="num">${fmt$(t.cost)}</td><td class="num">${pct(t.cost, totC)}</td><td class="num">${fmt$(t.cost / t.count)}</td>
      <td>${ppl.map(p => `<span class="chip">${esc(title(p.key))} · ${p.count}</span>`).join("") || '<span class="theme-desc">none recorded</span>'}</td>
      <td class="ex">${ex.map(e => "“" + esc(e.length > 120 ? e.slice(0, 118) + "…" : e) + "”").join("<br>")}</td></tr>`;
  });
  $("#t-theme").innerHTML = h + "</tbody>";
  $("#t-theme").querySelectorAll("tr[data-theme]").forEach(tr => tr.addEventListener("click", () => setFilter("theme", tr.dataset.theme)));
}

function heat(rows) {
  const named = rows.filter(r => r.person !== "(NOT RECORDED)");
  const people = group(named, "person").sort((a, b) => b.count - a.count).slice(0, 12);
  const themes = group(named, "theme").sort((a, b) => b.count - a.count).map(t => t.key);
  if (!people.length) { $("#t-heat").innerHTML = '<tbody><tr><td class="empty">No named responsible people in this selection.</td></tr></tbody>'; $("#heat-legend").innerHTML = ""; return; }
  let max = 1; people.forEach(p => themes.forEach(t => { max = Math.max(max, p.rows.filter(r => r.theme === t).length); }));
  const shade = v => v === 0 ? "transparent" : `var(--seq-${Math.min(6, 1 + Math.floor(5 * (v - 1) / Math.max(max - 1, 1)))})`;
  const ink = v => (v === 0 ? "var(--ink-3)" : (Math.min(6, 1 + Math.floor(5 * (v - 1) / Math.max(max - 1, 1))) >= 4 ? "#fff" : "var(--ink)"));
  let h = `<thead><tr><th>Person</th><th class="num">Total</th>${themes.map(t => `<th class="rot">${esc(t)}</th>`).join("")}</tr></thead><tbody>`;
  people.forEach(p => {
    h += `<tr><td><a href="#" data-p="${esc(p.key)}" style="color:var(--ink);text-decoration:none;font-weight:500">${esc(title(p.key))}</a></td><td class="num">${p.count}</td>`;
    themes.forEach(t => { const v = p.rows.filter(r => r.theme === t).length;
      h += `<td class="c" style="background:${shade(v)};color:${ink(v)}" title="${esc(title(p.key))} · ${esc(t)}: ${v}">${v || "·"}</td>`; });
    h += "</tr>";
  });
  $("#t-heat").innerHTML = h + "</tbody>";
  $("#t-heat").querySelectorAll("a[data-p]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); setFilter("person", a.dataset.p); }));
  $("#heat-legend").innerHTML = `fewer <span class="sw" style="background:var(--seq-1)"></span><span class="sw" style="background:var(--seq-3)"></span><span class="sw" style="background:var(--seq-5)"></span><span class="sw" style="background:var(--seq-6)"></span> more (max ${max})`;
}

function quality(rows) {
  const n = rows.length || 1;
  const q = [
    { key: "No responsible person", count: rows.filter(r => r.person === "(NOT RECORDED)").length },
    { key: "No type of property damaged", count: rows.filter(r => r.prop === "(Not recorded)").length },
    { key: "No responsible department", count: rows.filter(r => r.dept === "(NOT RECORDED)").length },
    { key: "No phase (or a date typed in)", count: rows.filter(r => r.phase === "(NOT RECORDED)").length },
    { key: "Vague note (undefined, no photo, ≤3 words)", count: rows.filter(r => r.vague).length },
    { key: "No incident date", count: rows.filter(r => !r.date).length },
    { key: "No PO / back charge #", count: rows.filter(r => !r.id).length },
  ].map(x => ({ ...x, cost: 0 }));
  hbar("#c-quality", q, "count", { labelW: 270, aria: "Missing or weak fields" });
  $("#c-quality").querySelectorAll("rect.bar").forEach((rect, i) => {
    const d = q.slice().sort((a, b) => b.count - a.count)[i];
    rect.addEventListener("pointermove", e => showTip(e, `<b>${d.count} of ${rows.length}</b>${esc(d.key)} (${pct(d.count, n)})`));
  });
}

function repeatTable(rows) {
  const g = group(rows.filter(r => r.addr), "addr").filter(x => x.count >= 2).sort((a, b) => b.count - a.count || b.cost - a.cost);
  if (!g.length) { $("#t-repeat").innerHTML = '<tbody><tr><td class="empty">No repeat addresses in this selection.</td></tr></tbody>'; return; }
  $("#t-repeat").innerHTML = `<thead><tr><th>Address</th><th class="num">BCs</th><th class="num">Cost</th><th>Themes</th></tr></thead><tbody>` +
    g.map(x => `<tr><td>${esc(x.key)}<div class="theme-desc">${esc(title(x.rows[0].payee))}</div></td><td class="num">${x.count}</td><td class="num">${fmt$(x.cost)}</td><td>${[...new Set(x.rows.map(r => r.theme))].map(t => `<span class="chip">${esc(t)}</span>`).join("")}</td></tr>`).join("") + "</tbody>";
}
function topTable(rows) {
  const t = rows.slice().sort((a, b) => b.cost - a.cost).slice(0, 12);
  $("#t-top").innerHTML = `<thead><tr><th class="num">Cost</th><th>Note</th><th>Theme</th></tr></thead><tbody>` +
    t.map(r => `<tr><td class="num">${fmt$(r.cost)}</td><td>${esc(r.desc)}<div class="theme-desc">${esc(r.addr)} · ${esc(title(r.person))}</div></td><td><span class="chip">${esc(r.theme)}</span></td></tr>`).join("") + "</tbody>";
}
function allTable(rows) {
  const q = $("#q").value.trim().toLowerCase();
  const list = rows.filter(r => !q || (r.desc + " " + r.addr + " " + r.person + " " + r.itype).toLowerCase().includes(q))
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  $("#all-count").textContent = `${list.length} of ${D.rows.length} back charges`;
  $("#t-all").innerHTML = `<thead><tr><th>Date</th><th>PO / BC #</th><th>Address</th><th>Description of incident</th><th>Theme</th><th>Person</th><th>Builder</th><th class="num">Cost</th><th>Passed on?</th></tr></thead><tbody>` +
    list.map(r => `<tr><td class="num">${esc(r.date)}</td><td>${esc(r.id)}</td><td>${esc(r.addr)}</td><td style="min-width:260px">${esc(r.desc)}${r.vague ? ' <span class="chip warn">vague</span>' : ""}</td><td><span class="chip">${esc(r.theme)}</span></td><td>${esc(title(r.person))}</td><td>${esc(title(r.payee))}</td><td class="num">${fmt$(r.cost)}</td><td>${esc(r.rec)}</td></tr>`).join("") + "</tbody>";
}
$("#q").addEventListener("input", () => allTable(filtered()));

document.querySelectorAll(".seg").forEach(seg => seg.addEventListener("click", e => {
  const b = e.target.closest("button"); if (!b) return;
  mode[seg.dataset.target] = b.dataset.m;
  seg.querySelectorAll("button").forEach(x => x.setAttribute("aria-pressed", x === b));
  render();
}));

function render() {
  hideTip();
  const rows = filtered();
  const active = Object.entries(F).filter(([, v]) => v);
  $("#f-note").textContent = active.length ? `Showing ${rows.length} of ${D.rows.length} back charges.` : `Showing all ${D.rows.length} back charges. Click bars or table rows to filter.`;
  kpis(rows);
  hbar("#c-theme", group(rows, "theme"), mode.theme, { labelW: 270, filterKey: "theme", aria: "Root-cause themes" });
  hbar("#c-cons", group(rows, "cons", true), mode.cons, { labelW: 200, aria: "Damage fixed" });
  themeTable(rows);
  columns("#c-month", rows, mode.month);
  hbar("#c-person", group(rows, "person"), mode.person, { limit: 15, labeler: title, filterKey: "person", aria: "Responsible person" });
  hbar("#c-payee", group(rows, "payee"), mode.payee, { limit: 15, labeler: title, filterKey: "payee", aria: "Builder" });
  heat(rows);
  hbar("#c-dept", group(rows, "dept"), mode.dept, { limit: 14, labeler: title, filterKey: "dept", aria: "Department" });
  hbar("#c-phase", group(rows, "phase"), mode.phase, { labeler: title, aria: "Phase" });
  hbar("#c-rec", group(rows, "rec"), mode.rec, { labelW: 240, aria: "Recovery" });
  quality(rows);
  repeatTable(rows); topTable(rows); allTable(rows);
}
let rt; window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(render, 150); });
render();
</script>
"""


def main():
    ap = argparse.ArgumentParser(description="Build the back charge dashboard from the Registry sheet.")
    ap.add_argument("workbook", nargs="?", default=DEFAULT_WORKBOOK,
                    help="Path to the back charge tracking workbook (.xlsb or .xlsx). "
                         "Defaults to DEFAULT_WORKBOOK at the top of this script.")
    ap.add_argument("--out", default=".", help="Folder for the outputs (default: current folder)")
    args = ap.parse_args()

    if not Path(args.workbook).exists():
        raise SystemExit(f"Workbook not found: {args.workbook}\n"
                         "Check that the P: drive is connected, or pass the workbook path as an argument.")
    df = load_registry(args.workbook)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    html_path = out / "back_charge_dashboard.html"
    xlsx_path = out / "back_charge_analysis.xlsx"
    write_html(df, html_path, Path(args.workbook).name)
    write_excel(df, xlsx_path)

    print(f"Rows analysed: {len(df)}   Total invoiced: {money(df['cost'].sum())}")
    print(summarize(df, "theme")[["Count", "Total_Cost", "Avg_Cost"]].to_string())
    print(f"\nWrote {html_path}\nWrote {xlsx_path}")


if __name__ == "__main__":
    main()
