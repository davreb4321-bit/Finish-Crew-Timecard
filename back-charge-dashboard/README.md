# Back Charge Dashboard

`backcharge_dashboard.py` reads the **Registry** sheet of the back charge tracking workbook,
sorts each *Description of Incident* (column K) into a root-cause theme, and writes:

- `back_charge_dashboard.html`: an interactive dashboard you open in a browser (filters, hover details, searchable incident list)
- `back_charge_analysis.xlsx`: the categorized registry plus summary tabs (findings, by theme, person, person × theme, builder, department, phase, incident type, recovery status, month)

## Run it

```bash
pip install pandas pyxlsb openpyxl
```

```bash
python backcharge_dashboard.py "2026_BACK_CHARGE_TRACKING.xlsb" --out reports
```

## Tuning the themes

The `THEMES` list near the top of the script holds the keyword rules. A note goes to the first
theme whose pattern matches. If a note lands in the wrong theme, check the
"Registry (Categorized)" tab, then add a keyword to the right theme or move that theme higher in the list.
