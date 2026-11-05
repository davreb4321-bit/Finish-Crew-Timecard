# Back Charge Dashboard

`backcharge_dashboard.py` reads the **Registry** sheet of the back charge tracking workbook,
sorts each *Description of Incident* (column K) into a root-cause theme, and writes:

- `back_charge_dashboard.html`: an interactive dashboard you open in a browser (filters, hover details, searchable incident list)
- `back_charge_analysis.xlsx`: the categorized registry plus summary tabs (findings, by theme, person, person × theme, builder, department, phase, incident type, recovery status, month)

## Run it

```bash
pip install pandas pyxlsb openpyxl
```

The script reads this workbook by default (set in `DEFAULT_WORKBOOK` at the top of the script):

`P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER\2026 BACK CHARGE TRACKING.xlsb`

On Windows, double-click `run_dashboard.bat`. It rebuilds both outputs next to the script and opens the dashboard.
Or from a command prompt:

```bash
python backcharge_dashboard.py
```

To read a different workbook or write somewhere else:

```bash
python backcharge_dashboard.py "D:\other\workbook.xlsb" --out "C:\Reports"
```

## Tuning the themes

The `THEMES` list near the top of the script holds the keyword rules. A note goes to the first
theme whose pattern matches. If a note lands in the wrong theme, check the
"Registry (Categorized)" tab, then add a keyword to the right theme or move that theme higher in the list.
