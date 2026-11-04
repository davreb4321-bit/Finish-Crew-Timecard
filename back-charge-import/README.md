# Back Charge Import

Replaces `IMPORTBC2.PY`, the rearranging of its output, the XLOOKUP formulas in
Back Charge Data and the old Property Damage PDF macro.

## Monthly steps

1. Drop the accounting PDFs into the loading folder
   (`P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER`).
2. Double-click `run_import.bat`. It writes `BC_Import_Ready.xlsx` into the loading folder.
   It also saves each back charge (accounting page + the builder's pages behind it) as its own PDF in
   `LOADING FOLDER\Accounting Sheets`, including back charges skipped because they're already in the Registry.
3. In the tracking workbook, click **Import Back Charges**. Rows land in Back Charge Data from row 10.
   - **Yellow** cells are suggestions (drop-downs, incident date, responsible person from an earlier back charge on the same job). Confirm or change them.
   - **Orange** cells could not be found. Fill them in.
   - The **Review** tab of `BC_Import_Ready.xlsx` lists back charges skipped as duplicates and the close calls between drop-down choices.
4. Click **Send to Registry** (unchanged).
5. Click **Create Property Damage PDFs**. Each PDF is saved to today's dated folder on P: and filed into its G: drive job folder,
   `G:\Mandi\Safety\Incident Investigations\2026 Incidents\Builder backcharges, 2026\<Lot> <Subdivision>, <Address>, <Deduction Date>, <Builder>`.
   If the house already has a folder (any date), the PDF goes into it; otherwise a new folder is created.
   Same house = same builder and the same street address (or the same lot and subdivision), using the glossary below.
   The same button then files every split accounting sheet into its house folder, even when Back Charge Data is empty.

## One-time setup

```bash
pip install pandas pyxlsb openpyxl pdfplumber pypdf
```

Put `bc_import.py` and `run_import.bat` in the same folder.

In the tracking workbook:

1. Press Alt+F11. Either use File > Import File and choose `BackChargeImport.bas`, or Insert > Module and paste the whole file in.
2. Add a button for `ImportBackCharges` on Back Charge Data.
3. Right-click the existing PDF button > Assign Macro > `CreatePropertyDamagePDFs`.
4. Save. Module4 (the old PDF macro) and the All Clear button are no longer needed.

The import overwrites columns A-W from row 10 down with values, so the XLOOKUP formulas there are replaced.
Drop-down lists stay in place.

## Glossary

Both files carry the same glossary. Keep them in step when you add to one.

- **Builders** (`BUILDERS` in `bc_import.py`, `BuilderGlossary` in `BackChargeImport.bas`): each builder lists the text that
  identifies it. A name matches when it *contains* that text, so "Fischer", "Fischer Homes" and "Fischer Home Management"
  are all Fischer. Capitals, punctuation and extra spaces are ignored ("M.I. Homes" = "MI HOMES").
- **Street words** (`STREET_WORDS` / `StreetGlossary`): Ln = Lane, St = Street, Dr = Drive, and so on. Addresses are compared
  by house number and street name, so "793 Rosebud", "793 Rosebud Ln" and "793 ROSEBUD LANE" are the same house.

The import script uses the glossary to name the builder when the job isn't in PRODUCTION DATA, and to find the job by
address when the PDF's job # is missing or wrong (marked yellow). The PDF macro uses it to find a house's existing
G: drive folder.

## Settings

- `bc_import.py` top: loading folder, tracking workbook path, default respondent.
- `BackChargeImport.bas` top: loading folder and G: drive folder.
- `FolderBuilder` in `BackChargeImport.bas`: how each builder is named in the G: drive folder names.
