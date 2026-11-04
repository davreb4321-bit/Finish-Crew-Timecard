# Back Charge Import

Replaces `IMPORTBC2.PY`, the rearranging of its output, the XLOOKUP formulas in
Back Charge Data and the old Property Damage PDF macro.

## Monthly steps

1. Drop the accounting PDFs into the loading folder
   (`P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER`).
2. Double-click `run_import.bat`. It writes `BC_Import_Ready.xlsx` into the loading folder.
3. In the tracking workbook, click **Import Back Charges**. Rows land in Back Charge Data from row 10.
   - **Yellow** cells are suggestions (drop-downs, incident date, responsible person from an earlier back charge on the same job). Confirm or change them.
   - **Orange** cells could not be found. Fill them in.
   - The **Review** tab of `BC_Import_Ready.xlsx` lists back charges skipped as duplicates and the close calls between drop-down choices.
4. Click **Send to Registry** (unchanged).
5. Click **Create Property Damage PDFs**. Each PDF is saved to today's dated folder on P: and filed into its G: drive job folder,
   `G:\Mandi\Safety\Incident Investigations\2026 Incidents\Builder backcharges, 2026\<Lot> <Subdivision>, <Address>, <Deduction Date>, <Builder>`.
   The folder is created if it doesn't exist.

## One-time setup

```bash
pip install pandas pyxlsb openpyxl pdfplumber
```

Put `bc_import.py` and `run_import.bat` in the same folder.

In the tracking workbook:

1. Press Alt+F11, then File > Import File, and choose `BackChargeImport.bas`.
2. Add a button for `ImportBackCharges` on Back Charge Data.
3. Right-click the existing PDF button > Assign Macro > `CreatePropertyDamagePDFs`.
4. Save. Module4 (the old PDF macro) and the All Clear button are no longer needed.

The import overwrites columns A-W from row 10 down with values, so the XLOOKUP formulas there are replaced.
Drop-down lists stay in place.

## Settings

- `bc_import.py` top: loading folder, tracking workbook path, default respondent.
- `BackChargeImport.bas` top: loading folder and G: drive folder.
- `FolderBuilder` in `BackChargeImport.bas`: how each builder is named in the G: drive folder names.
