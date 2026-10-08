# Finish Crew Timecard — Cover Sheet Prefill & Pay Math Build Plan

*Last updated 2026-10-08. Reflects the import flow, Send macro and app changes as built and
tested.*

Files in this repository:
- `docs/build-plan.md`: this plan
- `office-scripts/Import Finish Cover Sheet.ts`: Office Script used by the import flow
- `office-scripts/Populate Finish Crew Cover Sheet.ts`: Office Script for the daily Cover Sheet report
- `excel-macros/ClearCoverSheet.bas`: "Clear Sheet" button for the daily-entry workbook
- `excel-macros/SendToFinishCrewApp.bas`: "Send to Finish Crew App" button for the daily-entry
  workbook

Work through the phases in order. Each phase can be tested before starting the next.
Every formula below uses the control and column names from the current app export.

| Phase | What | Where |
|---|---|---|
| 1 | SCHEDULED JOBS list, SUBMISSIONS column changes | SharePoint |
| 2 | Import script | Excel for the web (Office Scripts) |
| 3 | Import flow | Power Automate |
| 3A | Office hand-off: fill in the sheet, click Send | Excel (desktop) |
| 3B | Daily Cover Sheet report script (one row per job) | Office Scripts |
| 4 | Restore pay splits and fix the pay math | Power Apps (scrJobInfo, scrPayItems, scrTimeCardDetail) |
| 5 | Builder as text | Power Apps and existing flows |
| 6 | Job picker, prefill, lead-only job list, navigation, delete scheduled job | Power Apps |
| 7 | Test | — |
| 8 | Field Supervisor Punch List (list view in Teams) | SharePoint |
| 9 | Scale (7,000+ jobs/year): server-side filters, indexes, delegation, optional cleanup | Power Automate, SharePoint, Power Apps |

Phase 4 fixes the current app, so do it (and test it) before the prefill work in Phase 6.
The "why the math is wrong" details are in the appendix.

---

## Phase 1 — SharePoint

### 1.1 New list: SCHEDULED JOBS

Create a blank list named **SCHEDULED JOBS** on the IndyWarrantyTracking-NRG365 site.

| # | Column | Type | Settings |
|---|---|---|---|
| 1 | Title | (built in) | Rename the display name to **Job Number** |
| 2 | JobDate | Date and time | Date only. **Index it** |
| 3 | Supervisor | Single line of text | |
| 4 | LeadInstaller | Single line of text | |
| 5 | HelperInstaller | Single line of text | |
| 6 | Builder | Single line of text | |
| 7 | Subdivision | Single line of text | |
| 8 | LotNumber | Single line of text | Text, not number (lots like "9B") |
| 9 | Phase | Single line of text | |
| 10 | JobStatus | Single line of text | Default value `Open`. **Index it** |
| 11 | SubmissionID | Number | 0 decimal places |
| 12 | ImportKey | Single line of text | **Index it** and turn on **Enforce unique values** |

Notes:
- Create each column with the name shown (no spaces) first, then rename the display name if you
  like. The internal name is what the flow and app use.
- JobStatus is text rather than Choice so the app's filter stays delegable (the Status.Value
  delegation warning noted in the project summary).
- ImportKey holds `JobNumber|yyyy-mm-dd`. The unique setting makes a re-saved cover sheet unable to
  create a duplicate job.
- Indexes: List settings → Indexed columns → Create a new index. Add them now, while the list is
  empty.

### 1.2 SUBMISSIONS: add two columns

| Column | Type | Why |
|---|---|---|
| BuilderName | Single line of text | Replaces the Builder lookup (Phase 5). Display name **Builder Name** |
| ScheduleJobID | Number, 0 decimals | Links a time card to its scheduled job |

The existing **Builder** column is a **Lookup** column, not a Choice column (the app export shows a
`Builder Id` field). SharePoint can't convert a lookup to text, so add the new text column and leave
the old one in place for existing records. Don't delete it; the app and flows fall back to it for
old time cards.

SUBMISSIONS also contains **Import Status**, **Source File**, **Import Key** and
**Assigned/Claimed by Email** columns that nothing in the app uses. They look like an earlier
import idea. Leave them alone; this plan doesn't use them.

### 1.3 Import folders

In the site's **Documents** library, inside **General**, create two folders side by side:
**Cover Sheet Import** (cover sheets arrive here) and **Cover Sheet Imported** (the flow keeps a
copy of every processed sheet here).

The Production office person needs **Edit** access to **Cover Sheet Import**.

---

## Phase 2 — Office Script

1. Open any workbook in Excel for the web → **Automate** → **New Script**.
2. Replace the contents with `office-scripts/Import Finish Cover Sheet.ts` from this repo. Open the
   file in Notepad (not Word), copy all, and paste into an **empty** editor. It should be about
   169 lines ending in a single `}`.
3. Name it **Import Finish Cover Sheet** and save it. It saves to the OneDrive of whoever owns the
   flow, so do this signed in as that account.

What it does:
- Reads the **FINISHCOVER** table by header name, so reordering columns doesn't break it.
- Reads the job date from **B4** and rejects invalid dates (such as the `9/25/25026` typo in the
  original template).
- Skips unused rows, collapses wrapped text, and adds the **N** prefix when it's missing (the same
  rule as the app).
- Treats `Not found` or `#REF!` from the Production Schedule lookups as blank and returns a warning
  for that row.
- If a job number appears twice, imports the first row and warns about the second.
- Ignores all four comment columns.

About the P: drive: Supervisor, Builder, Subdivision and Lot # are XLOOKUPs into the Production
Schedule on the P: drive. The cloud can't reach P:, so the sheet must be filled in **desktop
Excel**. The Send button (Phase 3A) replaces those lookups with their values before sending.

---

## Phase 3 — Power Automate import flow (as built and tested)

Automated cloud flow **Finish Crew Cover Sheet Import**. Final order:

```
Trigger → Delay → Copy file → Run script ─┬→ Apply to each → Delete file → Success email
                                          └→ Error email (Run script has failed)
```

1. **Trigger:** SharePoint → *When a file is created (properties only)*
   - Site Address: IndyWarrantyTracking-NRG365, Library Name: Documents
   - Folder: **General › Cover Sheet Import** (use the folder picker)
   - Trigger condition (Settings → Trigger conditions):
     `@endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.xlsx')`
   - Only a **new** file starts the flow. Renaming or overwriting an existing file does not.
2. **Delay:** 1 minute. Gives SharePoint time to finish saving the upload.
3. **Copy file** (SharePoint)
   - File to Copy: **fx** `triggerBody()?['{Identifier}']` (Identifier, **not** ID)
   - Destination Folder: **General › Cover Sheet Imported**
   - If another file is already there: **Copy with a new name**
4. **Run script** (Excel Online (Business)), on the **copy**, so the original is never opened
   and never locked:
   - Location: the SharePoint site, Document Library: Documents
   - File: **fx** `outputs('Copy_file')?['body/Id']`
   - Script: **Import Finish Cover Sheet**
5. **Apply to each** over `outputs('Run_script')?['body/result/rows']`
   1. **Get items** (SCHEDULED JOBS). Filter Query, as **one fx expression**:
      `concat('ImportKey eq ''', items('Apply_to_each')?['importKey'], '''')`. Top Count: `1`.
   2. **Condition**, one row: **fx** `length(outputs('Get_items')?['body/value'])` *is equal to* `0`
      - **True → Create item** (SCHEDULED JOBS):

        | Field | Value (all **fx**) |
        |---|---|
        | Job Number (Title) | `items('Apply_to_each')?['jobNumber']` |
        | JobDate | `outputs('Run_script')?['body/result/jobDate']` |
        | Supervisor | `items('Apply_to_each')?['supervisor']` |
        | LeadInstaller | `items('Apply_to_each')?['leadInstaller']` |
        | HelperInstaller | `items('Apply_to_each')?['helperInstaller']` |
        | Builder | `items('Apply_to_each')?['builder']` |
        | Subdivision | `items('Apply_to_each')?['subdivision']` |
        | LotNumber | `items('Apply_to_each')?['lot']` |
        | Phase | `items('Apply_to_each')?['phase']` |
        | JobStatus | `Open` (typed) |
        | ImportKey | `items('Apply_to_each')?['importKey']` |

      - **False → Condition 1**, one row: **fx**
        `first(outputs('Get_items')?['body/value'])?['JobStatus']` *is equal to* `Open` (typed)
        - **True → Update item**. Id: **fx** `first(outputs('Get_items')?['body/value'])?['ID']`,
          same fields as above.
        - **False →** nothing. Submitted jobs are never overwritten.

   **Designer rule:** inside this loop, enter every value with **fx**. Picking Get items fields
   from the dynamic-content list makes the designer add a hidden **For each**, and nothing gets
   created.
6. **Delete file** (SharePoint), **below** Apply to each (not inside it). File Identifier:
   **fx** `triggerBody()?['{Identifier}']`. Removes the original, so Cover Sheet Import stays
   empty.
7. **Success email** (Send an email (V2)), below Delete file. **Run after → Delete file:** Is
   successful **and** Has failed.
   - **To:** **fx** `triggerOutputs()?['body/Author/Email']` (the person who sent the sheet)
   - **Subject:** `Cover Sheet imported – ` + **fx** `outputs('Run_script')?['body/result/jobDate']`
   - **Body:** **fx** `length(outputs('Run_script')?['body/result/rows'])`, ` jobs imported.`,
     Enter, **fx** `join(outputs('Run_script')?['body/result/warnings'], '<br>')`

   If any job fails inside the loop, Delete file and this email are both skipped. So "imported"
   is only sent when every job was saved.
8. **Error email**, on a parallel branch after Run script. **Run after → Run script:** Has
   failed only.
   - **To:** **fx** `triggerOutputs()?['body/Author/Email']`
   - **Subject:** `Cover Sheet NOT imported`
   - **Body:** `The cover sheet could not be imported: ` + **fx**
     `coalesce(outputs('Run_script')?['body/error/message'], outputs('Run_script')?['body/message'], 'See the flow run history for details.')`

---

## Phase 3A — Office hand-off: fill in the sheet, click Send

The Production office person's whole job each day:

1. Open **FINISH COVER SHEET – DAILY ENTRY.xlsm** in desktop Excel.
2. Enter the job date in B4, then the Job Numbers, Lead, Helper and Phase. The Production
   Schedule lookups fill in the rest.
3. Click **Send to Finish Crew App**.

The button (macro `excel-macros/SendToFinishCrewApp.bas`) does the following:
- Refreshes the P: drive lookups.
- Checks that B4 is a real date and at least one job is entered.
- Warns about any "Not found" lookups.
- Saves a values-only `.xlsx` copy straight into **Cover Sheet Import**, named like
  `Finish Cover 2026-09-30 sent 2026-09-29 161502.xlsx`.

The flow does the rest, and the sender gets the "imported" email a few minutes later.

**Corrections:** fix the sheet and click **Send** again. Each click creates a new uniquely named
file, so the flow always starts. Jobs still Open are updated and submitted jobs are left alone.

### Setup (one time)

1. **Make the entry workbook.** Open the current cover sheet template in desktop Excel → **File →
   Save As** → type **Excel Macro-Enabled Workbook (*.xlsm)** → name it
   `FINISH COVER SHEET – DAILY ENTRY.xlsm`. Save it where the office person works (their desktop
   or a P: drive folder).
   *Keep the original **MASTER FINISH COVER SHEET TEMPLATE.xlsx** unchanged. The daily Cover
   Sheet report flow copies that file, and it must stay .xlsx.*
2. **Add the macro.** Press **Alt+F11** → **File → Import File…** → choose
   `SendToFinishCrewApp.bas` → close the VBA window.
3. **Add the button.** If there's no **Developer** tab: File → Options → Customize Ribbon →
   check **Developer**. Then **Developer → Insert → Button (Form Control)**, draw it near the top
   of the sheet (for example beside B4), pick **SendToFinishCrewApp** when asked, and change its
   text to **Send to Finish Crew App**.
4. **Add date validation to B4:** Data → Data Validation → Allow **Date**, between 1/1/2025 and
   12/31/2030.
5. **Save** the .xlsm.
6. **Macro security.** If Excel shows a yellow **Enable Content** bar, click it. If macros are
   blocked by policy, ask IT to add the entry workbook's folder as a **Trusted Location**.
7. **Test:** enter a date and one job, click **Send**. The file appears in Cover Sheet Import, then
   moves to Cover Sheet Imported within a few minutes, the job shows in SCHEDULED JOBS, and the
   "imported" email arrives.

The office person must be signed in to Excel with their Airtron account. The macro saves directly
to the SharePoint web address, so no OneDrive sync is needed.

### Clear button (start the next day)

Macro `excel-macros/ClearCoverSheet.bas` resets the daily-entry workbook after the sheet has been
sent. Import it the same way (Alt+F11 → File → Import File, or paste into a new module) and add a
second button, **Clear Sheet**, assigned to **ClearCoverSheet**.

It shows two warnings (job date and job count, then a final "cannot be undone"; **No** is the
default button on both), then:
- clears typed entries (Job Number, Lead, Helper, Phase, the four comment columns) and B4
- keeps the P: drive lookup formulas in Supervisor, Builder, Subdivision and Lot #, and puts back
  any formula someone typed over
- keeps the Lead/Helper drop-down menus, formatting and buttons, and resets row heights

### If the macro can't save to SharePoint

Some setups block saving to a web address from a macro. Use the synced folder instead:

1. In SharePoint, open **General › Cover Sheet Import** → **Sync** (or **Add shortcut to My
   files**). The folder then appears in File Explorer.
2. In File Explorer, right-click the folder → copy its path, for example
   `C:\Users\<name>\Airtron\IndyWarrantyTracking-NRG365 - Documents\General\Cover Sheet Import\`.
3. In the macro, replace the `IMPORT_FOLDER` web address with that path (keep the trailing `\`).

### No-macro alternative

If macros aren't allowed at all: after filling in the sheet, **File → Save a Copy** (or Save
As) into the synced **Cover Sheet Import** folder with a new name that includes the date, then
**close** the workbook. It's two more clicks than the button, and the office person has to
remember to use a new name each time.

---

## Phase 3B — Daily Finish Crew Cover Sheet report (new one-row-per-job format)

The report template is now **FINISH COVER SHEET MASTER.xlsx**, the same layout the office fills
in: sheet **MASTER COVER**, job date in **B4**, table **FINISHCOVER** (header row 6) with one row
per job:

Job Number | Supervisor | Lead Installer | Helper Installer | Builder | Subdivision | Lot # |
Phase | General Comments | Builder Comments | Field Super Comments | FQI Comments

Script: `office-scripts/Populate Finish Crew Cover Sheet.ts`. Paste it over the old
**POPULATE FINISH CREW COVER SHEET** script (Notepad → Ctrl+A, Ctrl+C → empty editor → Ctrl+V →
Save). It keeps the same parameters (`recordsJson`, `reportDate`), so the flow's Run script step
needs no changes. Point the flow's "copy template" step at the new master file.

What it does:
- Writes the job date into B4 as a real date (keeps the cell's date format).
- Clears the table body, including the template's P: drive XLOOKUP formulas, and writes one row
  per submission from the app.
- Builder uses **BuilderName**, falling back to the old Builder lookup for older cards. General
  Comments uses the **SupervisorComments** internal column.
- Job Number and Lot # are written as text ("9B", "0110" stay exact).
- Wraps text, fits row heights, sets the print area to the filled rows, and repeats row 6 on
  every printed page.
- Grows the table if a day ever has more than 94 jobs.

Street Address and Zip Code are no longer on the report.

Optional clean-up: the master still carries the external link to the P: drive Production
Schedule. It does no harm, but desktop Excel may ask about updating links. To remove it, open the
master in desktop Excel → **Data → Edit Links → Break Link** → save. Only do this on the report
master, not on the office's daily-entry workbook, which needs the lookups.

---

## Phase 4 — Restore pay splits and fix the pay math

**Before you start:**
- Save the app with a note such as "Before Phase 4" (restore later from **Details → Versions**
  if needed).
- **Installer names must match exactly** in three places, or the pay split (and later the job
  list and emails) won't find the person:
  1. **FINISH INSTALLER PAY SPLITS**: Lead Installer and Helper Installer columns
  2. **FINISH CREW EMAIL LIST**: Employee Name
  3. What the Production office types on the cover sheet

  Examples found in the pay splits list: "Matt Montgomery" / "Matthew Montgomery", "Lal Ro" /
  "Lal Ropuia", "Bryant Barrios" / "Bryant Barrios Chavez". Pick one spelling per person. A
  drop-down list (Data Validation → List) on the cover sheet's installer columns prevents typos.
- Pay splits are stored as whole numbers (50/50, 55/45) and always total 100. The formulas below
  handle that as written.

### 4.1 App.OnStart: add at the end

```
ClearCollect(colPaySplits, 'FINISH INSTALLER PAY SPLITS');
Set(varJobInfoLoaded, false);
```
(`varSchedJob` is added in Phase 6. Setting it before the job picker exists shows as an error.)

The pay split list is small, so it's loaded once into memory. Lookups are then instant, have no
delegation limit, and can ignore upper/lower case and extra spaces in names.

On **scrHome**, add the same `ClearCollect(colPaySplits, 'FINISH INSTALLER PAY SPLITS');` to the
Refresh Lists button so a pay-split change is picked up without restarting the app.

### 4.2 scrJobInfo: one hidden button holds all pay-split logic

Add a **Button** to scrJobInfo (outside the form) named **btnLoadPaySplit**. Set **Visible** to
`false`. It's called from several places, so the logic exists only once.

`btnLoadPaySplit.OnSelect`
```
With(
    {
        lead: Lower(Trim(DataCardValue7.Text)),
        helper: Lower(Trim(DataCardValue8.Text))
    },
    With(
        {
            pair: LookUp(
                colPaySplits,
                Lower(Trim('Lead Installer')) = lead &&
                Lower(Trim(Coalesce('Helper Installer', ""))) = helper
            ),
            leadRow: LookUp(colPaySplits, Lower(Trim('Lead Installer')) = lead),
            helperRow: LookUp(colPaySplits, Lower(Trim(Coalesce('Helper Installer', ""))) = helper)
        },
        With(
            {
                // Handles the list storing 60/40 or 0.60/0.40
                scale: If(
                    Coalesce(pair.'Lead Percentage', 0) + Coalesce(pair.'Helper Percentage', 0) <= 1,
                    100,
                    1
                )
            },
            Set(varPairFound, !IsBlank(pair));

            // Hourly rates: use the pair's row, else any row for that person
            Set(
                varLeadHourlyRate,
                If(lead = "", 0, Coalesce(pair.'Lead Hourly Pay Rate', leadRow.'Lead Hourly Pay Rate', 0))
            );
            Set(
                varHelperHourlyRate,
                If(helper = "", 0, Coalesce(pair.'Helper Hourly Pay Rate', helperRow.'Helper Hourly Pay Rate', 0))
            );

            // Percentages: keep the saved ones while editing, unless the installers were changed
            If(
                !varKeepSavedSplit,
                Set(
                    varLeadPercent,
                    If(
                        !IsBlank(pair), pair.'Lead Percentage' * scale,
                        lead <> "" && helper = "", 100,
                        Blank()
                    )
                );
                Set(
                    varHelperPercent,
                    If(
                        !IsBlank(pair), Coalesce(pair.'Helper Percentage', 0) * scale,
                        lead <> "" && helper = "", 0,
                        Blank()
                    )
                )
            )
        )
    )
)
```

This button does the following:
- **Loads both the percentages and the hourly rates.** The percentage lookup was lost when the
  rate lookup was added.
- **Handles lead-only jobs.** With no helper, the lead gets 100% and their hourly rate still
  loads. Before, the lookup ran only in Helper Installer's OnChange, so a lead-only job never
  loaded a rate.
- **Uses each installer's own rate when the pair isn't in the list.** Their percentages then stay
  blank, and a red message (4.4) tells the user to enter them.

### 4.3 scrJobInfo: when the lookup runs

**scrJobInfo.OnVisible** (the screen itself):
```
If(
    !varJobInfoLoaded,
    If(
        varEditMode,
        Set(varKeepSavedSplit, true);
        Set(varLeadPercent, Value(Substitute(varEditSubmission.'Lead Installer % of Pay', "%", "")));
        Set(varHelperPercent, Value(Substitute(varEditSubmission.'Helper Installer % of Pay', "%", ""))),
        Set(varKeepSavedSplit, false)
    );
    Select(btnLoadPaySplit);
    Set(varJobInfoLoaded, true)
)
```
It runs once each time a time card is opened, not when the user comes back from scrPayItems.
That way a manual percentage override isn't wiped out.

**DataCardValue7.OnChange** (Lead Installer text box) and **DataCardValue8.OnChange** (Helper
Installer text box). Replace the Helper's current rate formula; both get the same formula:
```
Set(varKeepSavedSplit, false);
Select(btnLoadPaySplit)
```

**DataCardValue9.OnChange** and **DataCardValue10.OnChange** (Lead/Helper Shop Hours): **delete
these formulas**. Shop pay now recalculates on its own (next step), whichever field is filled
in first.

### 4.4 scrJobInfo: pay fields

| Control | Property | Formula |
|---|---|---|
| DataCardValue13 (Lead %) | Default | `varLeadPercent` *(unchanged)* |
| Lead Installer % of Pay_DataCard1 | Update | `Substitute(Trim(DataCardValue13.Text), "%", "")` |
| DataCardValue12 (Helper %) | Default | `varHelperPercent` *(unchanged)* |
| Helper Installer % of Pay_DataCard1 | Update | `Substitute(Trim(DataCardValue12.Text), "%", "")` |

**DataCardValue20.Default** (Lead Installer Total Shop Pay):
```
If(
    varEditMode && varKeepSavedSplit &&
        Value(DataCardValue9.Text) = varEditSubmission.'Lead Shop Hours',
    Parent.Default,
    Round(Coalesce(Value(DataCardValue9.Text), 0) * Coalesce(varLeadHourlyRate, 0), 2)
)
```

**DataCardValue21.Default** (Helper Installer Total Shop Pay):
```
If(
    varEditMode && varKeepSavedSplit &&
        Value(DataCardValue10.Text) = varEditSubmission.'Helper Shop Hours',
    Parent.Default,
    Round(Coalesce(Value(DataCardValue10.Text), 0) * Coalesce(varHelperHourlyRate, 0), 2)
)
```
When you edit a saved time card, its saved shop pay is kept unless the hours or installers change.
An old card isn't silently repriced at today's rate.

**New label `lblSplitError`** (below the percentage fields). Color red, **Visible:**
`Self.Text <> ""`. **Text:**
```
With(
    {
        lp: Value(Substitute(DataCardValue13.Text, "%", "")),
        hp: Value(Substitute(DataCardValue12.Text, "%", "")),
        hasHelper: !IsBlank(Trim(DataCardValue8.Text))
    },
    If(
        IsBlank(Trim(DataCardValue7.Text)), "",
        IsBlank(lp) || (hasHelper && IsBlank(hp)),
            "No pay split on file for this Lead/Helper pair. Enter the % of Pay for each installer.",
        Round(lp + If(hasHelper, hp, 0), 2) <> 100,
            "Lead % + Helper % must equal 100 (currently " & Text(lp + If(hasHelper, hp, 0)) & ").",
        ""
    )
)
```
*If your splits can legitimately total something other than 100, delete the middle
`Round(...) <> 100` condition and its message.*

**Job Date**: the date picker inside Job Date_DataCard1 (DataCardValue2) has
`DefaultDate = Today()`, which ignores the saved date. Editing an old card overwrites its Job Date
with today's date. Fix:

| Control | Property | Formula |
|---|---|---|
| DataCardValue2 | DefaultDate | `Coalesce(Parent.Default, Today())` |

### 4.5 scrJobInfo: buttons and OnSuccess

**Next – Add Pay Items** OnSelect (replace):
```
If(
    IsBlank(DataCardValue1.Text),
    Notify("Please enter a Job Number before continuing.", NotificationType.Error),

    lblSplitError.Text <> "",
    Notify(lblSplitError.Text, NotificationType.Error),

    Set(varLeadPercent, Value(Substitute(DataCardValue13.Text, "%", "")));
    Set(varHelperPercent, Coalesce(Value(Substitute(DataCardValue12.Text, "%", "")), 0));
    Set(varLeadShopPay, Coalesce(Value(DataCardValue20.Text), 0));
    Set(varHelperShopPay, Coalesce(Value(DataCardValue21.Text), 0));
    SubmitForm(frmSubmission)
)
```

**btnSaveChanges** OnSelect (replace). The old version navigated away before the save finished and
showed "Changes Saved" twice:
```
If(
    lblSplitError.Text <> "",
    Notify(lblSplitError.Text, NotificationType.Error),

    !frmSubmission.Valid,
    Notify("Please complete all required fields before saving.", NotificationType.Error),

    SubmitForm(frmSubmission)
)
```

**frmSubmission.OnSuccess** (replace). Editing the header now also recalculates the saved
Lead/Helper **Total Pay**:
```
Set(varSubmissionID, frmSubmission.LastSubmit.ID);

If(
    varEditMode,

    With(
        {
            r: frmSubmission.LastSubmit,
            t: Coalesce(frmSubmission.LastSubmit.'Total Select Items Pay', 0)
        },
        Patch(
            SUBMISSIONS,
            r,
            {
                'Lead Installer Total Pay': Round(
                    Coalesce(r.'Lead Installer Total Shop Pay', 0) +
                    t * Coalesce(Value(Substitute(r.'Lead Installer % of Pay', "%", "")), 0) / 100,
                    2
                ),
                'Helper Installer Total Pay': Round(
                    Coalesce(r.'Helper Installer Total Shop Pay', 0) +
                    t * Coalesce(Value(Substitute(r.'Helper Installer % of Pay', "%", "")), 0) / 100,
                    2
                )
            }
        )
    );
    Set(varEditMode, false);
    Notify("Changes Saved", NotificationType.Success);
    Navigate(scrOpenSubmissions, ScreenTransition.Fade),

    Navigate(scrPayItems, ScreenTransition.Fade)
)
```

### 4.6 scrPayItems

**QuantityInput.OnChange** (inside galSelectItems):
```
Patch(
    colLineItems,
    ThisItem,
    {
        Quantity: Coalesce(Value(Self.Text), 0),
        'Extended Total': Round(Coalesce(Value(Self.Text), 0) * ThisItem.Rate, 2)
    }
)
```

**txtRate.OnChange** (inside galSelectItems):
```
Patch(
    colLineItems,
    ThisItem,
    {
        Rate: Coalesce(Value(Self.Text), 0),
        'Extended Total': Round(ThisItem.Quantity * Coalesce(Value(Self.Text), 0), 2)
    }
)
```

**btnToLineItems.OnSelect** (Submit Time Card). Replace the whole formula with:
```
If(
    CountRows(colLineItems) = 0 &&
    Coalesce(varLeadShopPay, 0) = 0 &&
    Coalesce(varHelperShopPay, 0) = 0,

    Notify(
        "You must add at least one pay item or enter shop pay before submitting.",
        NotificationType.Error
    ),

    Set(
        varJobNumber,
        LookUp(
            SUBMISSIONS,
            ID = varSubmissionID
        ).'Job Number'
    );

    // ---- Totals (rounded to the cent) ----
    Set(
        varSelectedItemsTotal,
        Round(Sum(colLineItems, 'Extended Total'), 2)
    );

    Set(
        varLeadInstallerTotalPay,
        Round(
            Coalesce(varLeadShopPay, 0) +
            varSelectedItemsTotal * Coalesce(varLeadPercent, 0) / 100,
            2
        )
    );

    Set(
        varHelperInstallerTotalPay,
        Round(
            Coalesce(varHelperShopPay, 0) +
            varSelectedItemsTotal * Coalesce(varHelperPercent, 0) / 100,
            2
        )
    );

    /*
       When editing an existing Time Card, remove its old
       LINE ITEMS records before saving the current collection.
    */
    If(
        varEditMode,
        RemoveIf(
            'LINE ITEMS',
            'Time Card Number' = Value(varSubmissionID)
        )
    );

    /*
       Recreate LINE ITEMS from the current contents of
       colLineItems. Deleted collection rows are not recreated.
    */
    ForAll(
        colLineItems,
        Patch(
            'LINE ITEMS',
            Defaults('LINE ITEMS'),
            {
                'Time Card Number': Value(varSubmissionID),
                'Job Number': varJobNumber,
                'Item Code': Coalesce(
                    'ItemCode',
                    ItemCode
                ),
                'Pay Item': 'Pay Item',
                Category: Category,
                Quantity: Quantity,
                Hours: Hours,
                Rate: Rate,
                'Extended Total': 'Extended Total',
                'Display Order': 'Display Order'
            }
        )
    );

    Patch(
        SUBMISSIONS,
        LookUp(
            SUBMISSIONS,
            ID = varSubmissionID
        ),
        {
            'Total Select Items Pay': varSelectedItemsTotal,
            'Lead Installer Total Pay': varLeadInstallerTotalPay,
            'Helper Installer Total Pay': varHelperInstallerTotalPay,
            Status: {
                Value: "Submitted"
            }
        }
    );

    // (Phase 6.5 inserts the "mark scheduled job submitted" block here.)

    Refresh('LINE ITEMS');
    Refresh(SUBMISSIONS);

    Set(
        varSelectedSubmission,
        LookUp(
            SUBMISSIONS,
            ID = varSubmissionID
        )
    );

    Clear(colLineItems);

    // ---- Reset for the next time card ----
    Set(varTotalPay, 0);
    Set(varSelectedItemsTotal, 0);
    Set(varLeadInstallerTotalPay, 0);
    Set(varHelperInstallerTotalPay, 0);
    Set(varLeadShopPay, 0);
    Set(varHelperShopPay, 0);
    Set(varLeadPercent, 0);
    Set(varHelperPercent, 0);
    Set(varLeadHourlyRate, 0);
    Set(varHelperHourlyRate, 0);
    Set(varEditMode, false);

    Notify(
        "Time Card Submitted",
        NotificationType.Success
    );

    Navigate(
        scrOpenSubmissions,
        ScreenTransition.Fade
    )
)
```
Changes from the original: totals rounded to the cent, and the hourly rates reset **at the end**
(not before saving). Clearing them earlier would zero shop pay if the user went back to
scrJobInfo. The odd-looking `'Item Code': Coalesce('ItemCode', ItemCode)` is original and keeps
item codes when a card is edited.

### 4.7 scrTimeCardDetail

**btnEditPayItems.OnSelect**. Replace the whole formula with:
```
Set(
    varEditMode,
    true
);

Set(
    varSubmissionID,
    varSelectedSubmission.ID
);

Set(
    varLeadPercent,
    Coalesce(Value(Substitute(varSelectedSubmission.'Lead Installer % of Pay', "%", "")), 0)
);

Set(
    varHelperPercent,
    Coalesce(Value(Substitute(varSelectedSubmission.'Helper Installer % of Pay', "%", "")), 0)
);

Set(
    varLeadHourlyRate,
    LookUp(
        'FINISH INSTALLER PAY SPLITS',
        'Lead Installer' = varSelectedSubmission.'Lead Installer' &&
        'Helper Installer' = varSelectedSubmission.'Helper Installer'
    ).'Lead Hourly Pay Rate'
);

Set(
    varHelperHourlyRate,
    LookUp(
        'FINISH INSTALLER PAY SPLITS',
        'Lead Installer' = varSelectedSubmission.'Lead Installer' &&
        'Helper Installer' = varSelectedSubmission.'Helper Installer'
    ).'Helper Hourly Pay Rate'
);

Set(
    varLeadShopPay,
    Coalesce(
        varSelectedSubmission.'Lead Installer Total Shop Pay',
        0
    )
);

Set(
    varHelperShopPay,
    Coalesce(
        varSelectedSubmission.'Helper Installer Total Shop Pay',
        0
    )
);

ClearCollect(
    colLineItems,
    ForAll(
        Filter('LINE ITEMS', 'Time Card Number' = varSelectedSubmission.ID) As li,
        {
            ItemCode: li.'Item Code',
            'Pay Item': li.'Pay Item',
            Category: li.Category,
            Quantity: li.Quantity,
            Hours: li.Hours,
            Rate: li.Rate,
            'Extended Total': li.'Extended Total',
            'Display Order': li.'Display Order',
            RequiresManualRate: LookUp('PAY SCHEDULE MASTER', 'Item Code' = li.'Item Code').'Requires Manual Rate'
        }
    )
);

Navigate(
    scrPayItems,
    ScreenTransition.Fade
)
```
- The old `ClearCollect(colLineItems, RenameColumns(...))` is **removed**. It dropped
  **RequiresManualRate**, so manual-rate items behaved backwards when a card was re-edited, and it
  would overwrite the new collection if left in.
- If `Category: li.Category` shows red, LINE ITEMS Category is a Choice column: use
  `Category: li.Category.Value`.

**btnEditTimeCard.OnSelect**: add at the top:
```
Set(varJobInfoLoaded, false);
```
(Phase 6 adds `Set(varSchedJob, Blank());` right after it.)

**btnApproved.OnSelect**. Replace the whole formula with:
```
Patch(
    SUBMISSIONS,
    varSelectedSubmission,
    {
        Status: {
            Value: "Approved"
        },
        'Supervisor Approval': User().FullName,
        'Approved Date': Now(),
        'General Comments':
            varSelectedSubmission.'General Comments'
    }
);

Office365Outlook.SendEmailV2(

    LookUp(
        'FINISH CREW EMAIL LIST',
        'Employee Name' = varSelectedSubmission.'Lead Installer'
    ).'Email Address' & ";" &

    LookUp(
        'FINISH CREW EMAIL LIST',
        'Employee Name' = varSelectedSubmission.'Helper Installer'
    ).'Email Address',

    "FINISH CREW TIME CARD #" &
    varSelectedSubmission.ID &
    " APPROVED",

    "Your Finish Crew Time Card has been approved." &

    "<br><br><b>Time Card #:</b> " &
    varSelectedSubmission.ID &

    "<br><b>Job Number:</b> " &
    varSelectedSubmission.'Job Number' &

    "<br><b>Job Date:</b> " &
    Text(
        varSelectedSubmission.'Job Date',
        "mm/dd/yyyy"
    ) &

    "<br><b>Builder:</b> " &
    Coalesce(varSelectedSubmission.'Builder Name', varSelectedSubmission.Builder.Value) &

    "<br><b>Subdivision:</b> " &
    varSelectedSubmission.Subdivision &

    "<br><b>Lot Number:</b> " &
    varSelectedSubmission.'Lot Number' &

    "<br><b>Street Address:</b> " &
    varSelectedSubmission.'Street Address' &

    "<br><b>Zip Code:</b> " &
    varSelectedSubmission.'Zip Code' &

    "<br><b>Supervisor:</b> " &
    varSelectedSubmission.Supervisor &

    "<br><br><b>Lead Installer:</b> " &
    varSelectedSubmission.'Lead Installer' &

    "<br><b>Lead Installer % of Pay:</b> " &
    Text(
        Coalesce(
            Value(Substitute(varSelectedSubmission.'Lead Installer % of Pay', "%", "")),
            0
        ),
        "0.##"
    ) &
    "%" &

    "<br><b>Lead Installer Total Pay:</b> " &
    Text(
        Coalesce(
            varSelectedSubmission.'Lead Installer Total Pay',
            0
        ),
        "$#,##0.00"
    ) &

    "<br><br><b>Helper Installer:</b> " &
    Coalesce(
        varSelectedSubmission.'Helper Installer',
        ""
    ) &

    "<br><b>Helper Installer % of Pay:</b> " &
    Text(
        Coalesce(
            Value(Substitute(varSelectedSubmission.'Helper Installer % of Pay', "%", "")),
            0
        ),
        "0.##"
    ) &
    "%" &

    "<br><b>Helper Installer Total Pay:</b> " &
    Text(
        Coalesce(
            varSelectedSubmission.'Helper Installer Total Pay',
            0
        ),
        "$#,##0.00"
    ) &

    "<br><br><b>Total Selected Items Pay:</b> " &
    Text(
        Coalesce(
            varSelectedSubmission.'Total Select Items Pay',
            0
        ),
        "$#,##0.00"
    ) &

    "<br><br><b>Approved By:</b> " &
    User().FullName &

    "<br><b>Approved Date:</b> " &
    Text(
        Now(),
        "mm/dd/yyyy h:mm AM/PM"
    ) &

    "<br><br><b>General Comments:</b>" &
    "<br>" &
    Coalesce(
        DataCardValue11.Text,
        ""
    ) &

    "<br><br>No further action is needed for this approved time card.",

    {
        Importance: "Normal"
    }
);

Notify(
    "Time Card Approved",
    NotificationType.Success
);

Notify(
    "Approval Email Sent",
    NotificationType.Success
);

Set(
    varSelectedSubmission,
    Blank()
);

Navigate(
    scrOpenSubmissions,
    ScreenTransition.Fade
)
```
- Both `% of Pay` lines strip a typed "%" before converting, so "60%" no longer shows as 0.6%.
- The Builder line uses the Phase 5 **Builder Name** column. If you do Phase 4 before creating
  that column, temporarily use `varSelectedSubmission.Builder.Value &` on that line.
- Every piece of the email body must be joined with `&`. One missing `&` turns every line after it
  red.

**Test Phase 4 now** (section 7.2) before continuing.

---

## Phase 5 — Builder as plain text

### 5.1 Form (scrJobInfo)
1. Select **Builder_DataCard1** in frmSubmission and **delete** it. Hiding it isn't enough,
   because a hidden card still saves into the old lookup.
2. frmSubmission → **Edit fields** → **Add field** → **Builder Name**. Drag it to where Builder
   was, and give it TabIndex 3 (the old Builder's position) so tabbing order stays the same.
3. The new card's Default is `ThisItem.'Builder Name'`. Phase 6 changes it for prefill.

### 5.2 Display (old records fall back to the old lookup)

| Screen / control | Text |
|---|---|
| scrOpenSubmissions → job/builder label | `"Job: " & ThisItem.'Job Number' & "   Builder: " & Coalesce(ThisItem.'Builder Name', ThisItem.Builder.Value)` |
| scrOpenSubmissions → subdivision/lot label | `"Subdivision: " & ThisItem.Subdivision & "   Lot: " & ThisItem.'Lot Number'` (delete the separate Lot label) |
| scrTimeCardDetail → lblBuilderName | `"Builder: " & Coalesce(varSelectedSubmission.'Builder Name', varSelectedSubmission.Builder.Value)` |
| scrTimeCardDetail → btnApproved email | Already done in the full formula in 4.7 |

To put Builder on its own line, use `Char(10)` in place of the spaces and set the label's
**AutoHeight** to `true`.

### 5.3 Existing flows
Anywhere the daily Cover Sheet flow or the approval/rejection flows read Builder
(`item()?['Builder']?['Value']` or the dynamic content *Builder Value*), use this instead:
```
coalesce(item()?['BuilderName'], item()?['Builder']?['Value'])
```
In the Cover Sheet report flow, this is the value passed to the Office Script for the Builder
column.

---

## Phase 6 — Job picker and prefill

### 6.1 Data source and deferred lines
Power Apps → Data → Add data → SharePoint → IndyWarrantyTracking-NRG365 → **SCHEDULED JOBS**.

**Order matters:** build **scrSelectJob and galSchedJobs.OnSelect (6.3) first**. Its
`Set(varSchedJob, ThisItem)` tells Power Apps what `varSchedJob` is. Until it exists, any
`Set(varSchedJob, Blank())` shows as an error, and an error anywhere in App.OnStart stops **all**
of OnStart from running (pay splits and supervisor access included).

After 6.3 is in place, `Set(varSchedJob, Blank());` goes in these five places:

| # | Where | Position |
|---|---|---|
| 1 | App → OnStart | At the end (see below) |
| 2 | scrHome → New Time Card → OnSelect (6.2) | First line |
| 3 | scrSelectJob → Job Not Listed → OnSelect (6.3) | Already in that formula |
| 4 | scrTimeCardDetail → btnEditTimeCard → OnSelect | At the top, after `Set(varJobInfoLoaded, false);` |
| 5 | scrPayItems → btnToLineItems → OnSelect (6.5) | End of the "mark job submitted" block |

**App.OnStart**: add at the end:
```
Set(varSchedJob, Blank());

Set(
    varMyName,
    LookUp(
        'FINISH CREW EMAIL LIST',
        Lower('Email Address') = Lower(User().Email)
    ).'Employee Name'
);
```
`varMyName` is the signed-in person's name from the email list. The job list (6.3) uses it so each
Lead Installer sees only their own jobs. The email list is small, so a delegation warning on this
line is harmless.

### 6.2 scrHome: New Time Card button, OnSelect (replace)
```
Set(varSchedJob, Blank());
Refresh('SCHEDULED JOBS');
Navigate(scrSelectJob, ScreenTransition.Fade)
```

### 6.3 New screen scrSelectJob
Copy scrOpenSubmissions' header and logo for a consistent look, then add:

- **txtSearchSchedJob** (Text input), HintText `"Search job # or lead installer"`
- **galSchedJobs** (vertical gallery)
- **btnJobNotListed** (Button), Text `"Job Not Listed"`
- **Open Submissions** button (6.7)
- A back arrow: `Navigate(scrHome, ScreenTransition.Fade)`

**scrSelectJob.OnVisible** shows jobs imported since the screen was last opened:
```
Refresh('SCHEDULED JOBS')
```

**galSchedJobs.Items** shows open jobs from the last 3 days onward (so late time cards still find
their job). **Supervisors** see every job; everyone else sees only jobs where they are the **Lead
Installer**:
```
With(
    {
        jobs: If(
            varSupervisor,
            Filter(
                'SCHEDULED JOBS',
                JobStatus = "Open",
                JobDate >= DateAdd(Today(), -3, TimeUnit.Days)
            ),
            Filter(
                'SCHEDULED JOBS',
                JobStatus = "Open",
                JobDate >= DateAdd(Today(), -3, TimeUnit.Days),
                LeadInstaller = varMyName
            )
        )
    },
    Sort(
        Filter(
            jobs,
            txtSearchSchedJob.Text = "" ||
            StartsWith(Title, txtSearchSchedJob.Text) ||
            StartsWith(LeadInstaller, txtSearchSchedJob.Text)
        ),
        JobDate,
        SortOrder.Ascending
    )
)
```
- Sorted **oldest to newest** by job date.
- To let **helpers** see their jobs too, change the last filter line to
  `LeadInstaller = varMyName || HelperInstaller = varMyName`.
- If a crew is swapped, the new lead won't see the job. A supervisor can, or the lead uses **Job
  Not Listed**.
- If the gallery shows one blank row reading "Lot" and "–", the filter returned no jobs. Test
  jobs dated more than 3 days ago are hidden; send a sheet with today's date, or temporarily
  change `-3` to `-30`.

**Gallery labels.** Each line is the **Text** of a separate label inside the gallery's first
(template) row. Add a label with Insert → Text label if the layout has only two.

| Label | Text |
|---|---|
| lblLotNumberJobDate | `ThisItem.Title & "   Lot " & ThisItem.LotNumber & "   " & Text(ThisItem.JobDate, "m/d")` |
| lblBuilderSub | `ThisItem.Builder & " – " & ThisItem.Subdivision` |
| lblInstallerPhase | `ThisItem.LeadInstaller & If(IsBlank(ThisItem.HelperInstaller), "", " / " & ThisItem.HelperInstaller) & "   " & ThisItem.Phase` |

If lines overlap, raise the gallery's **TemplateSize** (for example `110`). Set any label or
icon's OnSelect in the row to `Select(Parent)`.

**Empty-list message.** Add a label on scrSelectJob. **Visible:**
`CountRows(galSchedJobs.AllItems) = 0`. **Text:**
```
If(
    !varSupervisor && IsBlank(varMyName),
    "Your email isn't on the Finish Crew Email List. See your supervisor, or use Job Not Listed.",
    "No scheduled jobs found for you. If your job is missing, use Job Not Listed."
)
```

**galSchedJobs.OnSelect**
```
NewForm(frmSubmission);
Set(varSubmissionID, 0);
Set(varEditMode, false);
Set(varEditSubmission, Blank());
Clear(colLineItems);
Set(varSchedJob, ThisItem);
Set(varJobInfoLoaded, false);
Navigate(scrJobInfo, ScreenTransition.Fade)
```
Pay splits load automatically: scrJobInfo.OnVisible runs btnLoadPaySplit (Phase 4.3) against
the prefilled names.

**btnJobNotListed**: the way out when the job isn't on the list (left off or added after the
cover sheet was sent, sheet not imported yet, callbacks or warranty work, or dated outside the
3-day window). It opens a **blank** job info form that isn't linked to any scheduled job, like the
old New Time Card. Rename its Text if clearer, e.g. `"Enter Job Manually"`.

**OnSelect:**
```
NewForm(frmSubmission);
Set(varSubmissionID, 0);
Set(varEditMode, false);
Set(varEditSubmission, Blank());
Clear(colLineItems);
Set(varSchedJob, Blank());
Set(varJobInfoLoaded, false);
Navigate(scrJobInfo, ScreenTransition.Fade)
```

### 6.4 scrJobInfo: prefill card Defaults
Change the **card's** Default (the TypedDataCard, not the text box inside it). Edit mode is
unaffected because each formula checks `!varEditMode`.

| Card | Default |
|---|---|
| Job Number_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Title, ThisItem.'Job Number')` |
| Job Date_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.JobDate, ThisItem.'Job Date')` |
| Supervisor card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Supervisor, ThisItem.Supervisor)` |
| Lead Installer_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.LeadInstaller, ThisItem.'Lead Installer')` |
| Helper Installer_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.HelperInstaller, ThisItem.'Helper Installer')` |
| Builder Name card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Builder, ThisItem.'Builder Name')` |
| Subdivision card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Subdivision, ThisItem.Subdivision)` |
| Lot Number card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.LotNumber, ThisItem.'Lot Number')` |
| Phase card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Phase, ThisItem.Phase)` |

The four comment cards don't change.

**Lock the schedule data.** On the **text box inside** the Job Number, Supervisor, Builder Name,
Subdivision and Lot Number cards, set **DisplayMode**:
```
If(!varEditMode && !IsBlank(varSchedJob) && !IsBlank(Parent.Default), DisplayMode.View, DisplayMode.Edit)
```
Set it on the text box, not the card, so the card still saves the value. A field that came in blank
(a "Not found" lookup) stays editable. Lead, Helper and Phase stay editable because crews get
swapped during the day, and changing an installer reloads the pay split (Phase 4.3).

### 6.5 scrPayItems: mark the job submitted
In **btnToLineItems.OnSelect**, add this block right after the existing `Patch(SUBMISSIONS, …
Status: {Value: "Submitted"} …);`:
```
If(
    !varEditMode && !IsBlank(varSchedJob),
    Patch(
        SUBMISSIONS,
        LookUp(SUBMISSIONS, ID = varSubmissionID),
        { ScheduleJobID: varSchedJob.ID }
    );
    Patch(
        'SCHEDULED JOBS',
        LookUp('SCHEDULED JOBS', ID = varSchedJob.ID),
        { JobStatus: "Submitted", SubmissionID: varSubmissionID }
    )
);
Set(varSchedJob, Blank());
```

### 6.6 scrOpenSubmissions: reopen the job on delete
In **btnFINALDELETE.OnSelect**, add at the very top (before `RemoveIf('LINE ITEMS', …)`):
```
If(
    !IsBlank(varDeleteRecord.ScheduleJobID),
    Patch(
        'SCHEDULED JOBS',
        LookUp('SCHEDULED JOBS', ID = varDeleteRecord.ScheduleJobID),
        { JobStatus: "Open", SubmissionID: Blank() }
    )
);
```

### 6.7 Navigation

Every route into a **new** time card must run the same reset (NewForm plus the variable resets)
used by galSchedJobs and Job Not Listed. Leaving scrJobInfo must never require pressing **Next**,
because Next saves the record.

Put the **Open Submissions** button in the same spot (for example the header beside the logo) on
every screen that has one.

**scrSelectJob → Open Submissions button → OnSelect**
```
Refresh(SUBMISSIONS);
Navigate(scrOpenSubmissions, ScreenTransition.Fade)
```

**scrJobInfo → Open Submissions button → OnSelect.** Discards the unsaved form and clears working
variables. Nothing is created for a new card; unsaved edits to an existing card are discarded.
```
If(
    frmSubmission.Unsaved,
    Notify("Unsaved changes were discarded.", NotificationType.Warning)
);

ResetForm(frmSubmission);
Set(varEditMode, false);
Set(varEditSubmission, Blank());
Set(varSchedJob, Blank());
Set(varJobInfoLoaded, false);
Clear(colLineItems);

Refresh(SUBMISSIONS);
Navigate(scrOpenSubmissions, ScreenTransition.Fade)
```

**scrJobInfo → back arrow (ArrowBehind) → OnSelect.** Returns to where the user came from.
```
If(
    varEditMode,
    Set(varEditMode, false);
    Navigate(scrTimeCardDetail, ScreenTransition.Fade),
    Navigate(scrSelectJob, ScreenTransition.Fade)
)
```
(Replace `scrTimeCardDetail` with `scrHome` if editing should return Home.)

**Any "forward" arrow that jumps straight into scrJobInfo** (for example scrHome's
`btnToJobInfo`) skips the reset and can show the previous card's data. Delete it, or give it the
same OnSelect as **Job Not Listed** (6.3).

**scrPayItems → forward arrow to Open Submissions → OnSelect.** In new mode, **Next** has already
created the SUBMISSIONS record without pay items. This tells the user:
```
If(
    !varEditMode,
    Notify("This time card was saved without pay items. Open it from Open Submissions to finish or delete it.", NotificationType.Warning)
);
Clear(colLineItems);
Set(varEditMode, false);
Navigate(scrOpenSubmissions, ScreenTransition.Fade)
```

The scrPayItems back arrow (to scrJobInfo) and the scrSelectJob back arrow (to scrHome) stay as
they are.

### 6.8 Open Submissions: remember the search while reviewing

Supervisors filter Open Submissions by job number or Lead Installer, open a card, approve it, and
come back. The search is kept while they go back and forth, and cleared automatically when the
review is done.

| What happens | Search |
|---|---|
| Open a card, come back (Back or after approving) | **Kept** |
| Approve the **last** card that matches the search | **Cleared** |
| Back to Home, or open Open Submissions from Home or another screen | **Cleared** |
| Tap the ✕ beside the search box | **Cleared** |

**App.OnStart**: add at the end: `Set(varOpenSearch, "");`

**Sort order:** the Open Submissions gallery sorts **oldest to newest by Job Date**. In its
Items formula, the final two Sort arguments are `'Job Date', SortOrder.Ascending` (previously
`ID, SortOrder.Descending`). Use `ID` instead of `'Job Date'` to sort by when the card was
submitted.

**scrOpenSubmissions → txtSearchJobNumber**

| Property | Value |
|---|---|
| Default | `varOpenSearch` |
| OnChange | `Set(varOpenSearch, Self.Text)` |
| DelayOutput | `true` |

The gallery's Items formula doesn't change.

**Clear (✕) icon** (Icon.Cancel) at the right end of the search box.
Visible: `!IsBlank(txtSearchJobNumber.Text)`. OnSelect:
```
Set(varOpenSearch, "");
Reset(txtSearchJobNumber)
```

**Add `Set(varOpenSearch, "");` as the first line of OnSelect** on:
- scrOpenSubmissions → back arrow to Home (`ArrowBackHome_2`)
- scrHome → My Open Submissions (`btnMySubmission`)
- the Open Submissions buttons on scrSelectJob, scrJobInfo and scrPayItems (6.7)

Leave scrTimeCardDetail's **[Back]** button unchanged so the search survives.

**scrTimeCardDetail → btnApproved**: insert after the `SendEmailV2(...);` block and before
`Set(varSelectedSubmission, Blank());`:
```
If(
    varOpenSearch <> "" &&
    IsBlank(
        LookUp(
            SUBMISSIONS,
            ID <> varSelectedSubmission.ID &&
            Status.Value <> "Approved" &&
            (
                StartsWith('Job Number', varOpenSearch) ||
                StartsWith('Lead Installer', varOpenSearch)
            )
        )
    ),
    Set(varOpenSearch, "")
);
```

### 6.9 scrSelectJob: delete a scheduled job (single confirmation)

A trash icon on each job row, then one "Are you sure?" pop-up. The job is **marked Deleted**
(JobStatus = "Deleted") rather than erased. It disappears from the picker just the same, but if
the office re-sends a cover sheet that still lists it, the import flow sees a job that isn't
**Open** and leaves it alone instead of re-creating it.

**1. Trash icon in the gallery row.** Select galSchedJobs' first row → Insert → Icons → **Trash**.
Name it `icoDeleteSchedJob` and place it at the right end of the row.

| Property | Value |
|---|---|
| OnSelect | `Set(varSchedJobToDelete, ThisItem); Set(varShowSchedDelete, true)` |
| Visible | `varSupervisor` (only supervisors can delete; use `true` to allow everyone) |
| Color | `RGBA(192, 0, 0, 1)` |
| Tooltip | `"Delete this scheduled job"` |

The icon's OnSelect must **not** be `Select(Parent)`, or tapping it would also open the job.

**2. Confirmation pop-up.** On scrSelectJob (outside the gallery): Insert → **Container**, named
`conSchedDeleteConfirm`. Size it to the whole screen (X 0, Y 0, Width `Parent.Width`, Height
`Parent.Height`). **Visible:** `varShowSchedDelete`. Inside it add:

- **Rectangle** (full screen), Fill `RGBA(0, 0, 0, 0.5)`, so the screen behind is dimmed and
  can't be tapped.
- **Rectangle** (the white box, centered), Fill `White`, BorderColor `RGBA(56, 96, 178, 1)`.
- **Label** `lblSchedDeleteMsg` inside the box. **Text:**
  ```
  "Delete this scheduled job?" & Char(10) & Char(10) &
  varSchedJobToDelete.Title & "   Lot " & varSchedJobToDelete.LotNumber &
  "   " & Text(varSchedJobToDelete.JobDate, "m/d") & Char(10) &
  varSchedJobToDelete.Builder & " – " & varSchedJobToDelete.Subdivision & Char(10) & Char(10) &
  "It will be removed from the Scheduled Jobs list."
  ```
- **Button** `btnSchedDeleteNo`, Text `"No"`. **OnSelect:**
  ```
  Set(varShowSchedDelete, false);
  Set(varSchedJobToDelete, Blank())
  ```
- **Button** `btnSchedDeleteYes`, Text `"Yes, Delete"`, Fill `RGBA(192, 0, 0, 1)`. **OnSelect:**
  ```
  Patch(
      'SCHEDULED JOBS',
      LookUp('SCHEDULED JOBS', ID = varSchedJobToDelete.ID),
      { JobStatus: "Deleted" }
  );

  Set(varShowSchedDelete, false);
  Set(varSchedJobToDelete, Blank());

  Refresh('SCHEDULED JOBS');

  Notify("Scheduled job deleted.", NotificationType.Success)
  ```

**3. App.OnStart**: add at the end: `Set(varShowSchedDelete, false);`

No other changes are needed. galSchedJobs only shows `JobStatus = "Open"`, so the job drops off
the list right away. To bring one back, open SCHEDULED JOBS in SharePoint and change its JobStatus
back to `Open`.

*To erase the record instead (not recommended, because a re-sent cover sheet would re-create the
job), replace the Patch with:*
`Remove('SCHEDULED JOBS', LookUp('SCHEDULED JOBS', ID = varSchedJobToDelete.ID));`

Users who delete need **Edit** access to the SCHEDULED JOBS list.

---

## Phase 7 — Testing

### 7.1 Worked example (use it to check every screen)

Pay items: two items totaling **$400.00**. Split **60 / 40**. Lead **2 hrs @ $25.00**, Helper
**2 hrs @ $18.00**.

| Value | Math | Expected |
|---|---|---|
| Lead Total Shop Pay | 2 × 25.00 | **$50.00** |
| Helper Total Shop Pay | 2 × 18.00 | **$36.00** |
| Total Select Items Pay | 400.00 | **$400.00** |
| Lead Installer Total Pay | 50.00 + 400.00 × 60% | **$290.00** |
| Helper Installer Total Pay | 36.00 + 400.00 × 40% | **$196.00** |

Use your real pay split rows. Check these values on scrJobInfo (shop pay), on scrPayItems (the
live labels), and on scrTimeCardDetail and the approval email (saved values).

### 7.2 Phase 4 tests (pay math)
- [ ] New card: type Lead, then Helper → % and rates fill in. Then enter hours → shop pay is correct.
- [ ] New card: enter hours **first**, then the installers → shop pay still correct.
- [ ] Lead-only job (no helper) → Lead 100%, Helper 0, lead shop pay uses the lead's rate.
- [ ] Pair not in pay splits → red message appears; Next is blocked until % are entered.
- [ ] Type "60%" in Lead % → totals use 60, not 0.6.
- [ ] Override % manually, go to Pay Items, come back → the override is still there.
- [ ] Submit → SUBMISSIONS totals match 7.1.
- [ ] Next card after that → no leftover rates or % from the previous card.
- [ ] Edit Time Card, change nothing, Save Changes → %, shop pay, Job Date and totals unchanged.
- [ ] Edit Time Card, change Lead hours → shop pay and Lead Total Pay update after saving.
- [ ] Edit Pay Items, change a quantity, submit → totals update; manual-rate items keep an editable rate.

### 7.3 Phase 5 tests (Builder)
- [ ] New card saves Builder Name; it shows on Open Submissions, the detail screen and the approval email.
- [ ] An old card (lookup builder) still shows its builder everywhere.
- [ ] Daily Cover Sheet shows builders for both new and old cards.

### 7.4 Phase 6 tests (prefill)
- [ ] Import a cover sheet → jobs appear in the picker with the correct date, lot, builder and installers.
- [ ] Signed in as a **lead installer** → only that lead's jobs show. As a **supervisor** → all open jobs show.
- [ ] User not on the email list → empty-list message tells them to see a supervisor / use Job Not Listed.
- [ ] Pick a job → all eight fields plus the date prefill; schedule fields are locked; % and rates load.
- [ ] Swap the helper → split reloads for the new pair.
- [ ] Submit → the job leaves the picker; SUBMISSIONS.ScheduleJobID and SCHEDULED JOBS.SubmissionID are set.
- [ ] Job Not Listed → blank form, same as before.
- [ ] Delete that time card → the job reappears in the picker.
- [ ] Daily Cover Sheet report → unchanged output.

### 7.5 Navigation tests
- [ ] scrSelectJob → Open Submissions → nothing is created.
- [ ] New card on scrJobInfo → Open Submissions → nothing is created; the next New Time Card starts blank.
- [ ] Edited card on scrJobInfo → Open Submissions without saving → the saved values are unchanged.
- [ ] Back arrow on a **new** card → scrSelectJob; on an **edited** card → scrTimeCardDetail.
- [ ] Any forward arrow into scrJobInfo → blank form, no leftover data or pay split.
- [ ] scrPayItems (new card) → Open Submissions → warning shows; the incomplete card is listed and can be edited or deleted.
- [ ] Leave scrSelectJob, send a new cover sheet, come back → new jobs appear (OnVisible refresh).
- [ ] Open Submissions: search a lead with 3 cards → approve one → list still filtered (2 left).
- [ ] Approve the remaining matches → search clears and the full list shows.
- [ ] Search → open a card → [Back] → filter kept; back arrow to Home → reopen → unfiltered.
- [ ] ✕ clears the search.
- [ ] scrSelectJob: trash icon (supervisor) → pop-up → **No** leaves the job; **Yes, Delete** removes it from the list.
- [ ] Re-send a cover sheet containing a deleted job → it does **not** reappear (JobStatus stays Deleted).

---

## Phase 8 — Field Supervisor Punch List (SharePoint list view in Teams)

About six field supervisors follow up on **Field Supervisor Comments**. Instead of an Excel
workbook, they use a **view of the SUBMISSIONS list**, pinned as a Teams tab. It is web-only, always
current, and needs no refresh. The "Complete" status is saved on the job itself, so it can't get
out of step with the data.

- Comments with work still to do are shaded **yellow**.
- One click on the checkbox marks the item **Complete**, and the comment turns **red**. Clicking
  again undoes it.
- Each supervisor has a view showing only their own jobs.

(A Power Query workbook was tried first. It was dropped because refreshing a SharePoint list query
needs desktop Excel, and cells typed next to a query table don't stay with their rows after a
refresh.)

### 8.1 Columns in SUBMISSIONS

| Column (internal name) | Type | Purpose |
|---|---|---|
| Field Supervisor Comments (`FieldSupervisorComments`) | Single line of text | The comment. Single-line text can be filtered in views (max 255 characters) |
| Field Supervisor Comments Complete (`FieldSupervisorCommentsComplete`) | Yes/No, default No | The supervisor's checkbox; drives yellow/red |

No extra flag column or app change is needed. If Field Supervisor Comments is ever changed to
multi-line text (for comments longer than 255 characters), views can no longer filter on it. You
would then need a hidden Yes/No `HasFieldSuperComment` column, set by the app with
`!IsBlank(Trim(DataCardValue16.Text))`, and the views would filter on that instead.

### 8.2 Views

1. SUBMISSIONS → **+ Add view** → List → `Field Super Punch List`.
2. **Edit current view**:
   - **Columns:** Job Number, Job Date, Supervisor, BuilderName, Builder (old cards), Subdivision,
     Lot Number, Field Supervisor Comments, Field Supervisor Comments Complete
   - **Filter:** `Field Supervisor Comments` **is not equal to** *(leave the value blank)*, which
     means "has a comment"
   - **Sort:** Field Supervisor Comments Complete ascending (open first), then Job Date descending
3. **One view per supervisor:** from Field Super Punch List → **Save view as**
   `Punch List – ROSE02` → **Edit current view** → Filter:
   ```
   Show the items when column  [Field Supervisor Comments]  [is not equal to]  [ (blank) ]
     (•) And
   When column                 [Supervisor]                 [is equal to]      [ROSE02]
   ```
   Repeat for each supervisor code (ESCU00, GIBS03, …).
   - Don't filter on **Field Supervisor Comments Complete = Yes**. That shows only finished items.
   - Optional open-items-only view: add **And Field Supervisor Comments Complete is equal to No**.
4. To hide clutter, make Field Super Punch List the **default view**. A view can't be switched
   between public and private; to keep All Items for yourself, save a private copy
   (**Save view as**, untick **Make this a public view**), then delete the public All Items.
   Check that no flow's Get items uses "Limit Columns by View = All Items" before deleting it.

Views don't restrict access; any supervisor can switch views. For a locked-down tab per
supervisor, put a **List web part** set to that supervisor's view on a SharePoint page (command bar
off) and add the page as a Teams tab.

### 8.3 Comment colors (Field Supervisor Comments column)

Column header → **Column settings → Format this column → Advanced mode**:
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "txtContent": "@currentField",
  "style": {
    "background-color": "=if(@currentField == '', '', if([$FieldSupervisorCommentsComplete] == true, '#FF9999', '#FFFF99'))",
    "white-space": "pre-wrap",
    "padding": "4px",
    "width": "100%",
    "box-sizing": "border-box"
  }
}
```
Yellow `#FFFF99` = open, red `#FF9999` = complete, no color when there's no comment. If HTML tags
appear, set the column to **Plain text**.

### 8.4 One-click checkbox (Field Supervisor Comments Complete column)

A Yes/No column normally shows a checkbox only in grid-edit mode, and rows created before the
column existed show blank. This formatting draws a clickable checkbox on every row:
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "style": {
    "display": "flex",
    "align-items": "center",
    "justify-content": "center",
    "cursor": "pointer",
    "width": "100%",
    "height": "100%"
  },
  "attributes": {
    "title": "=if(@currentField == true, 'Complete - click to undo', 'Click to mark complete')"
  },
  "customRowAction": {
    "action": "setValue",
    "actionInput": {
      "FieldSupervisorCommentsComplete": "=if(@currentField == true, 'false', 'true')"
    }
  },
  "children": [
    {
      "elmType": "span",
      "attributes": {
        "iconName": "=if(@currentField == true, 'CheckboxComposite', 'Checkbox')"
      },
      "style": {
        "font-size": "22px",
        "color": "=if(@currentField == true, '#C00000', '#605E5C')"
      }
    },
    {
      "elmType": "span",
      "txtContent": "=if(@currentField == true, 'Complete', '')",
      "style": {
        "padding-left": "6px",
        "color": "#C00000",
        "font-weight": "600"
      }
    }
  ]
}
```
- Not done: grey empty box; the comment is yellow. Done: red ticked box + "Complete"; the comment
  is red.
- Supervisors need **Edit** access to SUBMISSIONS. Ask them to change only the checkbox.

### 8.5 Teams tab

Channel → **+** → **Lists** (or SharePoint) → **Add an existing list** → **SUBMISSIONS** → open the
Field Super Punch List view → name the tab `Field Super Punch List`.

### 8.6 Tests
- [ ] A new time card with a Field Supervisor comment shows in the punch list, shaded yellow.
- [ ] A time card without one does not appear.
- [ ] Clicking the checkbox ticks it and turns the comment red; clicking again turns it back to yellow.
- [ ] Each supervisor's view shows only their code's jobs.
- [ ] Everything works from the Teams tab in the browser and the Teams app, with no desktop Excel.

---

## Phase 9 — Scale (7,000+ jobs a year) and cleanup

SharePoint lists handle millions of items. The limits that matter are:
- **Get items returns 100 items by default, oldest first.** A flow that fetches the list and
  filters afterwards silently misses new records once the list grows. (This is what dropped the
  10/7 report.)
- **5,000-item threshold.** Past 5,000 items, filtered queries fail unless the filtered columns
  are indexed.
- **Power Apps delegation.** A filter SharePoint can't evaluate itself (warning icon) only checks
  the first 500–2,000 records.

### 9.1 Daily report flow: server-side date filter
Get items (SUBMISSIONS) → **Filter Query** (fx):
```
concat('Date ge ''', outputs('Compose_-_ReportStart'), ''' and Date lt ''', outputs('Compose_-_ReportEnd'), '''')
```
ReportStart/ReportEnd as `yyyy-MM-dd`. Top Count 500, Settings → Pagination on, threshold 5000.
Keep Filter array as a safety net. Optional: in the Condition's False branch, email "No
submissions found for <date>" before Terminate.

### 9.2 Indexes (List settings → Indexed columns)

| List | Index |
|---|---|
| SUBMISSIONS | Date (Job Date), Status, Created By, Title (Job Number), Supervisor, FieldSupervisorComments |
| LINE ITEMS | Time Card Number |
| SCHEDULED JOBS | JobDate, JobStatus, ImportKey, LeadInstaller |

Create them before a list reaches 5,000 items.

### 9.3 App: delegable filters
**scrOpenSubmissions gallery → Items:**
```
Sort(
    Filter(
        SUBMISSIONS,
        Status.Value = "Submitted",
        varSupervisor || 'Created By'.Email = User().Email,
        txtSearchJobNumber.Text = "" ||
            StartsWith('Job Number', txtSearchJobNumber.Text) ||
            StartsWith('Lead Installer', txtSearchJobNumber.Text)
    ),
    'Job Date',
    SortOrder.Ascending
)
```
- `Status.Value = "Submitted"` replaces the non-delegable `<> "Approved"`. Add
  `|| Status.Value = "Rejected"` (in parentheses) if other open statuses are used.
- `Lower()` is removed; email comparison is not case-sensitive in SharePoint.
- In **btnApproved** (6.8 search-clear check) change `Status.Value <> "Approved"` to
  `Status.Value = "Submitted"`.

The job picker, line-item lookups (Time Card Number) and ID lookups are already delegable.

### 9.4 Optional monthly cleanup flow
**First:** the daily Cover Sheet doesn't include pay amounts or pay items. Confirm payroll keeps
them elsewhere before deleting SUBMISSIONS / LINE ITEMS.

| List | Delete when |
|---|---|
| SCHEDULED JOBS | JobDate older than 60 days (any status) |
| SUBMISSIONS + LINE ITEMS | Approved, Job Date older than 365 days, and punch list finished (no field super comment, or Complete = Yes) |

Scheduled cloud flow, monthly:
1. Compose – SubmissionsCutoff: `formatDateTime(addDays(utcNow(), -365), 'yyyy-MM-dd')`
2. Compose – SchedCutoff: `formatDateTime(addDays(utcNow(), -60), 'yyyy-MM-dd')`
3. Get items (SCHEDULED JOBS), Filter Query
   `concat('JobDate lt ''', outputs('Compose_-_SchedCutoff'), '''')`, Top 5000, pagination on →
   Apply to each → Delete item.
4. Get items (SUBMISSIONS), Filter Query
   `concat('Date lt ''', outputs('Compose_-_SubmissionsCutoff'), ''' and Status eq ''Approved'' and (FieldSupervisorComments eq null or FieldSupervisorCommentsComplete eq 1)')`,
   Top 5000, pagination on → Apply to each:
   - Get items (LINE ITEMS), Filter Query `concat('TimeCardNumber eq ', <current submission ID>)`
     (check the real internal name of Time Card Number)
   - Apply to each → Delete item (line items), then Delete item (the submission)
5. Email a summary with the counts.

Run it once with Compose steps in place of the Deletes and check the counts first. Deleted items
stay in the site Recycle Bin for 93 days.

---

## Appendix — why the math was wrong past scrJobInfo

Found by reading the app export (`Finish_Crew_Timecard.msapp`). Numbers match the fixes above.

| # | Problem | Effect | Fixed in |
|---|---|---|---|
| 1 | The pay split lookup now loads only hourly rates; **% of Pay is never looked up**. `varLeadPercent` / `varHelperPercent` start blank and are only set when Next is pressed. | Unless the % was typed by hand, the pay-item share is **$0** for both installers. | 4.2 |
| 2 | The rate lookup runs only in **Helper Installer's OnChange** and needs an exact pair match. | Lead-only jobs never load a rate, and keep the **previous card's** rate. Changing the Lead after the Helper doesn't reload. | 4.2, 4.3, 4.6 |
| 3 | Shop pay is calculated only in the **Hours OnChange**. | Hours entered before the installers → shop pay **$0**. Changing an installer later doesn't recalculate. | 4.4 |
| 4 | On **Edit Time Card**, the % and Shop Pay boxes show variables, not the saved values. Those variables are reset to 0 after every submit. | Save Changes **overwrites the saved % and shop pay with 0**. | 4.3, 4.4 |
| 5 | Save Changes doesn't recalculate Lead/Helper **Total Pay**. | After a header edit, the saved totals no longer match the hours and %. | 4.5 |
| 6 | `Value("60%")` returns 0.6, then the code divides by 100. | A % typed with a % sign pays **0.6%** instead of 60%. | 4.4–4.7 |
| 7 | Job Date picker uses `DefaultDate = Today()`. | Editing a card changes its Job Date to today. | 4.4 |
| 8 | Extended totals and installer totals aren't rounded. | Stray fractions of a cent in the totals. | 4.6 |
| 9 | Edit Pay Items drops RequiresManualRate. | Manual-rate items behave backwards when re-edited. | 4.7 |
| 10 | Save Changes navigates before the save finishes. | Double "Changes Saved" message; the next screen can show stale data. | 4.5 |

The FINISH INSTALLER PAY SPLITS columns used: Lead Installer, Helper Installer, Lead Percentage,
Helper Percentage, Lead Hourly Pay Rate, Helper Hourly Pay Rate.
