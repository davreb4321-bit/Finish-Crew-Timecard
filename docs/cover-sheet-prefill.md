# Cover Sheet Prefill — Build Guide

The office keeps filling in the MASTER FINISH COVER SHEET each morning. Saving it to a SharePoint
folder imports the jobs into a **SCHEDULED JOBS** list. In the app, the installer picks the job and
the form is prefilled. The installer adds hours, comments and pay items and submits to **SUBMISSIONS**
as today. The approval flow and the daily Cover Sheet report do not change.

```
Office fills cover sheet ─► saves to "Cover Sheet Import" folder
        │  flow + Office Script "Import Finish Cover Sheet"
        ▼
SCHEDULED JOBS (one row per job per day, JobStatus = Open)
        │  app: New Time Card ─► pick job ─► form prefilled
        ▼
SUBMISSIONS + LINE ITEMS (unchanged) ─► approval flow ─► daily Cover Sheet report (unchanged)
```

## What is prefilled

| Cover sheet column | SCHEDULED JOBS column | SUBMISSIONS field (form card) |
|---|---|---|
| Job Date (cell B4) | JobDate | Job Date |
| Job Number | Title | Job Number |
| Supervisor | Supervisor | Supervisor |
| Lead Installer | LeadInstaller | Lead Installer |
| Helper Installer | HelperInstaller | Helper Installer |
| Builder | Builder | Builder (Choice) |
| Subdivision | Subdivision | Subdivision |
| Lot # | LotNumber | Lot Number |
| Phase | Phase | Phase |
| General / Builder / Field Super / FQI Comments | *not imported* | entered by installer |

Street Address and Zip Code are not on the cover sheet, so they stay manual.

---

## 1. SharePoint

### New list: SCHEDULED JOBS

| Column | Type | Notes |
|---|---|---|
| Title | Single line of text | Job Number (rename display to "Job Number") |
| JobDate | Date only | **Index** this column |
| Supervisor | Single line of text | |
| LeadInstaller | Single line of text | |
| HelperInstaller | Single line of text | |
| Builder | Single line of text | |
| Subdivision | Single line of text | |
| LotNumber | Single line of text | Text, not number (lots like "9B") |
| Phase | Single line of text | |
| JobStatus | Single line of text | Default value `Open`. **Index** this column. Text instead of Choice keeps the app filter delegable. |
| SubmissionID | Number | Filled when the time card is submitted |
| ImportKey | Single line of text | `JobNumber\|yyyy-mm-dd`. **Index** + **Enforce unique values** so a re-saved sheet can't create duplicates |

### SUBMISSIONS: add one column

| Column | Type |
|---|---|
| ScheduledJobID | Number |

### Builder choices

`Builder` in SUBMISSIONS is a **Choice** column. Every builder name the production schedule uses
(e.g. `DAVID WEEKLEY`, `PULTE HOMES OF INDIANA`) must exist as a choice, spelled exactly the same.
Otherwise turn on *Allow 'Fill-in' choices*. Without one of these, prefilled builders won't save.

### Document library folder

Create a folder, e.g. `Shared Documents/Cover Sheet Import`, and (optionally) a subfolder `Imported`.

---

## 2. Office Script

Save `office-scripts/Import Finish Cover Sheet.ts` as a new Office Script. In Excel for the web:
**Automate → New Script**, paste the file, and name it `Import Finish Cover Sheet`. Save it in the
OneDrive of the account that owns the flow.

What it does:
- Reads table **FINISHCOVER** by header name, so reordering columns won't break it.
- Reads the job date from **B4** and rejects bad dates. The template you sent has `9/25/25026` in B4,
  which is exactly the kind of typo this catches.
- Skips blank rows (the table runs to row 100), trims wrapped text, and adds the `N` prefix when it's
  missing (same rule as the app).
- Treats `Not found` / `#REF!` results from the production-schedule XLOOKUPs as blank and returns a
  warning for that row.
- Returns only the first row when the same job number is listed twice, with a warning.

Recommended template change: put **Data Validation → Date** on B4 so the office can't type an
invalid date in the first place.

---

## 3. Power Automate flow: "Import Finish Cover Sheet"

1. **Trigger:** SharePoint *When a file is created in a folder*. Site: IndyWarrantyTracking-NRG365,
   Folder: `/Shared Documents/Cover Sheet Import`.
2. **Excel Online (Business) → Run script.** Location: the same SharePoint site, Document Library:
   Documents, File: the file identifier from the trigger, Script: `Import Finish Cover Sheet`.
   - Set **Configure run after** on the next step's error branch (step 6) so a thrown script
     error emails the sender.
3. **Apply to each** over `outputs('Run_script')?['body/result/rows']`.
   1. **Get items** from SCHEDULED JOBS:
      Filter Query `ImportKey eq '@{items('Apply_to_each')?['importKey']}'`, Top Count `1`.
   2. **Condition:** `length(outputs('Get_items')?['body/value'])` is equal to `0`.
      - **Yes → Create item** in SCHEDULED JOBS:
        Title = `jobNumber`, JobDate = `body/result/jobDate`, Supervisor, LeadInstaller,
        HelperInstaller, Builder, Subdivision, LotNumber = `lot`, Phase, JobStatus = `Open`,
        ImportKey = `importKey`.
      - **No → Condition:** existing item's `JobStatus` is `Open`. If so, **Update item** with the
        same fields. This way a corrected re-save updates jobs that haven't been submitted yet and
        never changes submitted ones.
4. **Move file** to `Cover Sheet Import/Imported` (optional; keeps the inbox clean).
5. **Send an email (V2)** to the file's *Created By Email*: "Imported N jobs for {jobDate}", plus
   `join(outputs('Run_script')?['body/result/warnings'], '<br>')` when there are warnings.
6. **Error branch** (run after Run script *has failed*): email the file's *Created By Email* with
   the script error message, e.g. the B4 date problem.

A note on the lookups: Supervisor, Builder, Subdivision and Lot # in the template are XLOOKUPs into
the Production Schedule on the `P:` drive. The cloud flow can't reach `P:`. It reads the values
Excel **cached when the office saved the file**, which is fine as long as the office saves from
desktop Excel with the link updated (as they do today). Rows with no match come through blank
with a warning.

---

## 4. Power Apps changes

### 4a. Add the data source
Add **SCHEDULED JOBS** from the same SharePoint site.

### 4b. New screen `scrSelectJob`

Add a search box `txtSearchSchedJob`, a gallery `galSchedJobs`, and a button **Job Not Listed**.

`galSchedJobs.Items` shows open jobs from the last 3 days so late time cards still find their job:
```
Sort(
    Filter(
        'SCHEDULED JOBS',
        JobStatus = "Open",
        JobDate >= DateAdd(Today(), -3, TimeUnit.Days),
        txtSearchSchedJob.Text = "" ||
            StartsWith(Title, txtSearchSchedJob.Text) ||
            StartsWith(LeadInstaller, txtSearchSchedJob.Text)
    ),
    JobDate, SortOrder.Descending
)
```
Gallery labels: `ThisItem.Title & " – Lot " & ThisItem.LotNumber`,
`ThisItem.Builder & " / " & ThisItem.Subdivision`,
`ThisItem.LeadInstaller & " / " & ThisItem.HelperInstaller`, `Text(ThisItem.JobDate, "m/d")`.

`galSchedJobs.OnSelect`:
```
NewForm(frmSubmission);
Set(varSubmissionID, 0);
Set(varEditMode, false);
Set(varEditSubmission, Blank());
Clear(colLineItems);
Set(varSchedJob, ThisItem);

// Prefilling the installer boxes does not fire their OnChange,
// so load the pay-split hourly rates here (same lookup as Helper Installer OnChange).
Set(
    varSchedSplit,
    LookUp(
        'FINISH INSTALLER PAY SPLITS',
        'Lead Installer' = ThisItem.LeadInstaller &&
        'Helper Installer' = ThisItem.HelperInstaller
    )
);
Set(varLeadHourlyRate, varSchedSplit.'Lead Hourly Pay Rate');
Set(varHelperHourlyRate, varSchedSplit.'Helper Hourly Pay Rate');

Navigate(scrJobInfo, ScreenTransition.Fade)
```
If your published app also sets `varLeadPercent` / `varHelperPercent` from the pay-split list when
the installers change, add those same two `Set`s here from `varSchedSplit`. (The .msapp export
only sets the hourly rates in the Helper Installer OnChange.)

**Job Not Listed** `OnSelect` is the current New Time Card code plus clearing the scheduled job:
```
NewForm(frmSubmission);
Set(varSubmissionID, 0);
Set(varEditMode, false);
Set(varEditSubmission, Blank());
Set(varSchedJob, Blank());
Clear(colLineItems);
Navigate(scrJobInfo, ScreenTransition.Fade)
```

### 4c. Home screen
`New Time Card` button `OnSelect`:
```
Set(varSchedJob, Blank());
Refresh('SCHEDULED JOBS');
Navigate(scrSelectJob, ScreenTransition.Fade)
```

### 4d. `frmSubmission` card defaults (scrJobInfo)

Change only each **card's** `Default` (the card, not the text box inside it). Edit mode is unaffected
because every formula checks `!varEditMode`.

| Card | Default |
|---|---|
| Job Number_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Title, ThisItem.'Job Number')` |
| Job Date_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.JobDate, ThisItem.'Job Date')` |
| Supervisor card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Supervisor, ThisItem.Supervisor)` |
| Lead Installer_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.LeadInstaller, ThisItem.'Lead Installer')` |
| Helper Installer_DataCard1 | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.HelperInstaller, ThisItem.'Helper Installer')` |
| Builder card | `If(!varEditMode && !IsBlank(varSchedJob), {Value: varSchedJob.Builder}, ThisItem.Builder)` |
| Subdivision card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Subdivision, ThisItem.Subdivision)` |
| Lot Number card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.LotNumber, ThisItem.'Lot Number')` |
| Phase card | `If(!varEditMode && !IsBlank(varSchedJob), varSchedJob.Phase, ThisItem.Phase)` |

**Optional lock (recommended):** so the installer can't change schedule data by accident, set the
`DisplayMode` of the **input control inside** these cards (DataCardValue1 for Job Number, and the
Supervisor, Builder, Subdivision and Lot inputs). Lock the control rather than the card so the card's
`Update` still submits the value:
```
If(!varEditMode && !IsBlank(varSchedJob) && !IsBlank(Parent.Default), DisplayMode.View, DisplayMode.Edit)
```
For Builder, use `!IsBlank(varSchedJob.Builder)` in place of `!IsBlank(Parent.Default)`.
Leave Lead Installer, Helper Installer and Phase editable, because crews get swapped during the day.
A blank imported value (a "Not found" lookup) stays editable automatically.

### 4e. Mark the scheduled job submitted (scrPayItems)

In the submit button, add `ScheduledJobID` to the **existing** SUBMISSIONS Patch record (the one that
sets `Status: {Value: "Submitted"}`):
```
ScheduledJobID: If(IsBlank(varSchedJob), Blank(), varSchedJob.ID),
```
Then add this right after that Patch:
```
If(
    !IsBlank(varSchedJob),
    Patch(
        'SCHEDULED JOBS',
        LookUp('SCHEDULED JOBS', ID = varSchedJob.ID),
        { JobStatus: "Submitted", SubmissionID: varSubmissionID }
    );
    Set(varSchedJob, Blank())
);
```
The job then drops off everyone's pick list and can't be submitted twice.

### 4f. Re-open on delete (scrOpenSubmissions)

In the second delete-confirmation `OnSelect`, before the `RemoveIf('LINE ITEMS', …)` / `Remove(SUBMISSIONS, varDeleteRecord)` calls, add:
```
If(
    !IsBlank(varDeleteRecord.ScheduledJobID),
    Patch(
        'SCHEDULED JOBS',
        LookUp('SCHEDULED JOBS', ID = varDeleteRecord.ScheduledJobID),
        { JobStatus: "Open", SubmissionID: Blank() }
    )
);
```

---

## 5. Test checklist

- [ ] Save the sample cover sheet with a bad B4 date: the flow emails the error and imports nothing.
- [ ] Fix B4 and re-save: the jobs appear in SCHEDULED JOBS; re-saving again creates no duplicates.
- [ ] Pick a job in the app: all eight fields prefill; hourly rates match the pay split.
- [ ] Submit: SUBMISSIONS has ScheduledJobID; the SCHEDULED JOBS row is `Submitted` and gone from the list.
- [ ] Job Not Listed: the blank form works exactly as before.
- [ ] Edit an existing submission: values come from the saved record, not the schedule.
- [ ] Delete a submission: its scheduled job re-opens.
- [ ] Daily Cover Sheet report: output unchanged.

## Later: skip the spreadsheet step

The four lookup columns come from Natasha Coffman's Production Schedule on the `P:` drive. If that
schedule is moved to SharePoint/OneDrive (or exported there daily), the flow can build SCHEDULED JOBS
straight from it. The office would then only assign crews and phase, and nobody would retype
builder, subdivision or lot.
