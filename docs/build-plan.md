# Finish Crew Timecard — Cover Sheet Prefill & Pay Math Build Plan

Work through the phases in order. Each phase can be tested before starting the next.
Every formula below uses the control and column names from the current app export.

| Phase | What | Where |
|---|---|---|
| 1 | SCHEDULED JOBS list, SUBMISSIONS column changes | SharePoint |
| 2 | Import script | Excel for the web (Office Scripts) |
| 3 | Import flow | Power Automate |
| 4 | Restore pay splits and fix the pay math | Power Apps (scrJobInfo, scrPayItems, scrTimeCardDetail) |
| 5 | Builder as text | Power Apps and existing flows |
| 6 | Job picker and prefill | Power Apps |
| 7 | Test | — |

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
| ScheduledJobID | Number, 0 decimals | Links a time card to its scheduled job |

The existing **Builder** column is a **Lookup** column, not a Choice column (the app export shows a
`Builder Id` field). SharePoint can't convert a lookup to text, so add the new text column and leave
the old one in place for existing records. Don't delete it; the app and flows fall back to it for
old time cards.

SUBMISSIONS also contains **Import Status**, **Source File**, **Import Key** and
**Assigned/Claimed by Email** columns that nothing in the app uses. They look like an earlier
import idea. Leave them alone; this plan doesn't use them.

### 1.3 Import folder

In the site's **Documents** library, create two folders side by side (not one inside the other):
**Cover Sheet Import** (the office saves here) and **Cover Sheet Imported** (the flow moves
processed files here).

---

## Phase 2 — Office Script

1. Open any workbook in Excel for the web → **Automate** → **New Script**.
2. Replace the contents with `office-scripts/Import Finish Cover Sheet.ts` from this repo.
3. Name it **Import Finish Cover Sheet** and save it. It saves to the OneDrive of whoever will own
   the flow, so do this signed in as that account.

What it does:
- Reads the **FINISHCOVER** table by header name, so reordering columns doesn't break it.
- Reads the job date from **B4** and rejects invalid dates. The template you sent contains
  `9/25/25026`, which this catches.
- Skips unused rows, collapses wrapped text, and adds the **N** prefix when it's missing (the same
  rule as the app).
- Treats `Not found` or `#REF!` from the Production Schedule lookups as blank and returns a warning
  for that row.
- If a job number appears twice, imports the first row and warns about the second.
- Ignores all four comment columns.

Template change: select **B4** → Data → Data Validation → Allow **Date**, between 1/1/2025 and
12/31/2030. This stops the date typo at the source.

About the P: drive: Supervisor, Builder, Subdivision and Lot # are XLOOKUPs into the Production
Schedule on the P: drive. The cloud flow can't reach P:, so the script reads the values Excel
saved into the file. The office should keep filling the sheet in **desktop Excel** (where the
lookups calculate) and then save it into the import folder.

---

## Phase 3 — Power Automate import flow

Create an **Automated cloud flow** named **Import Finish Cover Sheet**.

1. **Trigger:** SharePoint → *When a file is created (properties only)*
   - Site Address: IndyWarrantyTracking-NRG365
   - Library Name: Documents
   - Folder: `/Shared Documents/Cover Sheet Import`
   - Trigger condition (Settings → Trigger conditions), so only Excel files start the flow:
     `@endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.xlsx')`
   - The trigger supplies the file **Identifier** and the creator's email
     (`triggerOutputs()?['body/Author/Email']`) used below.
2. **Run script** (Excel Online (Business))
   - Location: the SharePoint site, Document Library: Documents
   - File: the trigger's **Identifier**
   - Script: **Import Finish Cover Sheet**
3. **Apply to each** over `outputs('Run_script')?['body/result/rows']`
   1. **Get items** (SCHEDULED JOBS)
      - Filter Query: type `ImportKey eq '` then click **fx** and insert the expression
        `items('Apply_to_each')?['importKey']`, then type a closing `'`. The box should read
        `ImportKey eq '` [fx items(...)] `'`. Don't paste `@{...}` into the box; the new designer
        rejects it as an invalid expression.
      - Top Count: `1`
   2. **Condition** (directly under Get items, still inside Apply to each). **One row only:**
      left side **fx** `length(outputs('Get_items')?['body/value'])`, *is equal to*, right side `0`.
      - **If yes → Create item** (SCHEDULED JOBS):

        | Field | Value |
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
        | JobStatus | `Open` |
        | ImportKey | `items('Apply_to_each')?['importKey']` |

      - **If no → Condition:** `first(outputs('Get_items')?['body/value'])?['JobStatus']` *is equal
        to* `Open`
        - **If yes → Update item**. Id: `first(outputs('Get_items')?['body/value'])?['ID']`, with
          the same fields. A corrected re-save updates jobs nobody has submitted yet.
        - **If no →** do nothing. Submitted jobs are never overwritten.

   **Designer tip:** enter every value in steps 3.1–3.2 with **fx** expressions, as shown. If you
   pick a Get items field from the dynamic-content list instead, the designer silently wraps the
   step in an extra **For each**. That makes the loop run once per matching item, which is zero
   times for a new job, so nothing is ever created.
4. **Move file** (SharePoint). File to Move: the trigger's **Identifier**, Destination Folder:
   `/Shared Documents/Cover Sheet Imported`, If another file is already there: **Replace**.
   This keeps the import folder empty.
5. **Send an email (V2)**, placed below Apply to each. Enter each value with **fx** (don't paste
   `@{...}` text):
   - **To:** click in the box → **fx** → `triggerOutputs()?['body/Author/Email']` → **Add**.
     This is the email of whoever saved the cover sheet. The designer calls it *Created By Email*,
     and it's often hidden under **See more** in the trigger's dynamic content.
   - **Subject:** type `Cover Sheet imported – ` then **fx** `outputs('Run_script')?['body/result/jobDate']`
   - **Body:** **fx** `length(outputs('Run_script')?['body/result/rows'])`, then type
     ` jobs imported.`, press Enter, then **fx**
     `join(outputs('Run_script')?['body/result/warnings'], '<br>')`
6. **Error email.** On the arrow between **Run script** and **Apply to each**, click **+** → **Add a
   parallel branch** → **Send an email (V2)**. Then on that email's **Settings** tab → **Run
   after**, expand **Run script**, uncheck **Is successful** and check **Has failed**.
   - **To:** **fx** `triggerOutputs()?['body/Author/Email']`
   - **Subject:** `Cover Sheet NOT imported`
   - **Body:** type `The cover sheet could not be imported: ` then **fx**
     `coalesce(outputs('Run_script')?['body/error/message'], outputs('Run_script')?['body/message'], 'See the flow run history for details.')`

   The sender then learns about a bad date or the wrong template right away.

**Test:** save the sample cover sheet into the folder. With the B4 typo you should get the error
email. Fix B4, save again, and the two jobs appear in SCHEDULED JOBS. Save a third time and there
are still only two.

---

## Phase 4 — Restore pay splits and fix the pay math

### 4.1 App.OnStart: add at the end

```
ClearCollect(colPaySplits, 'FINISH INSTALLER PAY SPLITS');
Set(varSchedJob, Blank());
Set(varJobInfoLoaded, false);
```
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

**btnToLineItems.OnSelect** (Submit Time Card): replace the three `Set(varSelectedItemsTotal…)`,
`Set(varLeadInstallerTotalPay…)` and `Set(varHelperInstallerTotalPay…)` blocks with:
```
Set(varSelectedItemsTotal, Round(Sum(colLineItems, 'Extended Total'), 2));

Set(
    varLeadInstallerTotalPay,
    Round(Coalesce(varLeadShopPay, 0) + varSelectedItemsTotal * Coalesce(varLeadPercent, 0) / 100, 2)
);

Set(
    varHelperInstallerTotalPay,
    Round(Coalesce(varHelperShopPay, 0) + varSelectedItemsTotal * Coalesce(varHelperPercent, 0) / 100, 2)
);
```
In the same formula, add these two lines to the reset block at the end (next to
`Set(varLeadShopPay, 0);`). Otherwise the next time card starts with the previous card's rates:
```
Set(varLeadHourlyRate, 0);
Set(varHelperHourlyRate, 0);
```
(Phase 6 adds one more block to this button.)

### 4.7 scrTimeCardDetail

**btnEditPayItems.OnSelect**: replace the two percentage `Set`s and the `ClearCollect`:
```
Set(
    varLeadPercent,
    Coalesce(Value(Substitute(varSelectedSubmission.'Lead Installer % of Pay', "%", "")), 0)
);
Set(
    varHelperPercent,
    Coalesce(Value(Substitute(varSelectedSubmission.'Helper Installer % of Pay', "%", "")), 0)
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
```
The old version dropped **RequiresManualRate**. When a saved card was edited, manual-rate items
had their rate locked and quantity unlocked, the reverse of new cards. Leave the rest of the
button as is (edit mode, shop pay, Navigate).

**btnEditTimeCard.OnSelect**: add at the top:
```
Set(varJobInfoLoaded, false);
Set(varSchedJob, Blank());
```

**btnApproved** email: in the two `% of Pay` lines, change
`Value(varSelectedSubmission.'Lead Installer % of Pay')` to
`Value(Substitute(varSelectedSubmission.'Lead Installer % of Pay', "%", ""))`, and do the same for
Helper. Otherwise a percent typed as "60%" shows as 0.6% in the email.

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
| scrOpenSubmissions → lblBuilder | `"Builder: " & Coalesce(ThisItem.'Builder Name', ThisItem.Builder.Value)` |
| scrTimeCardDetail → lblBuilderName | `"Builder: " & Coalesce(varSelectedSubmission.'Builder Name', varSelectedSubmission.Builder.Value)` |
| scrTimeCardDetail → btnApproved email | replace `varSelectedSubmission.Builder.Value` with `Coalesce(varSelectedSubmission.'Builder Name', varSelectedSubmission.Builder.Value)` |

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

### 6.1 Data source
Power Apps → Data → Add data → SharePoint → IndyWarrantyTracking-NRG365 → **SCHEDULED JOBS**.

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
- A back arrow: `Navigate(scrHome, ScreenTransition.Fade)`

**galSchedJobs.Items** lists open jobs from the last 3 days, so a late time card still finds its
job:
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
    JobDate,
    SortOrder.Descending
)
```

Gallery labels:
- `ThisItem.Title & "   Lot " & ThisItem.LotNumber & "   " & Text(ThisItem.JobDate, "m/d")`
- `ThisItem.Builder & " – " & ThisItem.Subdivision`
- `ThisItem.LeadInstaller & If(IsBlank(ThisItem.HelperInstaller), "", " / " & ThisItem.HelperInstaller) & "   " & ThisItem.Phase`

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

**btnJobNotListed.OnSelect** (same as the old New Time Card):
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
        { ScheduledJobID: varSchedJob.ID }
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
    !IsBlank(varDeleteRecord.ScheduledJobID),
    Patch(
        'SCHEDULED JOBS',
        LookUp('SCHEDULED JOBS', ID = varDeleteRecord.ScheduledJobID),
        { JobStatus: "Open", SubmissionID: Blank() }
    )
);
```

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
- [ ] Pick a job → all eight fields plus the date prefill; schedule fields are locked; % and rates load.
- [ ] Swap the helper → split reloads for the new pair.
- [ ] Submit → the job leaves the picker; SUBMISSIONS.ScheduledJobID and SCHEDULED JOBS.SubmissionID are set.
- [ ] Job Not Listed → blank form, same as before.
- [ ] Delete that time card → the job reappears in the picker.
- [ ] Daily Cover Sheet report → unchanged output.

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
