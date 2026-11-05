# RNC Rough Walk (RQI) — Build Plan

*Started 2026-11-05. Draft 1: plan only, nothing built yet.*

Files in this folder:
- `docs/build-plan.md`: this plan
- `office-scripts/Fill Rough Walk Sheet.ts`: Office Script that fills the Excel template that
  becomes the PDF (Phase 3)
- `courier/Copy-RoughWalksToP.ps1`: script on the 24-hour server that copies finished walks
  from SharePoint into the job folders on P: (Phase 6)
- `courier/courier-config.sample.json`: settings for the courier script

Work through the phases in order. Each phase can be tested before starting the next.

---

## How it works

```
 iPad / phone (Power Apps)                SharePoint (cloud)                24-hour server
 ─────────────────────────                ──────────────────                ──────────────
 Pick job (prefilled from      ──────►    RQI WALKS list (the answers)
 RNC JOBS list)                           WalkDrop library:
 Fill in walk, take photos,                 2601001-57/
 sign                                         N2601001-RWC.pdf     ◄── flow builds the PDF,
 Submit (waits in the Outbox                  Rough Walk Photos/       emails the Field
 if there's no signal)                          N2601001-RW-01.jpg     Supervisor
                                                       │
                                                       ▼
                                          Courier script (every 10 min) ──► P:\Indianapolis\JOB FOLDERS\
                                                                              2026 Jobs\01 JANUARY\
                                                                                N2601001_ARCHER_169 COLD SPRINGS\
                                                                                  N2601001-RWC.pdf
                                                                                  Rough Walk Photos\
```

**P: stays the master copy.** The WalkDrop library is only a holding area: each walk sits there
until the courier has copied it into its job folder. Nobody has to change how they use P:. The
existing job folders do **not** move. Moving all job folders into SharePoint is a separate,
optional step (Phase 9).

**Finding the job folder.** The job number gives the path:

| Job # | Year folder | Month folder | Job folder |
|---|---|---|---|
| 2601001 | `2026 Jobs` | `01 JANUARY` | the one folder whose name starts `N2601001_` |
| 2610023 | `2026 Jobs` | `10 OCTOBER` | `N2610023_LENNAR_177_BEYERS ESTATES` |

The courier matches on the `N2601001_` prefix only, so the rest of the name (builder, lot,
subdivision, spaces or underscores) can be anything. If there's no matching folder, or more than
one, the walk waits and the office gets an email. When someone creates or fixes the folder, the
next run files the walk automatically.

**Only standard licenses.** Power Apps, SharePoint, Excel Online, OneDrive and Outlook connectors
are all included with Microsoft 365. Nothing here needs a Power Apps or Power Automate premium
license.

### Assumptions to confirm

| # | Assumption | Status |
|---|---|---|
| A1 | The RQI report is **now** emailed to the Field Supervisor for the builder (you wrote "not sent"; I read it as "now sent"). | Confirm |
| A2 | The PDF goes in the root of the job folder, named `N2601001-RWC.pdf`. | Confirmed (keep current names) |
| A3 | Photos go in a subfolder named **Rough Walk Photos**, named `N2601001-RW-01.jpg`, `-02`, … | Confirm the subfolder name |
| A4 | A re-walk of the same job never overwrites: it's saved as `N2601001-RWC-2.pdf`, and its photos continue the numbering. | Confirm |
| A5 | Punch Out (POC Name, Date) is filled in later, after the punch items are done. The app lets the assistant reopen a filed walk to add it, which re-sends an updated PDF. | Confirm who fills it in and when |
| A6 | The Field Supervisor forwards punch items to the Installer, and "100% Complete" jobs to Production, as today. Sending "100% Complete" straight to Production is an option (Phase 9). | Confirm |
| A7 | The `500` folder under JOB FOLDERS isn't used for RNC jobs. | Confirm |

| Phase | What | Where |
|---|---|---|
| 0 | Prerequisites: service account, server access, app registration | IT |
| 1 | Lists and the WalkDrop library | SharePoint |
| 2 | Production Schedule → RNC JOBS list | Server script or Power Automate |
| 3 | PDF template and Office Script | Excel for the web |
| 4 | The walk app (offline Outbox, photos, signature) | Power Apps |
| 5 | Flows: upload a file, walk submitted, alerts | Power Automate |
| 6 | Courier: copy finished walks to P: | 24-hour server (PowerShell) |
| 7 | Test | — |
| 8 | Rollout to the iPads and phones | — |
| 9 | Later: Production email, photo page in PDF, reporting, SharePoint job folders | — |

Before building all of Phase 4, do the **photo spike test (7.1)**: it proves the riskiest part
(photos saved offline, then uploaded) in about an hour.

---

## Phase 0 — Prerequisites (IT)

### 0.1 Service account

One licensed Microsoft 365 account, for example **svc-rqi@yourcompany.com**, that:
- owns the flows (Phase 5), so they don't stop working if an employee leaves
- holds the Excel template and the Office Script in its OneDrive (Phase 3)
- has Contribute access to the IndyWarrantyTracking-NRG365 site
- has "Send As" on a shared mailbox such as **RQI Reports** (emails come from there)

If you'd rather not add an account, the flows can run under your own account to start with and be
moved later. They all have to be owned by the same account.

### 0.2 Server access to P:

On the 24-hour server:
1. A Windows account for the courier (for example `DOMAIN\svc-rqi-courier`) with **Modify** on
   the JOB FOLDERS share.
2. The **UNC path** behind P: (for example `\\fileserver\Public\Indianapolis\JOB FOLDERS`).
   Scheduled tasks don't see mapped drive letters, so the script must use the UNC path.
3. PowerShell 7.4 or later and the PnP.PowerShell module:
   ```
   winget install Microsoft.PowerShell
   pwsh -Command "Install-Module PnP.PowerShell -Scope AllUsers"
   ```

### 0.3 App registration (lets the server read and update SharePoint)

A Microsoft 365 admin creates this once in the Entra admin center:
1. **App registrations → New registration**: name **RQI Courier**, single tenant.
2. **Certificates & secrets → Certificates**: upload a certificate. The certificate is installed
   on the server in the courier account's certificate store. Note the **thumbprint**.
3. **API permissions → SharePoint → Application permissions → Sites.Selected**, then **Grant
   admin consent**. Sites.Selected gives the app no access at all until it's granted a site.
4. Grant the app **write** to this one site only (run once, as a SharePoint admin):
   ```
   Connect-PnPOnline -Url https://yourtenant.sharepoint.com/sites/IndyWarrantyTracking-NRG365 -Interactive -ClientId <an admin app id>
   Grant-PnPAzureADAppSitePermission -AppId <RQI Courier app id> -DisplayName "RQI Courier" -Permissions Write -Site https://yourtenant.sharepoint.com/sites/IndyWarrantyTracking-NRG365
   ```
5. Give me the **tenant name**, **app (client) ID** and **certificate thumbprint** for the courier
   config (Phase 6). None of these are passwords, but keep the certificate's private key on the
   server only.

**No-admin fallback:** sign the courier account into the OneDrive sync app on the server and sync
the WalkDrop library, then point the script at the synced folder. It works, but the sync app only
runs while that account is signed in, so the app registration is the reliable way.

---

## Phase 1 — SharePoint

Use the IndyWarrantyTracking-NRG365 site (or a new site, if you'd rather keep RQI separate).

**Column rule:** create every column with the **name with no spaces** shown below, then rename it
to the display name. The internal name stays the short one, which keeps flow expressions and the
courier simple.

### 1.1 List: RNC JOBS (filled from the Production Schedule)

| Column (create as) | Display name | Type | Notes |
|---|---|---|---|
| Title | Job Number | Single line | 7 digits, no N (e.g. `2601001`). Index it. |
| Builder | Builder | Single line | |
| Lot | Lot | Single line | Text, because lots like `A161` exist |
| Subdivision | Subdivision | Single line | |
| StreetAddress | Street Address | Single line | |
| FieldSupervisor | Field Supervisor | Single line | From the schedule if it has one |
| InstallerName | Installer Name | Single line | |
| InstallerNumber | Installer # | Single line | |
| RoughDate | Rough-in Date | Date | Index it. Used to pick which jobs go offline. |

Exact columns depend on the Production Schedule (Phase 2).

### 1.2 List: BUILDER SUPERVISORS

One row per builder: which of the six Field Supervisors gets the RQI report.

| Column | Display name | Type |
|---|---|---|
| Title | Builder | Single line (matches the Builder text in RNC JOBS, e.g. `PULTE`) |
| SupervisorName | Field Supervisor | Single line |
| SupervisorEmail | Supervisor Email | Single line |
| CcEmails | Cc | Single line, optional (`a@x.com;b@x.com`) |

Plain text, not a Person column, so the app can keep a copy offline.

### 1.3 List: RQI WALKS (one row per walk)

| Column | Display name | Type | Paper form |
|---|---|---|---|
| Title | Job Number | Single line | Job # |
| Builder, Lot, Subdivision | (same) | Single line | Home Information |
| InspectorName | Inspector Name | Single line | RQI Details |
| InspectorEmail | Inspector Email | Single line | |
| InspectionDate | Inspection Date | Date | RQI Details → Date |
| InstallerName | Installer Name | Single line | |
| InstallerNumber | Installer # | Single line | |
| FieldSupervisor | Field Supervisor | Single line | |
| SupervisorEmail | Supervisor Email | Single line | |
| NotReady | Not Ready for Walk | Yes/No | |
| NIF | NIF (Not Installer's Fault) | Yes/No | |
| Notes | Notes | Multiple lines (plain) | |
| Floor2 | 2nd Floor | Choice: Yes, No, N/A | Home Information |
| Floor1 | 1st Floor | Choice: Yes, No, N/A | |
| Basement | Basement | Choice: Yes, No, N/A | |
| HomeComments | Home Comments | Multiple lines (plain) | Comments |
| HomeMaterials | Home Materials Needed | Multiple lines (plain) | Materials Needed |
| ExteriorOfHome | Exterior of Home | Choice: Yes, No, N/A | Equipment Exterior |
| FurnaceSet | Furnace Set | Choice: Yes, No, N/A | |
| Zoning | Zoning | Choice: Yes, No, N/A | |
| MiscLabel | Misc Item | Single line | Misc. ____ |
| Misc | Misc | Choice: Yes, No, N/A | |
| EquipComments | Equipment Comments | Multiple lines (plain) | Comments |
| EquipMaterials | Equipment Materials Needed | Multiple lines (plain) | Materials Needed |
| Complete100 | 100% Complete | Yes/No | Written as "100%" in Comments today |
| POCName | POC Name | Single line | Punch Out |
| PunchOutDate | Punch Out Date | Date | Punch Out |
| SignedDate | Signed Date | Date | Inspector Signature → Date |
| RQIReportSent | RQI Report Sent | Yes/No | RQI Report (set by the flow) |
| RQISentOn | RQI Sent On | Date and time | RQI Report Date / Time (set by the flow) |
| PhotoCount | Photo Count | Number | |
| DropFolder | Drop Folder | Single line | e.g. `WalkDrop/2601001-57` (set by the app) |
| WalkStatus | Status | Choice (below) | |
| PdfName | PDF Name | Single line | set by the courier |
| FiledOn | Filed On | Date and time | set by the courier |
| FiledPath | Filed To | Single line | set by the courier |
| FilingError | Filing Problem | Multiple lines (plain) | set by the flow or courier |
| LocalId | Local Id | Single line | Unique id made on the iPad. Index it. Stops a weak-signal retry from saving the walk twice. |

**WalkStatus choices**, in order: `Uploading`, `Submitted`, `Processing`, `Ready to File`,
`Filed`, `Folder Not Found`, `Error`.

| Status | Set by | Meaning |
|---|---|---|
| Uploading | app | Row saved, photos still uploading |
| Submitted | app | Everything uploaded; starts the flow |
| Processing | flow | PDF being built |
| Ready to File | flow | PDF made, supervisor emailed, waiting for the courier |
| Filed | courier | Copied into the P: job folder |
| Folder Not Found | courier | No folder (or more than one) starts with `N<job#>_`. Retried every run. |
| Error | flow or courier | See Filing Problem |

Signatures aren't stored in the list. A signature image is too big for a text column, so it's
uploaded to the drop folder as `_work/signature.png`.

### 1.4 Library: WalkDrop

1. **New → Document library**, name it `WalkDrop` (no space, so the URL is `/WalkDrop`), then
   rename it to **WALK DROP** if you like.
2. Field assistants don't need any access to it. Files are uploaded by a flow running as the
   service account.
3. Layout, made automatically:
   ```
   WalkDrop/
     2601001-57/                 ← <Job #>-<RQI WALKS ID>, so re-walks never collide
       N2601001-RWC.pdf
       Rough Walk Photos/
         N2601001-RW-01.jpg
       _work/
         signature.png           ← used for the PDF only; the courier skips _work
   ```

### 1.5 Permissions

| Who | RNC JOBS | BUILDER SUPERVISORS | RQI WALKS | WalkDrop |
|---|---|---|---|---|
| Field assistants | Read | Read | Contribute | none |
| Field Supervisors, office | Read | Read | Read (optional) | Read (optional) |
| Service account, courier app | Edit | Edit | Edit | Edit |

---

## Phase 2 — Production Schedule → RNC JOBS

The Production Schedule has everything the walk prefills: Builder, Lot, Subdivision, Field
Supervisor, Installer. It lives on P:, which the cloud can't reach, but the 24-hour server can.

**Recommended:** a second scheduled script on the same server reads the Production Schedule every
night (and maybe at noon), and adds or updates RNC JOBS rows by Job Number. Nobody has to export
anything.

**Alternative:** the same pattern as the Finish Crew import: someone saves the schedule (or an
export) to a SharePoint folder, and a flow with an Office Script upserts the rows.

**Needed from you to write this phase:**
- a copy of the Production Schedule (or a few rows of it with the headers)
- its path on P:
- which column is the rough-in date, so the app can take only recent jobs offline

Until Phase 2 is built, RNC JOBS can be filled by hand: **Microsoft Lists → Import from Excel**
from a copy of the schedule. That's enough to build and test the app.

---

## Phase 3 — PDF template and Office Script

The PDF is made from an Excel template that looks like the paper form. Excel to PDF conversion is
included with Microsoft 365. Word templates need a premium license, so they're not used.

### 3.1 Build the template

Signed in as the service account:
1. In OneDrive, make a folder **RQI** with two subfolders: **Template** and **Temp**.
2. Make a workbook **RQI Rough Walk Template.xlsx** in **RQI/Template** with one sheet, **RWC**.
   Lay it out like the paper form (both halves on one landscape page, or portrait if that reads
   better). Use borders and grey fills for the headings. The look is up to you.
3. Name each answer cell (select the cell → type the name in the **Name Box** left of the formula
   bar → Enter). The script fills cells by these names, so the layout can change without changing
   any code.

| Name | Cells | Filled with |
|---|---|---|
| Builder, Lot, Subdivision, JobNumber | 1 each | text (JobNumber with the N, e.g. `N2601001`) |
| InspectorName, InspectionDate, InstallerName, InstallerNumber, FieldSupervisor | 1 each | text |
| RQIReportSent | 1 | `X` |
| RQIDate, RQITime | 1 each | date and time the report was emailed |
| POCName, PunchOutDate | 1 each | text, often blank on the first PDF |
| SignedDate | 1 | date |
| Signature | a block of merged cells, e.g. 4 × 2 | the signature image, scaled to fit |
| NotReady, NIF | 1 each | `X` or blank (put a box border around these cells) |
| Notes | a merged block, wrap text on | text |
| Floor2, Floor1, Basement | **3 cells in a row**: Yes, No, N/A | `X` in the matching cell |
| ExteriorOfHome, FurnaceSet, Zoning, Misc | **3 cells in a row** | `X` in the matching cell |
| MiscLabel | 1 | text |
| HomeComments, EquipComments | merged block, wrap on | text (`100% Complete` when ticked) |
| HomeMaterials, EquipMaterials | merged block, wrap on | text, one item per line |

4. **Page Layout:** set the print area to the form, **Fit to 1 page wide by 1 tall**, and pick
   portrait or landscape.
5. The Excel → PDF conversion uses the print settings, so check by doing **File → Print → Print
   preview**.

### 3.2 Add the script

1. Open the template in Excel for the web → **Automate → New Script**.
2. Replace everything with `office-scripts/Fill Rough Walk Sheet.ts` (open it in Notepad, copy
   all, paste into an empty editor).
3. Name it **Fill Rough Walk Sheet** and save. It saves to the service account's OneDrive.

What it does:
- `dataJson` holds one value for each cell name above. Each value goes into the cell with that
  name. Values are written as text, so `100%` and dates stay exactly as sent.
- For a 3-cell name (Yes / No / N/A), it puts `X` under the answer and clears the other two.
- Places the signature image in the **Signature** block, scaled to fit.
- Returns the names it couldn't find, so a missing name shows up in the flow's error email instead
  of silently leaving a blank.

---

## Phase 4 — Power Apps: RQI Rough Walk app

### 4.1 App settings

- **Create → Blank canvas app → Tablet format.** All screens are built from vertical containers
  with **Flexible height** sections, so one app fits an iPad and a phone.
- **Settings → Display:** Scale to fit **Off**, Lock aspect ratio **Off**, Lock orientation
  **Off**.
- **Data:** RNC JOBS, BUILDER SUPERVISORS, RQI WALKS, and the flow **RQI Upload File**
  (Phase 5.1, which you build before the Submit button).
- Assistants run it in the **Power Apps** app from the App Store, not in Safari. Offline storage
  (SaveData / LoadData) is meant for the mobile app.

### 4.2 Screens

| Screen | What it shows |
|---|---|
| scrHome | **Start a walk**, **My drafts** (on this iPad), **Outbox** (waiting for signal), **My recent walks** (from RQI WALKS, with status) |
| scrPickJob | Search by job #, builder, subdivision or lot. Offline it searches the saved copy. |
| scrWalk | The form: Home Information (read-only, from the job), RQI Details, Not Ready / NIF / Notes, Home Information walk, Equipment Exterior, Comments, Materials |
| scrPhotos | **Take photo** (opens the iPad camera at full resolution), thumbnails, delete |
| scrReview | Summary, signature pad, **Submit** |

### 4.3 App.OnStart

```
Set(gblMe, User());
Concurrent(
    LoadData(colOutbox, "rqiOutbox", true),
    LoadData(colDrafts, "rqiDrafts", true)
);
If(
    Connection.Connected,
    // Online: refresh the offline copies
    ClearCollect(
        colJobs,
        Filter('RNC JOBS', RoughDate >= DateAdd(Today(), -90, TimeUnit.Days))
    );
    ClearCollect(colSupervisors, 'BUILDER SUPERVISORS');
    SaveData(colJobs, "rqiJobs");
    SaveData(colSupervisors, "rqiSupervisors"),
    // Offline: use the last copies saved on this iPad
    LoadData(colJobs, "rqiJobs", true);
    LoadData(colSupervisors, "rqiSupervisors", true)
);
```

`colJobs` holds up to 2,000 jobs (the app's limit). 90 days of rough-ins is well under that. If
it isn't, shorten the window.

### 4.4 scrPickJob

- `txtJobSearch` (text input), `galJobs` (vertical gallery):
  ```
  galJobs.Items =
  SortByColumns(
      Filter(
          colJobs,
          IsBlank(txtJobSearch.Value) ||
          StartsWith(Title, Substitute(Upper(txtJobSearch.Value), "N", "")) ||
          txtJobSearch.Value in Builder ||
          txtJobSearch.Value in Subdivision ||
          txtJobSearch.Value in Lot
      ),
      "Title", SortOrder.Descending
  )
  ```
  Typing `N2601001` or `2601001` both work.
- Gallery row: `"N" & ThisItem.Title`, `ThisItem.Builder & " · Lot " & ThisItem.Lot & " · " & ThisItem.Subdivision`.
- **Job not in the list** button: lets the assistant type Job #, Builder, Lot and Subdivision by
  hand (for a job the schedule missed).
- `galJobs.OnSelect` starts a new walk:
  ```
  Set(gblJob, ThisItem);
  Set(gblWalkLocalId, GUID());
  Set(gblSup, LookUp(colSupervisors, Upper(Title) = Upper(ThisItem.Builder)));
  Clear(colPhotos);
  Navigate(scrWalk)
  ```

### 4.5 scrWalk controls

| Section | Controls |
|---|---|
| Home Information | Labels from `gblJob` (Builder, Lot, Subdivision, Job #) |
| RQI Details | `txtInspector` Default `gblMe.FullName`; `dteInspection` Default `Today()`; `txtInstaller` Default `gblJob.InstallerName`; `txtInstallerNo` Default `gblJob.InstallerNumber`; `txtSupervisor` Default `Coalesce(gblJob.FieldSupervisor, gblSup.SupervisorName)` |
| Status | `tglNotReady`, `tglNIF`, `txtNotes` (multi-line) |
| Home Information walk | `rdoFloor2`, `rdoFloor1`, `rdoBasement`: horizontal Radio, Items `["Yes","No","N/A"]`, large touch targets |
| Comments | `txtHomeComments`, `txtHomeMaterials` |
| Equipment Exterior | `rdoExterior`, `rdoFurnace`, `rdoZoning`, `txtMiscLabel`, `rdoMisc` |
| Comments | `tglComplete100`, `txtEquipComments`, `txtEquipMaterials` |

- `tglComplete100.OnCheck`: `Reset(txtEquipComments)`, with `txtEquipComments.Default =
  If(tglComplete100.Checked, "100% Complete", "")`.
- When **Not Ready for Walk** is on, the walk questions are optional. Otherwise Submit stays
  disabled until every Yes/No/N/A has an answer (`DisplayMode` on `btnSubmit`, 4.7).
- **Save draft** (top bar) keeps the walk on this iPad so it can be finished later. It works the
  same way as the Outbox (4.7), with collection `colDrafts`.

### 4.6 scrPhotos

- `addPhoto` (Add picture control). On an iPad and a phone it offers **Take Photo** or **Photo
  Library**, at full camera resolution.
- `addPhoto.OnChange`:
  ```
  Collect(
      colPhotos,
      {
          Seq: CountRows(colPhotos) + 1,
          Img: addPhoto.Media
      }
  );
  Reset(addPhoto)
  ```
- `galPhotos.Items = colPhotos` (thumbnails), with a trash icon: `Remove(colPhotos, ThisItem)`.
- Label: `CountRows(colPhotos) & " photos"`.

### 4.7 scrReview: sign and Submit

- `penSig` (Pen input) for the inspector's signature, `dteSigned` Default `Today()`.
- `btnSubmit.DisplayMode`:
  ```
  If(
      !IsBlank(penSig.Image) &&
      (tglNotReady.Checked ||
          (!IsBlank(rdoFloor2.Selected) && !IsBlank(rdoFloor1.Selected) &&
           !IsBlank(rdoBasement.Selected) && !IsBlank(rdoExterior.Selected) &&
           !IsBlank(rdoFurnace.Selected) && !IsBlank(rdoZoning.Selected))),
      DisplayMode.Edit,
      DisplayMode.Disabled
  )
  ```
- `btnSubmit.OnSelect`: everything goes into the **Outbox** first, then sends if there's signal.
  Nothing is lost if the signal drops halfway.
  ```
  Collect(
      colOutbox,
      {
          LocalId: gblWalkLocalId,
          JobNumber: gblJob.Title,
          Rec: {
              Title: gblJob.Title,
              Builder: gblJob.Builder,
              Lot: gblJob.Lot,
              Subdivision: gblJob.Subdivision,
              InspectorName: txtInspector.Value,
              InspectorEmail: gblMe.Email,
              InspectionDate: dteInspection.SelectedDate,
              InstallerName: txtInstaller.Value,
              InstallerNumber: txtInstallerNo.Value,
              FieldSupervisor: txtSupervisor.Value,
              SupervisorEmail: gblSup.SupervisorEmail,
              NotReady: tglNotReady.Checked,
              NIF: tglNIF.Checked,
              Notes: txtNotes.Value,
              Floor2: {Value: rdoFloor2.Selected.Value},
              Floor1: {Value: rdoFloor1.Selected.Value},
              Basement: {Value: rdoBasement.Selected.Value},
              HomeComments: txtHomeComments.Value,
              HomeMaterials: txtHomeMaterials.Value,
              ExteriorOfHome: {Value: rdoExterior.Selected.Value},
              FurnaceSet: {Value: rdoFurnace.Selected.Value},
              Zoning: {Value: rdoZoning.Selected.Value},
              MiscLabel: txtMiscLabel.Value,
              Misc: {Value: rdoMisc.Selected.Value},
              Complete100: tglComplete100.Checked,
              EquipComments: txtEquipComments.Value,
              EquipMaterials: txtEquipMaterials.Value,
              SignedDate: dteSigned.SelectedDate,
              PhotoCount: CountRows(colPhotos),
              LocalId: gblWalkLocalId
          },
          Photos: colPhotos,
          Sig: penSig.Image
      }
  );
  SaveData(colOutbox, "rqiOutbox");
  If(Connection.Connected, Select(btnSendOutbox));
  Notify(
      If(Connection.Connected, "Walk sent.", "No signal. The walk is in the Outbox and will send when you're back online."),
      NotificationType.Success
  );
  Navigate(scrHome)
  ```
  (Control property names such as `.Value` and `.Media` are for the current modern controls.
  With classic controls they're `.Text` and `.Image`.)

### 4.8 scrHome: send the Outbox

Hidden button `btnSendOutbox`, pressed by Submit, by a **Send now** button next to the Outbox
count, and by a timer every 60 seconds while the Outbox isn't empty:
```
ClearCollect(colSending, colOutbox);
ForAll(
    colSending As o,
    With(
        {
            existing: LookUp('RQI WALKS', LocalId = o.LocalId)
        },
        With(
            {
                row: Coalesce(
                    existing,
                    Patch('RQI WALKS', Defaults('RQI WALKS'), o.Rec, {WalkStatus: {Value: "Uploading"}})
                )
            },
            With(
                {folder: "WalkDrop/" & row.Title & "-" & row.ID},
                ForAll(
                    o.Photos As p,
                    'RQIUploadFile'.Run(
                        folder & "/Rough Walk Photos",
                        "N" & row.Title & "-RW-" & Text(p.Seq, "00") & ".jpg",
                        {file: {contentBytes: p.Img, name: "photo.jpg"}}
                    )
                );
                'RQIUploadFile'.Run(
                    folder & "/_work",
                    "signature.png",
                    {file: {contentBytes: o.Sig, name: "signature.png"}}
                );
                Patch('RQI WALKS', row, {DropFolder: folder, WalkStatus: {Value: "Submitted"}});
                Remove(colOutbox, LookUp(colOutbox, LocalId = o.LocalId))
            )
        )
    )
);
SaveData(colOutbox, "rqiOutbox")
```

- If the signal drops partway, that walk stays in the Outbox. Next time it finds its row by
  **LocalId** (no duplicate) and uploads the files again under the same names, overwriting the
  partial copies.
- The `Timer`: Duration `60000`, Repeat on, AutoStart `CountRows(colOutbox) > 0`, OnTimerEnd
  `If(Connection.Connected && CountRows(colOutbox) > 0, Select(btnSendOutbox))`.
- Re-photo numbering starts at 01 for every walk. The courier renumbers if the job folder already
  has photos from an earlier walk.

### 4.9 Reopen a walk to add Punch Out (A5)

`My recent walks` gallery → **Add Punch Out** → POC Name and Date → Patch the row and set
`WalkStatus` back to **Submitted**. The flow rebuilds the PDF and emails it. The courier files it
as `N2601001-RWC-2.pdf` (A4). To overwrite the first PDF instead, the courier needs a different
rule. Tell me which you want.

---

## Phase 5 — Power Automate (all owned by the service account)

### 5.1 Instant flow: RQIUploadFile

Used by the app to upload each photo and the signature.

1. **Trigger:** Power Apps (V2). Inputs:
   - Text **FolderPath** (e.g. `WalkDrop/2601001-57/Rough Walk Photos`)
   - Text **FileName**
   - File **file**
2. **Create file** (SharePoint): Site IndyWarrantyTracking-NRG365, Folder Path
   **fx** `concat('/', triggerBody()?['text'])` (the FolderPath input), File Name the FileName
   input, File Content **fx** `triggerBody()?['file']?['contentBytes']`.
   Create file makes any missing folders and replaces a file with the same name.
3. **Respond to a PowerApp or flow:** Text **ok** = `yes`.

The flow uses the service account's SharePoint connection, so assistants don't need access to
WalkDrop. When you add the flow to the app, check that the run-only users setting for SharePoint
is **Use this connection**.

### 5.2 Automated flow: RQI Walk Submitted

```
Trigger (Status = Submitted) → Status Processing → Get signature → Copy template → Run script
  → Convert to PDF → Create PDF in drop folder → Delete temp → Email supervisor
  → Update item (Ready to File, RQI sent)
     └─ any failure → Update item (Error) → email office
```

1. **Trigger:** SharePoint → *When an item is created or modified*, list RQI WALKS. Trigger
   condition:
   `@equals(triggerOutputs()?['body/WalkStatus/Value'], 'Submitted')`
2. **Update item:** WalkStatus `Processing`. The flow changes the status straight away, so its
   own later updates don't start it again.
3. **Get file content using path** (SharePoint):
   **fx** `concat('/', triggerOutputs()?['body/DropFolder'], '/_work/signature.png')`
4. **Copy file** (OneDrive for Business): `/RQI/Template/RQI Rough Walk Template.xlsx` →
   `/RQI/Temp`, **Copy with a new name**.
5. **Compose Walk Data.** Type it as a JSON object. Each value is a separate expression, so
   quotes or line breaks in comments can't break it:
   ```
   {
     "JobNumber": "N@{triggerOutputs()?['body/Title']}",
     "Builder": "@{triggerOutputs()?['body/Builder']}",
     "Lot": "@{triggerOutputs()?['body/Lot']}",
     "Subdivision": "@{triggerOutputs()?['body/Subdivision']}",
     "InspectorName": "@{triggerOutputs()?['body/InspectorName']}",
     "InspectionDate": "@{formatDateTime(triggerOutputs()?['body/InspectionDate'], 'M-d-yy')}",
     "InstallerName": "@{triggerOutputs()?['body/InstallerName']}",
     "InstallerNumber": "@{triggerOutputs()?['body/InstallerNumber']}",
     "FieldSupervisor": "@{triggerOutputs()?['body/FieldSupervisor']}",
     "RQIReportSent": "X",
     "RQIDate": "@{formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'M-d-yy')}",
     "RQITime": "@{formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'h:mm tt')}",
     "POCName": "@{coalesce(triggerOutputs()?['body/POCName'], '')}",
     "PunchOutDate": "@{if(empty(triggerOutputs()?['body/PunchOutDate']), '', formatDateTime(triggerOutputs()?['body/PunchOutDate'], 'M-d-yy'))}",
     "SignedDate": "@{formatDateTime(triggerOutputs()?['body/SignedDate'], 'M-d-yy')}",
     "NotReady": "@{if(triggerOutputs()?['body/NotReady'], 'X', '')}",
     "NIF": "@{if(triggerOutputs()?['body/NIF'], 'X', '')}",
     "Notes": "@{coalesce(triggerOutputs()?['body/Notes'], '')}",
     "Floor2": "@{triggerOutputs()?['body/Floor2/Value']}",
     "Floor1": "@{triggerOutputs()?['body/Floor1/Value']}",
     "Basement": "@{triggerOutputs()?['body/Basement/Value']}",
     "HomeComments": "@{coalesce(triggerOutputs()?['body/HomeComments'], '')}",
     "HomeMaterials": "@{coalesce(triggerOutputs()?['body/HomeMaterials'], '')}",
     "ExteriorOfHome": "@{triggerOutputs()?['body/ExteriorOfHome/Value']}",
     "FurnaceSet": "@{triggerOutputs()?['body/FurnaceSet/Value']}",
     "Zoning": "@{triggerOutputs()?['body/Zoning/Value']}",
     "MiscLabel": "@{coalesce(triggerOutputs()?['body/MiscLabel'], '')}",
     "Misc": "@{triggerOutputs()?['body/Misc/Value']}",
     "EquipComments": "@{coalesce(triggerOutputs()?['body/EquipComments'], '')}",
     "EquipMaterials": "@{coalesce(triggerOutputs()?['body/EquipMaterials'], '')}"
   }
   ```
   Indianapolis is on Eastern time, and `'Eastern Standard Time'` handles daylight saving.
6. **Run script** (Excel Online (Business)): Location OneDrive for Business, File
   **fx** `outputs('Copy_file')?['body/Id']`, Script **Fill Rough Walk Sheet**:
   - dataJson: **fx** `string(outputs('Compose_Walk_Data'))`
   - signatureBase64: **fx** `body('Get_file_content_using_path')?['$content']`
7. **Condition:** **fx** `outputs('Run_script')?['body/result']` *is equal to* (blank). If it's
   not blank, a cell name is missing from the template. Put it in FilingError and carry on (the
   PDF is still useful).
8. **Convert file** (OneDrive for Business): File **fx** `outputs('Copy_file')?['body/Id']`,
   Target type **PDF**.
9. **Create file** (SharePoint): Folder **fx** `concat('/', triggerOutputs()?['body/DropFolder'])`,
   Name **fx** `concat('N', triggerOutputs()?['body/Title'], '-RWC.pdf')`, Content the Convert
   file output.
10. **Delete file** (OneDrive for Business): the temp copy.
11. **Get item** (BUILDER SUPERVISORS) where Title = Builder, for the email address and Cc. If the
    walk already has SupervisorEmail, use that.
12. **Send an email from a shared mailbox (V2):** from **RQI Reports**:
    - To: supervisor email. Cc: the list's Cc, plus the inspector.
    - Subject, one **fx**:
      ```
      concat('RQI ', if(triggerOutputs()?['body/NotReady'], 'NOT READY', if(triggerOutputs()?['body/Complete100'], '100% COMPLETE', 'PUNCH ITEMS')), ' – N', triggerOutputs()?['body/Title'], ' – ', triggerOutputs()?['body/Builder'], ' Lot ', triggerOutputs()?['body/Lot'], ' ', triggerOutputs()?['body/Subdivision'])
      ```
    - Body: the Comments and Materials Needed (so the supervisor can forward them to the
      installer without opening the PDF), NIF if ticked, the photo count, and the job folder
      path.
    - Attachment: `N2601001-RWC.pdf`. Photos aren't attached (too big for 20+ photos). They're in
      the job folder after the next courier run.
13. **Update item:** WalkStatus `Ready to File`, RQIReportSent Yes, RQISentOn `utcNow()`.
14. **Error branch.** Use a **Scope** around steps 3–13. A parallel step after it, **Run after →
    Scope: has failed / timed out**, does **Update item** (WalkStatus `Error`, FilingError with
    the failed step's message) and emails the office.

### 5.3 Automated flow: RQI Filing Alert

- **Trigger:** RQI WALKS item created or modified, condition
  `@or(equals(triggerOutputs()?['body/WalkStatus/Value'], 'Folder Not Found'), equals(triggerOutputs()?['body/WalkStatus/Value'], 'Error'))`
- **Send email** to the office: job number, what's wrong (FilingError), and what to do ("create
  the job folder `N2601001_…` in `2026 Jobs\01 JANUARY`; the walk will file itself within 10
  minutes").
- The courier only changes the status when it changes, so this sends one email per problem, not
  one every 10 minutes.

---

## Phase 6 — Courier: copy finished walks to P:

`courier/Copy-RoughWalksToP.ps1` runs on the 24-hour server every 10 minutes.

### 6.1 What it does each run

1. Connects to SharePoint as the **RQI Courier** app (certificate, no password).
2. Gets RQI WALKS rows with Status **Ready to File** or **Folder Not Found**.
3. For each one, works out the folder from the job number: `2601001` → `<JobRoot>\2026 Jobs\01
   JANUARY\` → the one folder starting with `N2601001_`.
   - None or more than one → Status **Folder Not Found** (FilingError says which). It's tried
     again next run.
4. Downloads the walk's files from WalkDrop to a local staging folder, then copies them in:
   - `N2601001-RWC.pdf` → job folder root. If that name is taken: `N2601001-RWC-2.pdf`, `-3`, …
   - Photos → `Rough Walk Photos\`. If `N2601001-RW-01.jpg` is taken, it uses the next free
     number, so re-walk photos never overwrite earlier ones.
   - `_work\` is skipped.
5. Checks every copied file's size against SharePoint, then sets Status **Filed**, Filed On,
   Filed To and PDF Name.
6. Never deletes anything on P:, and never deletes from WalkDrop. Clean up WalkDrop with the
   optional cleanup (9.4).
7. Writes a log per day to the log folder and keeps 60 days.

`-WhatIf` shows what it would do without copying or updating anything. Use it for the first
runs.

### 6.2 Install

1. Copy the `courier` folder to the server, for example `D:\Scripts\RQI Courier\`.
2. Copy `courier-config.sample.json` to `courier-config.json` and fill in the values from 0.2 and
   0.3.
3. Test as the courier account:
   ```
   pwsh -File "D:\Scripts\RQI Courier\Copy-RoughWalksToP.ps1" -WhatIf
   ```
4. **Task Scheduler → Create Task:**
   - General: run as `DOMAIN\svc-rqi-courier`, **Run whether user is logged on or not**
   - Triggers: daily, repeat every **10 minutes** for **1 day**
   - Action: `pwsh.exe`, arguments
     `-NoProfile -File "D:\Scripts\RQI Courier\Copy-RoughWalksToP.ps1"`
   - Settings: **If the task is already running: Do not start a new instance**

---

## Phase 7 — Testing

### 7.1 Photo spike test (do first, about an hour)

A throwaway app with one Add picture control, a gallery, a Submit button and the RQIUploadFile
flow:
1. On an iPad in the Power Apps app, **airplane mode on**, take 15 photos, tap Submit (they go to
   the Outbox), close the app completely.
2. Reopen the app, still offline: the 15 photos must still be there (LoadData).
3. Airplane mode off: all 15 must arrive in WalkDrop at full size (2–5 MB each, not thumbnails).
4. Repeat with 40 photos to find the limit of offline storage on the iPad.

If photos don't survive the restart, the fallback is to store each photo as text
(`JSON(Img, JSONFormat.IncludeBinaryData)`) before SaveData. If uploads are too slow on cell
signal, the app can shrink photos before upload. Both are small changes, made before Phase 4 goes
further.

### 7.2 Worked example

Use the paper form you sent: **N2601001**, Archer, Lot 169, Cold Springs. Inspector Steve, 1-20-26,
Installer CHL #702, Supervisor Jared. 2nd Floor N/A, 1st Floor Yes, Basement No. Exterior of Home
Yes, Furnace Set Yes, Zoning No, Misc N/A. Comments "100%".

Make a test job folder `N2601001_ARCHER_169 COLD SPRINGS` in a **test copy** of the JOB FOLDERS
tree first (set `JobRoot` in the config to the test copy until 7.3 passes).

| # | Test | Expected |
|---|---|---|
| 1 | Submit online | Row goes Uploading → Submitted → Processing → Ready to File within 2 minutes; supervisor email with PDF |
| 2 | PDF | Every answer in the right place, X under the right Yes/No/N/A, signature fits its box, one page |
| 3 | Courier run | PDF in the job folder root, photos in `Rough Walk Photos`, Status Filed |
| 4 | Submit with airplane mode on | Walk in the Outbox; sends itself within a minute of signal coming back |
| 5 | Kill the signal halfway through sending | One row only (LocalId), all photos there after the retry |
| 6 | Job folder missing | Status Folder Not Found, one email; create the folder; filed on the next run, no second email |
| 7 | Two folders start with `N2601001_` | Folder Not Found, FilingError names both |
| 8 | Re-walk same job | `N2601001-RWC-2.pdf`, photos numbered after the existing ones |
| 9 | Not Ready for Walk | Walk questions optional; email subject says NOT READY |
| 10 | 100% Complete | Comments say "100% Complete"; subject says 100% COMPLETE |
| 11 | Comments with quotes, line breaks, "100%" | PDF shows them exactly |
| 12 | Phone (iPhone) | Every screen usable without sideways scrolling |
| 13 | Job typed as `n2601001` | Found |
| 14 | Builder with no row in BUILDER SUPERVISORS | Error status and office email, not a silent failure |

### 7.3 Go live on P:

Set `JobRoot` to the real UNC path, run once with `-WhatIf`, read the log, then run for real.

---

## Phase 8 — Rollout

1. Share the app with the field assistants (Read access to RNC JOBS and BUILDER SUPERVISORS,
   Contribute to RQI WALKS).
2. On each iPad and phone: install **Power Apps** from the App Store, sign in, open the app once
   **with signal** so the offline job list is saved, and add it to the home screen.
3. Each morning: open the app with signal (refreshes the jobs). If they're offline all day,
   everything they submit sends when they're next online.
4. Run paper and app side by side for a week, then switch.

---

## Phase 9 — Later

- **9.1 100% Complete straight to Production.** The submitted flow emails the payroll contact
  directly for 100% jobs, instead of the supervisor forwarding.
- **9.2 Photo page in the PDF.** A second page with thumbnails of the photos.
- **9.3 Reporting.** RQI WALKS already has everything for a list view or Power BI: walks per
  assistant, 100% rate by installer or builder, NIF counts. Same approach as the Finish Crew
  Punch List (Phase 8 there).
- **9.4 WalkDrop cleanup.** Monthly flow deletes WalkDrop folders whose walk has been Filed for
  30+ days.
- **9.5 Job folders in SharePoint.** If the company moves JOB FOLDERS into SharePoint later, the
  courier goes away. The flow saves straight into the job folder using the same folder rule.

---

## Open questions

1. A1 to A7 above.
2. A copy of the Production Schedule (headers and a few rows) and its path, for Phase 2.
3. The UNC path behind P: and the server name (Phase 0.2).
4. Who in the office gets the Folder Not Found and Error emails.
