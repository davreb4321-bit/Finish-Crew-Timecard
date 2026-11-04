' ============================================================================
' Send to Finish Crew App
'
' One-click hand-off for the Production office. Fill in the daily cover sheet
' (table FINISHCOVER, job date in B4) in desktop Excel, then click the button.
'
' The macro:
'   0. Treats rows hidden by a filter (or hidden by hand) as leftovers from an
'      earlier day: after a warning it deletes their typed entries (lookup
'      formulas stay), then removes the filter so every row is visible.
'   1. Checks B4 holds a real date and at least one Job Number is entered.
'   2. Warns if any Production Schedule lookup shows "Not found".
'   3. Copies the cover sheet into a new workbook, replaces the P: drive
'      lookups with their values, and saves it as .xlsx straight into the
'      SharePoint "Cover Sheet Import" folder with a unique, dated name.
'   4. Closes the copy. The open workbook is untouched.
'
' The "Finish Crew Cover Sheet Import" flow picks the file up automatically,
' fills SCHEDULED JOBS, and emails the sender.
' ============================================================================
Option Explicit

' Folder the import flow watches. Must end with "/".
Private Const IMPORT_FOLDER As String = _
    "https://airtronhvac.sharepoint.com/sites/IndyWarrantyTracking-NRG365/Shared Documents/General/Cover Sheet Import/"

Private Const TABLE_NAME As String = "FINISHCOVER"
Private Const DATE_CELL As String = "B4"

Public Sub SendToFinishCrewApp()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim jobDate As Date
    Dim jobCount As Long
    Dim notFound As Long
    Dim fileName As String
    Dim newWb As Workbook
    Dim links As Variant
    Dim i As Long

    ' --- Find the cover sheet table -------------------------------------
    Set lo = FindCoverTable(ws)
    If lo Is Nothing Then
        MsgBox "Table " & TABLE_NAME & " was not found. Use the Finish Cover Sheet template.", _
            vbCritical, "Send to Finish Crew App"
        Exit Sub
    End If

    ' --- Delete jobs hidden by a filter, then show every row ---------------
    If Not ClearHiddenRows(ws, lo) Then Exit Sub

    ' --- Pull the latest Production Schedule values ----------------------
    On Error Resume Next
    links = ThisWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            ThisWorkbook.UpdateLink Name:=links(i), Type:=xlExcelLinks
        Next i
    End If
    On Error GoTo 0
    Application.CalculateFull

    ' --- Validate the job date -------------------------------------------
    If Not IsDate(ws.Range(DATE_CELL).Value) Then
        MsgBox "The Job Date in " & DATE_CELL & " (" & ws.Range(DATE_CELL).Text & _
            ") is not a valid date. Enter it as m/d/yyyy and try again.", _
            vbExclamation, "Send to Finish Crew App"
        ws.Activate
        ws.Range(DATE_CELL).Select
        Exit Sub
    End If
    jobDate = CDate(ws.Range(DATE_CELL).Value)
    If Year(jobDate) < Year(Date) - 1 Or Year(jobDate) > Year(Date) + 1 Then
        MsgBox "The Job Date in " & DATE_CELL & " (" & Format(jobDate, "m/d/yyyy") & _
            ") doesn't look right. Check the year and try again.", _
            vbExclamation, "Send to Finish Crew App"
        ws.Activate
        ws.Range(DATE_CELL).Select
        Exit Sub
    End If

    ' --- Make sure there is at least one job -----------------------------
    If lo.DataBodyRange Is Nothing Then
        MsgBox "No jobs are entered on the cover sheet.", vbExclamation, "Send to Finish Crew App"
        Exit Sub
    End If
    jobCount = Application.WorksheetFunction.CountA(lo.ListColumns("Job Number").DataBodyRange)
    If jobCount = 0 Then
        MsgBox "No jobs are entered on the cover sheet.", vbExclamation, "Send to Finish Crew App"
        Exit Sub
    End If

    ' --- Warn about failed Production Schedule lookups -------------------
    notFound = Application.WorksheetFunction.CountIf(lo.DataBodyRange, "Not found")
    If notFound > 0 Then
        If MsgBox(notFound & " cell(s) show ""Not found"" (job number not in the Production Schedule)." & _
                vbCrLf & vbCrLf & "Send anyway? Installers will have to fill those fields in.", _
                vbYesNo + vbQuestion, "Send to Finish Crew App") = vbNo Then
            Exit Sub
        End If
    End If

    ' --- Build a values-only .xlsx copy and save it to SharePoint ---------
    fileName = "Finish Cover " & Format(jobDate, "yyyy-mm-dd") & " sent " & _
        Format(Now, "yyyy-mm-dd hhmmss") & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo SaveFailed

    ws.Copy                               ' new workbook containing just this sheet
    Set newWb = ActiveWorkbook

    ' Replace the P: drive XLOOKUPs with their current values
    links = newWb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            newWb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If

    newWb.SaveAs fileName:=IMPORT_FOLDER & fileName, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    Set newWb = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ThisWorkbook.Activate

    MsgBox jobCount & " job(s) for " & Format(jobDate, "m/d/yyyy") & " sent to the Finish Crew App." & _
        vbCrLf & vbCrLf & "You'll get a confirmation email in a few minutes.", _
        vbInformation, "Send to Finish Crew App"
    Exit Sub

SaveFailed:
    Dim errText As String
    errText = Err.Description
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    MsgBox "The cover sheet could not be sent." & vbCrLf & vbCrLf & errText & vbCrLf & vbCrLf & _
        "Check that you are signed in to Excel with your Airtron account and have access to " & _
        "the Cover Sheet Import folder.", vbCritical, "Send to Finish Crew App"
End Sub

' Rows hidden by a filter (or hidden by hand) are treated as leftovers from an
' earlier day, so they are never sent. If any of them holds a Job Number the
' user is warned first (No is the default). Typed entries in the hidden rows are
' cleared; the lookup formulas stay. Then every row is shown again.
' Returns False when the user stops or the rows can't be cleared.
Private Function ClearHiddenRows(ws As Worksheet, lo As ListObject) As Boolean
    Dim r As ListRow
    Dim hiddenRows As Collection
    Dim hr As Variant
    Dim jobIdx As Long
    Dim job As String
    Dim jobList As String
    Dim jobsHidden As Long
    Dim jobsListed As Long
    Dim constants As Range

    ClearHiddenRows = True
    Set hiddenRows = New Collection
    If lo.DataBodyRange Is Nothing Then Exit Function

    jobIdx = lo.ListColumns("Job Number").Index
    For Each r In lo.ListRows
        If r.Range.EntireRow.Hidden Then
            hiddenRows.Add r
            job = Trim(r.Range.Cells(1, jobIdx).Text)
            If job <> "" Then
                jobsHidden = jobsHidden + 1
                If jobsListed < 10 Then
                    jobList = jobList & vbCrLf & "      " & job
                    jobsListed = jobsListed + 1
                End If
            End If
        End If
    Next r

    If hiddenRows.Count = 0 Then
        ShowAllRows ws, lo
        Exit Function
    End If

    If jobsHidden > 0 Then
        If jobsHidden > jobsListed Then
            jobList = jobList & vbCrLf & "      ...and " & (jobsHidden - jobsListed) & " more"
        End If
        If MsgBox(jobsHidden & " job(s) are hidden by a filter and will NOT be sent:" & vbCrLf & _
                  jobList & vbCrLf & vbCrLf & _
                  "These are usually left over from an earlier day." & vbCrLf & vbCrLf & _
                  "Yes  =  DELETE the hidden jobs and send only the visible ones." & vbCrLf & _
                  "No   =  stop and show all rows so you can check them.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Send to Finish Crew App") <> vbYes Then
            ShowAllRows ws, lo
            ClearHiddenRows = False
            Exit Function
        End If
    End If

    ' Clear typed entries in the hidden rows (SpecialCells includes hidden cells).
    On Error GoTo ClearFailed
    For Each hr In hiddenRows
        Set constants = Nothing
        On Error Resume Next
        Set constants = hr.Range.SpecialCells(xlCellTypeConstants)
        On Error GoTo ClearFailed
        If Not constants Is Nothing Then constants.ClearContents
    Next hr
    On Error GoTo 0

    ShowAllRows ws, lo
    Exit Function

ClearFailed:
    ClearHiddenRows = False
    MsgBox "The hidden rows could not be cleared (" & Err.Description & ")." & vbCrLf & vbCrLf & _
        "Nothing was sent. Clear the filter, delete the old jobs (or use All Clear), and try again.", _
        vbCritical, "Send to Finish Crew App"
End Function

' Removes the table and sheet filters and unhides every table row.
Private Sub ShowAllRows(ws As Worksheet, lo As ListObject)
    On Error Resume Next
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    If ws.FilterMode Then ws.ShowAllData
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.EntireRow.Hidden = False
    On Error GoTo 0
End Sub

' Returns the FINISHCOVER table and sets ws to its sheet.
Private Function FindCoverTable(ByRef ws As Worksheet) As ListObject
    Dim sh As Worksheet
    Dim t As ListObject
    For Each sh In ThisWorkbook.Worksheets
        For Each t In sh.ListObjects
            If StrComp(t.Name, TABLE_NAME, vbTextCompare) = 0 Then
                Set ws = sh
                Set FindCoverTable = t
                Exit Function
            End If
        Next t
    Next sh
End Function
