' ============================================================================
' Clear Cover Sheet
'
' Resets the daily-entry cover sheet for the next day:
'   - Clears typed entries in table FINISHCOVER (Job Number, Lead Installer,
'     Helper Installer, Phase and the four comment columns) and the Job Date (B4).
'   - Keeps the Production Schedule (P: drive) formulas in Supervisor, Builder,
'     Subdivision and Lot #. If someone typed over a formula, it is put back.
'   - Keeps drop-down menus (data validation), formatting and buttons.
'   - Resets row heights.
' ============================================================================
Option Explicit

Private Const TABLE_NAME As String = "FINISHCOVER"
Private Const DATE_CELL As String = "B4"

' Production Schedule used by the lookup formulas (only used if a formula has to be rebuilt).
Private Const SCHEDULE_REF As String = _
    "'P:\Indianapolis\Natasha Coffman\[2.22.23 Production Schedule v1.xlsm]Sheet1'!"

Public Sub ClearCoverSheet()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim constants As Range

    Set lo = FindCoverTable(ws)
    If lo Is Nothing Then
        MsgBox "Table " & TABLE_NAME & " was not found.", vbCritical, "Clear Cover Sheet"
        Exit Sub
    End If

    ' ---- Two-step warning (No is the default button, so pressing Enter cancels) ----
    Dim jobCount As Long
    Dim dateText As String

    If Not lo.DataBodyRange Is Nothing Then
        jobCount = Application.WorksheetFunction.CountA(lo.ListColumns("Job Number").DataBodyRange)
    End If
    dateText = ws.Range(DATE_CELL).Text
    If Trim(dateText) = "" Then dateText = "(no date entered)"

    If MsgBox("CAUTION: You are about to clear the Finish Crew Cover Sheet." & vbCrLf & vbCrLf & _
              "Job Date:  " & dateText & vbCrLf & _
              "Jobs entered:  " & jobCount & vbCrLf & vbCrLf & _
              "All jobs, installers, phases, comments and the Job Date will be erased." & vbCrLf & _
              "Make sure this sheet has already been sent with Send to Teams." & vbCrLf & vbCrLf & _
              "Do you want to continue?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Clear Cover Sheet - Warning") <> vbYes Then
        Exit Sub
    End If

    If MsgBox("FINAL WARNING" & vbCrLf & vbCrLf & _
              "This cannot be undone. Ctrl+Z will NOT bring the jobs back." & vbCrLf & vbCrLf & _
              "Clear the cover sheet now?", _
              vbYesNo + vbCritical + vbDefaultButton2, "Clear Cover Sheet - Final Warning") <> vbYes Then
        MsgBox "Nothing was cleared.", vbInformation, "Clear Cover Sheet"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Put back any lookup formula that was typed over, before clearing typed values.
    RestoreLookup lo, "Supervisor", "$D:$D"
    RestoreLookup lo, "Builder", "$A:$A"
    RestoreLookup lo, "Subdivision", "$C:$C"
    RestoreLookup lo, "Lot #", "$B:$B"

    ' Clear only typed values (constants). Formulas, drop-downs and formatting stay.
    If Not lo.DataBodyRange Is Nothing Then
        Set constants = Nothing
        On Error Resume Next
        Set constants = lo.DataBodyRange.SpecialCells(xlCellTypeConstants)
        On Error GoTo 0
        If Not constants Is Nothing Then constants.ClearContents

        lo.DataBodyRange.Rows.AutoFit
    End If

    ' Job Date
    ws.Range(DATE_CELL).ClearContents

    Application.ScreenUpdating = True

    ws.Activate
    ws.Range(DATE_CELL).Select
    MsgBox "Cover sheet cleared. Enter the new Job Date in " & DATE_CELL & ".", _
        vbInformation, "Clear Cover Sheet"
End Sub

' Makes every row of a lookup column hold its formula again.
' Uses the formula from a row that still has it; rebuilds it only if none do.
Private Sub RestoreLookup(lo As ListObject, columnName As String, returnColumn As String)
    Dim col As ListColumn
    Dim cell As Range
    Dim f As String

    On Error Resume Next
    Set col = lo.ListColumns(columnName)
    On Error GoTo 0
    If col Is Nothing Then Exit Sub
    If col.DataBodyRange Is Nothing Then Exit Sub

    For Each cell In col.DataBodyRange.Cells
        If cell.HasFormula Then
            f = cell.Formula2
            Exit For
        End If
    Next cell

    If f = "" Then
        f = "=IFERROR(IF([@[Job Number]]="""","""",XLOOKUP([@[Job Number]]&""""," & _
            SCHEDULE_REF & "$G:$G," & SCHEDULE_REF & returnColumn & ",""Not found"",0)),"""")"
    End If

    For Each cell In col.DataBodyRange.Cells
        If Not cell.HasFormula Then cell.Formula2 = f
    Next cell
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
