Attribute VB_Name = "BackChargeImport"
' ============================================================================
' Back Charge Import + Property Damage PDFs
'
' ImportBackCharges
'   Brings the rows from BC_Import_Ready.xlsx (made by bc_import.py) into
'   Back Charge Data, starting at row 10, columns A-W. Yellow cells are
'   suggestions to confirm, orange cells could not be found. Replaces the
'   copy/rearrange step, the XLOOKUP formulas and the All Clear button.
'
' CreatePropertyDamagePDFs
'   Run after the drop-downs are confirmed. Fills the Property Damage Report
'   for each row, saves the PDF to today's folder in the P: drive loading
'   folder, and files a copy straight into its G: drive job folder, named
'   "<Lot> <Subdivision>, <Address>, <Deduction Date>, <Builder>".
'   The folder is created if it does not exist yet.
' ============================================================================
Option Explicit

Private Const LOADING_FOLDER As String = "P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER\"
Private Const IMPORT_FILE As String = "BC_Import_Ready.xlsx"
Private Const G_DRIVE_ROOT As String = "G:\Mandi\Safety\Incident Investigations\2026 Incidents\Builder backcharges, 2026\"

Private Const DATA_SHEET As String = "Back Charge Data"
Private Const REPORT_SHEET As String = "Property Damage Report"
Private Const FIRST_ROW As Long = 10
Private Const LAST_CLEAR_ROW As Long = 400

Private Const COLOR_MISSING As Long = 11389944   ' RGB(248, 203, 173) orange, set by bc_import.py

' ----------------------------------------------------------------------------
Public Sub ImportBackCharges()
    Dim wsData As Worksheet, src As Workbook, wsSrc As Worksheet
    Dim path As String, lastSrc As Long, lastData As Long
    Dim r As Long, c As Long, n As Long

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)

    path = LOADING_FOLDER & IMPORT_FILE
    If Dir(path) = "" Then
        MsgBox "Could not find " & path & vbNewLine & vbNewLine & _
               "Run bc_import.py first (double-click run_import.bat).", vbExclamation, "Import Back Charges"
        Exit Sub
    End If

    lastData = LastUsedRow(wsData)
    If lastData >= FIRST_ROW Then
        If MsgBox("Back Charge Data already has rows " & FIRST_ROW & " to " & lastData & "." & vbNewLine & vbNewLine & _
                  "Make sure they were sent to the Registry. Replace them with the new import?", _
                  vbYesNo + vbExclamation, "Import Back Charges") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Clear values and highlight only; drop-down lists and formatting stay
    With wsData.Range(wsData.Cells(FIRST_ROW, 1), wsData.Cells(LAST_CLEAR_ROW, 23))
        .ClearContents
        .Interior.Pattern = xlNone
    End With

    Set src = Workbooks.Open(path, ReadOnly:=True)
    Set wsSrc = src.Sheets(1)
    lastSrc = wsSrc.Cells(wsSrc.Rows.Count, 14).End(xlUp).Row    ' column N = PO/Back Charge #
    If wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).Row > lastSrc Then lastSrc = wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).Row

    For r = 2 To lastSrc
        For c = 1 To 23
            With wsSrc.Cells(r, c)
                wsData.Cells(FIRST_ROW + n, c).Value = .Value
                If .Interior.Pattern <> xlNone Then wsData.Cells(FIRST_ROW + n, c).Interior.Color = .Interior.Color
            End With
        Next c
        n = n + 1
    Next r
    src.Close SaveChanges:=False

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    wsData.Activate
    wsData.Cells(FIRST_ROW, 1).Select
    MsgBox n & " back charge(s) imported." & vbNewLine & vbNewLine & _
           "Yellow = suggested, please confirm." & vbNewLine & _
           "Orange = not found, please fill in." & vbNewLine & vbNewLine & _
           "Skipped duplicates and close calls are on the Review tab of " & IMPORT_FILE & ".", _
           vbInformation, "Import Back Charges"
End Sub

' ----------------------------------------------------------------------------
Public Sub CreatePropertyDamagePDFs()
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim i As Long, lastRow As Long, made As Long, filed As Long, missingRows As Long
    Dim pFolder As String, gFolder As String, fileName As String, failures As String
    Dim jobAddr As String, subDiv As String, lotNum As String, filePrefix As String
    Dim builderName As String, jobID As String, poNum As String

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsReport = ThisWorkbook.Sheets(REPORT_SHEET)
    lastRow = LastUsedRow(wsData)
    If lastRow < FIRST_ROW Then
        MsgBox "Back Charge Data has no rows to print.", vbInformation, "Property Damage PDFs"
        Exit Sub
    End If

    For i = FIRST_ROW To lastRow
        If RowHasColor(wsData, i, COLOR_MISSING) Then missingRows = missingRows + 1
    Next i
    If missingRows > 0 Then
        If MsgBox(missingRows & " row(s) still have orange cells that were not filled in." & vbNewLine & _
                  "Create the PDFs anyway?", vbYesNo + vbExclamation, "Property Damage PDFs") = vbNo Then Exit Sub
    End If

    pFolder = LOADING_FOLDER & Format(Date, "yyyy-mm-dd") & "\"
    If Dir(pFolder, vbDirectory) = "" Then MkDir pFolder
    If Dir(G_DRIVE_ROOT, vbDirectory) = "" Then
        MsgBox "Cannot reach the G: drive folder:" & vbNewLine & G_DRIVE_ROOT & vbNewLine & vbNewLine & _
               "The PDFs will be saved on P: only.", vbExclamation, "Property Damage PDFs"
    End If

    ' Keep the sheet's buttons off the PDF
    On Error Resume Next
    wsReport.Buttons.PrintObject = False
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = FIRST_ROW To lastRow
        Application.StatusBar = "Creating PDF " & (i - FIRST_ROW + 1) & " of " & (lastRow - FIRST_ROW + 1) & "..."

        wsReport.Range("G36, B36, B12, G33, B21, B33, B30, G21, B24, G24, B27, G27, G30, G9, B9, B15, B18").Value = ""

        jobAddr = Trim(wsData.Cells(i, "D").Value)
        subDiv = Trim(wsData.Cells(i, "E").Value)
        lotNum = Trim(wsData.Cells(i, "F").Value)
        If lotNum <> "" Or subDiv <> "" Then
            filePrefix = Trim(lotNum & " " & subDiv)
        ElseIf jobAddr <> "" Then
            filePrefix = jobAddr
        Else
            filePrefix = "Unknown_Location"
        End If

        With wsReport
            .Range("B9").Value = ValOrNA(wsData.Cells(i, "N"))    ' PO#
            .Range("B12").Value = IIf(jobAddr <> "", jobAddr, filePrefix)
            .Range("B15").Value = ValOrNA(wsData.Cells(i, "W"))   ' Customer Name
            .Range("B18").Value = ValOrNA(wsData.Cells(i, "B"))   ' Phone
            If .Range("B18").Value <> "N/A" Then .Range("B18").NumberFormat = "(###) ###-####"
            .Range("B21").Value = ValOrNA(wsData.Cells(i, "H"))   ' Incident Date
            .Range("B24").Value = ValOrNA(wsData.Cells(i, "O"))   ' Responsible Dept
            .Range("B27").Value = ValOrNA(wsData.Cells(i, "T"))   ' Property Damaged
            .Range("B30").Value = ValOrNA(wsData.Cells(i, "M"))   ' Phase
            .Range("B33").Value = ValOrNA(wsData.Cells(i, "I"))   ' Deduction Date
            .Range("B36").Value = ValOrNA(wsData.Cells(i, "C"))   ' Job #
            .Range("G9").Value = ValOrNA(wsData.Cells(i, "V"))    ' Respondent
            .Range("G21").Value = ValOrNA(wsData.Cells(i, "L"))   ' Claim Type
            .Range("G24").Value = ValOrNA(wsData.Cells(i, "S"))   ' Incident Type
            .Range("G27").Value = ValOrNA(wsData.Cells(i, "Q"))   ' Description
            .Range("G30").Value = ValOrNA(wsData.Cells(i, "U"))   ' Responsible Person
            .Range("G33").Value = ValOrNA(wsData.Cells(i, "G"))   ' BC Amount
            .Range("G36").Value = ValOrNA(wsData.Cells(i, "A"))   ' Payee (builder)
        End With

        builderName = IIf(Trim(wsData.Cells(i, "A").Value) <> "", Trim(wsData.Cells(i, "A").Value), "No_Builder")
        jobID = IIf(Trim(wsData.Cells(i, "C").Value) <> "", Trim(wsData.Cells(i, "C").Value), "Pending")
        poNum = Trim(wsData.Cells(i, "N").Value)
        If poNum = "" Then poNum = "Row" & i & "-" & Format(Now, "yyyymmdd-hhnnss")
        ' PO# keeps two back charges on the same job from overwriting each other
        fileName = CleanName(filePrefix & " - " & builderName & " - PROPERTY DAMAGE - " & jobID & " - " & poNum) & ".pdf"

        wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pFolder & fileName, _
            Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If Dir(pFolder & fileName) <> "" Then
            made = made + 1
            If Dir(G_DRIVE_ROOT, vbDirectory) <> "" Then
                gFolder = G_DRIVE_ROOT & JobFolderName(wsData, i) & "\"
                On Error Resume Next
                If Dir(gFolder, vbDirectory) = "" Then MkDir gFolder
                FileCopy pFolder & fileName, gFolder & fileName
                If Err.Number = 0 Then
                    filed = filed + 1
                Else
                    failures = failures & vbNewLine & "Row " & i & ": " & Err.Description
                End If
                Err.Clear
                On Error GoTo 0
            End If
        Else
            failures = failures & vbNewLine & "Row " & i & ": PDF was not created"
        End If
    Next i

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "PDFs created on P: " & made & vbNewLine & _
           "Filed into G: job folders: " & filed & _
           IIf(failures <> "", vbNewLine & vbNewLine & "Problems:" & failures, ""), _
           IIf(failures = "", vbInformation, vbExclamation), "Property Damage PDFs"
End Sub

' ----------------------------------------------------------------------------
' "2 HICKORY RUN, 10109 BITTERNUT RUN LANE, 10-19-2026, MI HOMES"
Private Function JobFolderName(ws As Worksheet, r As Long) As String
    Dim lotSub As String, addr As String, dt As String, parts As String

    lotSub = UCase(Trim(Trim(ws.Cells(r, "F").Value) & " " & Trim(ws.Cells(r, "E").Value)))
    addr = ExpandStreet(UCase(Trim(ws.Cells(r, "D").Value)))
    If IsDate(ws.Cells(r, "I").Value) Then
        dt = Format(CDate(ws.Cells(r, "I").Value), "m-d-yyyy")
    Else
        dt = "NO DEDUCTION DATE"
    End If

    If lotSub <> "" Then parts = lotSub & ", "
    If addr <> "" Then parts = parts & addr & ", "
    JobFolderName = CleanName(parts & dt & ", " & FolderBuilder(ws.Cells(r, "A").Value))
End Function

' Builder names as they appear on the G: drive folders. Add new builders here.
Private Function FolderBuilder(ByVal b As String) As String
    b = UCase(Trim(b))
    Select Case True
        Case b Like "*PULTE*": FolderBuilder = "PULTE"
        Case b Like "*M.I.*", b Like "MI HOMES*", b Like "M I HOMES*": FolderBuilder = "MI HOMES"
        Case b Like "*ARBOR*": FolderBuilder = "ARBOR"
        Case b Like "*FISCHER*": FolderBuilder = "FISCHER HOMES"
        Case b Like "*EC NEW VISION*": FolderBuilder = "EC NEW VISION"
        Case b Like "*LENNAR*": FolderBuilder = "LENNAR"
        Case b Like "*BEAZER*": FolderBuilder = "BEAZER"
        Case b Like "*RYAN*": FolderBuilder = "RYAN HOMES"
        Case b Like "*DREES*": FolderBuilder = "DREES"
        Case b Like "*HORTON*": FolderBuilder = "DR HORTON"
        Case b Like "*WEEKLEY*": FolderBuilder = "DAVID WEEKLEY"
        Case b Like "*ESTRIDGE*": FolderBuilder = "ESTRIDGE"
        Case b Like "*SILVERTHORNE*": FolderBuilder = "SILVERTHORNE"
        Case b Like "*TAYLOR MORRISON*": FolderBuilder = "TAYLOR MORRISON"
        Case b Like "*DAVIS*": FolderBuilder = "DAVIS"
        Case b = "": FolderBuilder = "NO BUILDER"
        Case Else: FolderBuilder = b
    End Select
End Function

' "785 ROSEBUD LN" -> "785 ROSEBUD LANE" (last word only)
Private Function ExpandStreet(ByVal a As String) As String
    Dim p As Long, lastWord As String, full As String
    p = InStrRev(a, " ")
    If p = 0 Then ExpandStreet = a: Exit Function
    lastWord = Replace(Mid(a, p + 1), ".", "")
    Select Case lastWord
        Case "LN": full = "LANE"
        Case "DR": full = "DRIVE"
        Case "CT": full = "COURT"
        Case "ST": full = "STREET"
        Case "AVE", "AV": full = "AVENUE"
        Case "RD": full = "ROAD"
        Case "PL": full = "PLACE"
        Case "CIR": full = "CIRCLE"
        Case "BLVD": full = "BOULEVARD"
        Case "TRL": full = "TRAIL"
        Case "PKWY": full = "PARKWAY"
        Case "TER": full = "TERRACE"
        Case Else: full = lastWord
    End Select
    ExpandStreet = Left(a, p) & full
End Function

Private Function CleanName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "-")
    Next ch
    CleanName = Trim(s)
End Function

Private Function ValOrNA(c As Range) As String
    If Trim(c.Value) = "" Then ValOrNA = "N/A" Else ValOrNA = Trim(c.Value)
End Function

Private Function RowHasColor(ws As Worksheet, r As Long, clr As Long) As Boolean
    Dim c As Long
    For c = 1 To 23
        If ws.Cells(r, c).Interior.Pattern <> xlNone Then
            If ws.Cells(r, c).Interior.Color = clr Then RowHasColor = True: Exit Function
        End If
    Next c
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Range("A" & FIRST_ROW & ":W" & LAST_CLEAR_ROW).Find("*", , xlValues, , xlByRows, xlPrevious)
    If f Is Nothing Then LastUsedRow = FIRST_ROW - 1 Else LastUsedRow = f.Row
End Function
