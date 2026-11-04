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
'   If the house already has a folder (any date), the PDF goes into it;
'   otherwise a new one is created. "Same house" uses the glossary below, so
'   Ln/Lane, St/Street, Fischer/Fischer Homes etc. count as the same.
'
'   It then files the split accounting sheets (one PDF per back charge, made by
'   bc_import.py and listed on the "Accounting Sheets" tab of
'   BC_Import_Ready.xlsx) into the same house folders. That includes back
'   charges that were skipped because they are already in the Registry.
' ============================================================================
Option Explicit

Private Const LOADING_FOLDER As String = "P:\Indianapolis\IT STUFF FOR INDY\BACK CHARGES\LOADING FOLDER\"
Private Const IMPORT_FILE As String = "BC_Import_Ready.xlsx"
Private Const SHEETS_TAB As String = "Accounting Sheets"
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
    Dim reused As Long, created As Long
    Dim pFolder As String, gFolder As String, fileName As String, failures As String
    Dim houses As Collection
    Dim jobAddr As String, subDiv As String, lotNum As String, filePrefix As String
    Dim builderName As String, jobID As String, poNum As String
    Dim hasRows As Boolean, acctFiled As Long

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsReport = ThisWorkbook.Sheets(REPORT_SHEET)
    lastRow = LastUsedRow(wsData)
    hasRows = (lastRow >= FIRST_ROW)    ' no rows: still file the accounting sheets

    For i = FIRST_ROW To lastRow
        If RowHasColor(wsData, i, COLOR_MISSING) Then missingRows = missingRows + 1
    Next i
    If missingRows > 0 Then
        If MsgBox(missingRows & " row(s) still have orange cells that were not filled in." & vbNewLine & _
                  "Create the PDFs anyway?", vbYesNo + vbExclamation, "Property Damage PDFs") = vbNo Then Exit Sub
    End If

    pFolder = LOADING_FOLDER & Format(Date, "yyyy-mm-dd") & "\"
    If Dir(pFolder, vbDirectory) = "" Then MkDir pFolder
    If Dir(G_DRIVE_ROOT, vbDirectory) <> "" Then
        Set houses = LoadHouseFolders()
    Else
        MsgBox "Cannot reach the G: drive folder:" & vbNewLine & G_DRIVE_ROOT & vbNewLine & vbNewLine & _
               "The PDFs will be saved on P: only, and the accounting sheets will not be filed.", _
               vbExclamation, "Property Damage PDFs"
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
            If Not houses Is Nothing Then
                gFolder = HouseFolder(houses, JobFolderName(wsData, i), reused, created)
                On Error Resume Next
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

    If Not houses Is Nothing Then
        Application.StatusBar = "Filing accounting sheets..."
        FileAccountingSheets houses, acctFiled, reused, created, failures
    End If

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox IIf(hasRows, "", "Back Charge Data has no rows, so no property damage PDFs were made." & vbNewLine & vbNewLine) & _
           "Property damage PDFs created on P: " & made & vbNewLine & _
           "Property damage PDFs filed into G: " & filed & vbNewLine & _
           "Accounting sheets filed into G: " & acctFiled & vbNewLine & _
           "(existing house folders used: " & reused & ", new folders: " & created & ")" & _
           IIf(failures <> "", vbNewLine & vbNewLine & "Problems:" & failures, ""), _
           IIf(failures = "", vbInformation, vbExclamation), "Property Damage PDFs"
End Sub

' ----------------------------------------------------------------------------
' The house's G: drive folder (existing one if found, else newly created) with "\" on the end
Private Function HouseFolder(houses As Collection, ByVal newFolder As String, _
                             ByRef reused As Long, ByRef created As Long) As String
    Dim matchFolder As String
    matchFolder = FindHouseFolder(houses, newFolder)
    If matchFolder <> "" Then
        reused = reused + 1
        HouseFolder = G_DRIVE_ROOT & matchFolder & "\"
    Else
        On Error Resume Next
        MkDir G_DRIVE_ROOT & newFolder
        If Err.Number = 0 Then
            houses.Add ParseFolder(newFolder)
            created = created + 1
        End If
        On Error GoTo 0
        HouseFolder = G_DRIVE_ROOT & newFolder & "\"
    End If
End Function

' Copies each split accounting sheet listed on the Accounting Sheets tab into its house folder
Private Sub FileAccountingSheets(houses As Collection, ByRef acctFiled As Long, _
                                 ByRef reused As Long, ByRef created As Long, ByRef failures As String)
    Dim src As Workbook, ws As Worksheet, r As Long, lastR As Long
    Dim pdfPath As String, pdfName As String, gFolder As String, newFolder As String

    If Dir(LOADING_FOLDER & IMPORT_FILE) = "" Then Exit Sub
    Set src = Workbooks.Open(LOADING_FOLDER & IMPORT_FILE, ReadOnly:=True)
    On Error Resume Next
    Set ws = src.Sheets(SHEETS_TAB)
    On Error GoTo 0
    If ws Is Nothing Then src.Close SaveChanges:=False: Exit Sub   ' made by an older bc_import.py

    lastR = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For r = 2 To lastR
        pdfPath = Trim(ws.Cells(r, 3).Value)                        ' C: Split PDF
        If pdfPath <> "" Then
            If Dir(pdfPath) = "" Then
                failures = failures & vbNewLine & ws.Cells(r, 1).Value & ": accounting sheet PDF not found"
            Else
                pdfName = Mid(pdfPath, InStrRev(pdfPath, "\") + 1)
                ' H lot, G subdivision, F address, I deduction date, D builder
                newFolder = FolderNameFrom(ws.Cells(r, 8).Value, ws.Cells(r, 7).Value, ws.Cells(r, 6).Value, _
                                           ws.Cells(r, 9).Value, ws.Cells(r, 4).Value)
                gFolder = HouseFolder(houses, newFolder, reused, created)
                On Error Resume Next
                FileCopy pdfPath, gFolder & pdfName
                If Err.Number = 0 Then
                    acctFiled = acctFiled + 1
                Else
                    failures = failures & vbNewLine & ws.Cells(r, 1).Value & " accounting sheet: " & Err.Description
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next r
    src.Close SaveChanges:=False
End Sub

' ----------------------------------------------------------------------------
' GLOSSARY (keep in step with the glossary in bc_import.py)
'
' Builders: "<name used in G: drive folder names>=<text that identifies it>|<more>"
' A builder matches when any of its text is CONTAINED in the name being read.
' Capitals, punctuation and extra spaces are ignored ("M.I. Homes" = "M I HOMES").
' More specific entries go first. Add new builders or spellings here.
Private Function BuilderGlossary() As String
    Dim g As String
    g = g & "PULTE=PULTE" & vbLf
    g = g & "MI HOMES=M I HOMES|MI HOMES|M I SCHOTTENSTEIN" & vbLf
    g = g & "ARBOR=ARBOR" & vbLf
    g = g & "FISCHER HOMES=FISCHER|FISHER HOMES|FISHER MANAGEMENT" & vbLf
    g = g & "EC NEW VISION=EC NEW VISION|NEW VISION" & vbLf
    g = g & "LENNAR=LENNAR" & vbLf
    g = g & "BEAZER=BEAZER" & vbLf
    g = g & "RYAN HOMES=RYAN HOMES" & vbLf
    g = g & "DREES=DREES" & vbLf
    g = g & "DR HORTON=HORTON" & vbLf
    g = g & "DAVID WEEKLEY=WEEKLEY" & vbLf
    g = g & "ESTRIDGE=ESTRIDGE" & vbLf
    g = g & "SILVERTHORNE=SILVERTHORNE" & vbLf
    g = g & "TAYLOR MORRISON=TAYLOR MORRISON" & vbLf
    g = g & "DAVIS=DAVIS BUILDING|DAVIS BLDG" & vbLf          ' not plain DAVIS: Paul Davis is a restoration company
    g = g & "TWG WHISLETT=WHISLETT|TWG" & vbLf
    g = g & "WESTPORT=WESTPORT" & vbLf
    g = g & "BEACON BUILDERS=BEACON BUILDERS" & vbLf
    g = g & "BEDROCK=BEDROCK" & vbLf
    g = g & "OLD TOWN=OLD TOWN" & vbLf
    g = g & "PRIDE ONE=PRIDE ONE" & vbLf
    g = g & "JOYNER=JOYNER" & vbLf
    g = g & "ONYX AND EAST=ONYX"
    BuilderGlossary = g
End Function

' Street words: "<full word>=<abbreviation>|<abbreviation>"
Private Function StreetGlossary() As String
    Dim g As String
    g = "LANE=LN|LA;STREET=ST|STR;DRIVE=DR|DRV;COURT=CT|CRT;AVENUE=AVE|AV;ROAD=RD;PLACE=PL;"
    g = g & "CIRCLE=CIR|CIRC;BOULEVARD=BLVD|BL;TRAIL=TRL|TR;PARKWAY=PKWY|PKY;TERRACE=TER|TERR;WAY=WY;"
    g = g & "POINTE=PT|POINT;CROSSING=XING|CRSG;COVE=CV;SQUARE=SQ;HIGHWAY=HWY;BEND=BND;GLEN=GLN;"
    g = g & "HOLLOW=HOLW;CREEK=CRK;RIDGE=RDG;COMMONS=CMNS;MEADOWS=MDWS;TRACE=TRCE;RUN=;PASS=;PIKE=;"
    g = g & "LOOP=;ROW=;WALK=;PATH=;VIEW=VW;CHASE=;NORTH=N;SOUTH=S;EAST=E;WEST=W"
    StreetGlossary = g
End Function

Private Function StreetWords() As Object
    Static d As Object
    Dim entry As Variant, kv() As String, ab As Variant
    If d Is Nothing Then
        Set d = CreateObject("Scripting.Dictionary")
        For Each entry In Split(StreetGlossary(), ";")
            kv = Split(entry, "=")
            d(kv(0)) = kv(0)
            For Each ab In Split(kv(1), "|")
                If ab <> "" Then d(ab) = kv(0)
            Next ab
        Next entry
    End If
    Set StreetWords = d
End Function

' Upper case, punctuation to spaces, single spaces: "M.I. Homes" -> "M I HOMES"
Private Function Plain(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase(s)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Z0-9&]" Then out = out & ch Else out = out & " "
    Next i
    out = Trim(out)
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    Plain = out
End Function

' Glossary builder CONTAINED in the text ("Fischer Home Management" -> "FISCHER HOMES"), or ""
Private Function GlossaryBuilder(ByVal b As String) As String
    Dim entry As Variant, kv() As String, v As Variant, t As String
    t = " " & Plain(b) & " "
    For Each entry In Split(BuilderGlossary(), vbLf)
        kv = Split(entry, "=")
        For Each v In Split(kv(1), "|")
            If InStr(t, " " & Plain(CStr(v)) & " ") > 0 Then GlossaryBuilder = kv(0): Exit Function
        Next v
    Next entry
End Function

' Builder name as used in G: drive folder names
Private Function FolderBuilder(ByVal b As String) As String
    FolderBuilder = GlossaryBuilder(b)
    If FolderBuilder = "" Then FolderBuilder = Plain(b)
    If FolderBuilder = "" Then FolderBuilder = "NO BUILDER"
End Function

' "793 Rosebud Ln.", "793 ROSEBUD LANE", "793 Rosebud" -> "793 ROSEBUD"
' House number + street name. Street-type words at the end and a direction after the
' number are dropped so a missing or abbreviated "Lane" still matches.
Private Function AddressKey(ByVal a As String) As String
    Dim w() As String, i As Long, first As Long, last As Long, sw As Object, out As String
    Set sw = StreetWords()
    a = Plain(a)
    If a = "" Then Exit Function
    w = Split(a, " ")
    If UBound(w) < 1 Or Not (Left(w(0), 1) Like "#") Then Exit Function
    For i = 1 To UBound(w)
        If sw.Exists(w(i)) Then w(i) = sw(w(i))
    Next i
    first = 1: last = UBound(w)
    If last > first Then
        If w(first) = "NORTH" Or w(first) = "SOUTH" Or w(first) = "EAST" Or w(first) = "WEST" Then first = first + 1
    End If
    Do While last > first
        If Not sw.Exists(w(last)) Then Exit Do
        last = last - 1
    Loop
    out = w(0)
    For i = first To last
        out = out & " " & w(i)
    Next i
    AddressKey = out
End Function

' "785 ROSEBUD LN" -> "785 ROSEBUD LANE" (last word, for new folder names)
Private Function ExpandStreet(ByVal a As String) As String
    Dim p As Long, lastWord As String, sw As Object
    Set sw = StreetWords()
    p = InStrRev(a, " ")
    If p = 0 Then ExpandStreet = a: Exit Function
    lastWord = Replace(Mid(a, p + 1), ".", "")
    If sw.Exists(lastWord) Then lastWord = sw(lastWord)
    ExpandStreet = Left(a, p) & lastWord
End Function

' ----------------------------------------------------------------------------
' Folder name -> Array(name, builder, "|address keys|", lot/sub part, date)
' "60B PARKROSE, 793 ROSEBUD LANE, 10-23-2026, FISCHER HOMES"
Private Function ParseFolder(ByVal nm As String) As Variant
    Dim part As Variant, p As String, bits() As String, dt As Date
    Dim parts As New Collection, builder As String, keys As String, k As String, i As Long, bIdx As Long

    For Each part In Split(nm, ",")
        p = Trim(part)
        bits = Split(p, "-")
        If p Like "#*-#*-####" And UBound(bits) = 2 Then
            dt = DateSerial(Val(bits(2)), Val(bits(0)), Val(bits(1)))
        ElseIf p <> "" Then
            parts.Add p
        End If
    Next part
    If parts.Count = 0 Then ParseFolder = Array(nm, "", "|", "", dt): Exit Function

    ' Builder: the last part, unless another part is a glossary builder
    bIdx = parts.Count
    For i = parts.Count To 1 Step -1
        If GlossaryBuilder(parts(i)) <> "" Then bIdx = i: Exit For
    Next i
    If parts.Count = 1 Then bIdx = 0
    If bIdx > 0 Then builder = FolderBuilder(parts(bIdx))

    keys = "|"
    For i = 1 To parts.Count
        If i <> bIdx Then
            k = AddressKey(parts(i))
            If k <> "" Then keys = keys & k & "|"
        End If
    Next i
    ParseFolder = Array(nm, builder, keys, Plain(parts(1)), dt)
End Function

' Every folder already in the G: drive root
Private Function LoadHouseFolders() As Collection
    Dim c As New Collection, nm As String
    nm = Dir(G_DRIVE_ROOT & "*", vbDirectory)
    Do While nm <> ""
        If nm <> "." And nm <> ".." Then
            If (GetAttr(G_DRIVE_ROOT & nm) And vbDirectory) = vbDirectory Then c.Add ParseFolder(nm)
        End If
        nm = Dir()
    Loop
    Set LoadHouseFolders = c
End Function

' Existing folder for the same house, or "" if none. Same house = same builder (by the
' glossary) and the same address (by the glossary) or the same lot/subdivision. When a
' house has several folders, the one with the latest date is used.
Private Function FindHouseFolder(folders As Collection, ByVal newName As String) As String
    Dim want As Variant, f As Variant, k As Variant, sameHouse As Boolean, bestDate As Date

    want = ParseFolder(newName)
    For Each f In folders
        If SameBuilder(CStr(f(1)), CStr(want(1))) Then
            sameHouse = False
            For Each k In Split(want(2), "|")
                If k <> "" Then
                    If InStr(f(2), "|" & k & "|") > 0 Then sameHouse = True: Exit For
                End If
            Next k
            If Not sameHouse And want(3) <> "" Then sameHouse = (f(3) = want(3))
            If sameHouse And (FindHouseFolder = "" Or f(4) > bestDate) Then
                FindHouseFolder = f(0)
                bestDate = f(4)
            End If
        End If
    Next f
End Function

' Builders not in the glossary still match when one name contains the other
Private Function SameBuilder(ByVal a As String, ByVal b As String) As Boolean
    If a = "" Or b = "" Then Exit Function
    SameBuilder = (a = b) Or InStr(" " & a & " ", " " & b & " ") > 0 Or InStr(" " & b & " ", " " & a & " ") > 0
End Function

' "2 HICKORY RUN, 10109 BITTERNUT RUN LANE, 10-19-2026, MI HOMES"
Private Function JobFolderName(ws As Worksheet, r As Long) As String
    JobFolderName = FolderNameFrom(ws.Cells(r, "F").Value, ws.Cells(r, "E").Value, ws.Cells(r, "D").Value, _
                                   ws.Cells(r, "I").Value, ws.Cells(r, "A").Value)
End Function

Private Function FolderNameFrom(ByVal lot As Variant, ByVal subDiv As Variant, ByVal addr As Variant, _
                                ByVal dedDate As Variant, ByVal builder As Variant) As String
    Dim lotSub As String, a As String, dt As String, parts As String

    lotSub = UCase(Trim(Trim(CStr(lot)) & " " & Trim(CStr(subDiv))))
    a = ExpandStreet(UCase(Trim(CStr(addr))))
    If IsDate(dedDate) Then
        dt = Format(CDate(dedDate), "m-d-yyyy")
    Else
        dt = "NO DEDUCTION DATE"
    End If

    If lotSub <> "" Then parts = lotSub & ", "
    If a <> "" Then parts = parts & a & ", "
    FolderNameFrom = CleanName(parts & dt & ", " & FolderBuilder(CStr(builder)))
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
