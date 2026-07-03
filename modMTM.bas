Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler and import tools.
'
'  Reads: modConfig (PropConfig, LoadConfig, BAR_GREY, GetGroupForCode)
'         modReaders (PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists, IsSectionBar, MonthSheetName)
'
'  Column layout (A-K):
'    A Unit | B Name | C Floor Plan | D Lease Expiry | E Current Rent
'    F Last Increase (manual) | G Market Rent | H Next Increase (calc)
'    I Status | J Notes | K Select (☑/☐)
'
'  Version 2.1.0
' ================================================================

Private Const MTM_SHEET  As String = "MTM"
Private Const DATA_START As Long = 3    ' row 1=title, row 2=headers, row 3+=data
Private Const COL_COUNT  As Long = 11   ' A through K

Private Const CHK_ON  As String = "☑"  ' ChrW(9745)
Private Const CHK_OFF As String = "☐"  ' ChrW(9744)

' ================================================================
'  PUBLIC — button handlers
' ================================================================

' ----------------------------------------------------------------
'  RefreshMTMSheet  -  prompts for Yardi Rent Roll and refreshes
'                      the MTM tracker. Creates the tab if missing.
' ----------------------------------------------------------------
Public Sub RefreshMTMSheet()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    If Not SheetExists(MTM_SHEET) Then
        Dim newWs As Worksheet
        Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newWs.Name = MTM_SHEET
    End If

    Dim yardiPath As String
    yardiPath = PickFile("Select Yardi Rent Roll for MTM Refresh", "xlsx")
    If yardiPath = "" Then Exit Sub

    Dim yardiWB As Workbook
    On Error GoTo ErrHandler
    Set yardiWB = Workbooks.Open(yardiPath, ReadOnly:=True, UpdateLinks:=False)

    Dim mtmDict As Object
    Set mtmDict = ReadYardiMTM(cfg, yardiWB)
    yardiWB.Close False
    Set yardiWB = Nothing

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    FormatMTMSheet ws, cfg
    WriteHeaders ws, cfg
    DoRefreshMTM ws, mtmDict

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "MTM tracker refreshed. " & mtmDict.Count & " active MTM unit(s) found.", _
           vbInformation, "MTM Refresh"
    Exit Sub

ErrHandler:
    If Not yardiWB Is Nothing Then yardiWB.Close False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "MTM Refresh error: " & Err.Description, vbCritical, "MTM Refresh"
End Sub

' ----------------------------------------------------------------
'  ImportSelectedMTM  -  places ☑ rows into the correct month
'                        sheet based on Next Increase date.
'                        Run the monthly import afterwards to fill
'                        in rent data via FillMTMRows.
' ----------------------------------------------------------------
Public Sub ImportSelectedMTM()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    If Not SheetExists(MTM_SHEET) Then
        MsgBox "MTM tracker sheet not found.", vbExclamation, "Import MTM"
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then
        MsgBox "No MTM units to import.", vbInformation, "Import MTM"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim imported As Long: imported = 0
    Dim skipped As String: skipped = ""
    Dim r As Long

    For r = DATA_START To lastRow
        If Trim(CStr(ws.Cells(r, 11).Value)) <> CHK_ON Then GoTo NextImportRow

        Dim unitNum As String: unitNum = Trim(CStr(ws.Cells(r, 1).Value))
        Dim fpCode  As String: fpCode  = Trim(CStr(ws.Cells(r, 3).Value))
        Dim niCell  As Variant: niCell = ws.Cells(r, 8).Value

        If unitNum = "" Then GoTo NextImportRow

        If Not IsDate(niCell) Then
            skipped = skipped & "  " & unitNum & " — no Next Increase date" & vbCrLf
            GoTo NextImportRow
        End If

        Dim mNum    As Long: mNum = Month(CDate(niCell))
        Dim yr      As Long: yr   = Year(CDate(niCell))
        Dim shName  As String: shName = MonthSheetName(mNum, yr)

        If Not SheetExists(shName) Then
            skipped = skipped & "  " & unitNum & " — sheet '" & shName & "' not found" & vbCrLf
            GoTo NextImportRow
        End If

        Dim grp As String: grp = GetGroupForCode(cfg, fpCode)
        If grp = "" Then
            skipped = skipped & "  " & unitNum & " — floor plan '" & fpCode & "' not in config" & vbCrLf
            GoTo NextImportRow
        End If

        Dim mws As Worksheet: Set mws = ThisWorkbook.Sheets(shName)
        Dim placed As Boolean: placed = PlaceUnitInSection(mws, unitNum, grp)

        If placed Then
            ws.Cells(r, 11).Value = CHK_OFF
            imported = imported + 1
        Else
            skipped = skipped & "  " & unitNum & " — already on " & shName & vbCrLf
            ws.Cells(r, 11).Value = CHK_OFF
        End If

NextImportRow:
    Next r

    Application.ScreenUpdating = True

    Dim msg As String
    msg = imported & " unit(s) placed in month sheet(s)."
    If skipped <> "" Then msg = msg & vbCrLf & vbCrLf & "Skipped:" & vbCrLf & skipped
    msg = msg & vbCrLf & vbCrLf & "Run the monthly import to fill in rent and other data."
    MsgBox msg, vbInformation, "Import MTM"
End Sub

' ----------------------------------------------------------------
'  SelectAllMTM  -  checks every unit in the tracker.
' ----------------------------------------------------------------
Public Sub SelectAllMTM()
    If Not SheetExists(MTM_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then Exit Sub
    ws.Range(ws.Cells(DATA_START, 11), ws.Cells(lastRow, 11)).Value = CHK_ON
End Sub

' ----------------------------------------------------------------
'  ClearSelectionMTM  -  unchecks every unit in the tracker.
' ----------------------------------------------------------------
Public Sub ClearSelectionMTM()
    If Not SheetExists(MTM_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then Exit Sub
    ws.Range(ws.Cells(DATA_START, 11), ws.Cells(lastRow, 11)).Value = CHK_OFF
End Sub

' ================================================================
'  PRIVATE — formatting, refresh logic, helpers
' ================================================================

' ----------------------------------------------------------------
'  FormatMTMSheet  -  monthly-sheet-style formatting.
' ----------------------------------------------------------------
Private Sub FormatMTMSheet(ws As Worksheet, cfg As PropConfig)
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    ws.Columns("A").ColumnWidth = 9.14
    ws.Columns("B").ColumnWidth = 21.57
    ws.Columns("C").ColumnWidth = 12.29
    ws.Columns("D").ColumnWidth = 13#
    ws.Columns("E").ColumnWidth = 13.86
    ws.Columns("F").ColumnWidth = 13#
    ws.Columns("G").ColumnWidth = 11.43
    ws.Columns("H").ColumnWidth = 13#
    ws.Columns("I").ColumnWidth = 20#
    ws.Columns("J").ColumnWidth = 48.71
    ws.Columns("K").ColumnWidth = 5#

    ws.Rows(1).RowHeight = 32.25
    ws.Rows(2).RowHeight = 90.75

    ws.Range("D3:D2000").NumberFormat = "mm/dd/yy;@"
    ws.Range("E3:E2000").NumberFormat = "$#,##0"
    ws.Range("F3:F2000").NumberFormat = "mm/dd/yy;@"
    ws.Range("G3:G2000").NumberFormat = "$#,##0"
    ws.Range("H3:H2000").NumberFormat = "mm/dd/yy;@"

    With ws.Range("A3:K2000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("J3:J2000").HorizontalAlignment = xlLeft

    ' Select column dropdown: ☑ or ☐
    With ws.Range("K3:K2000")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                        Formula1:=CHK_ON & "," & CHK_OFF
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With

    With ws.Range("A2:K2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ws.Range("A3:K2000").FormatConditions.Delete

    Dim fcG As Object
    Set fcG = ws.Range("A3:K2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$I3=""Active MTM""")
    fcG.Interior.Color = RGB(226, 239, 218)

    Dim fcR As Object
    Set fcR = ws.Range("A3:K2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=LEFT($I3,1)=""" & ChrW(9888) & """")
    fcR.Interior.Color = RGB(252, 220, 220)

    With ws.Parent.Windows(1)
        .FreezePanes = False
    End With
    ws.Cells(DATA_START, 2).Select
    ws.Parent.Windows(1).FreezePanes = True
    ws.Parent.Windows(1).DisplayGridlines = False
End Sub

' ----------------------------------------------------------------
'  WriteHeaders  -  title bar (row 1) and column headers (row 2).
' ----------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet, cfg As PropConfig)
    ws.Range("A1:K1").Merge
    With ws.Range("A1")
        .Value = cfg.ShortName & " - MTM Tracker"
        .Font.Name = "Garamond"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim headers As Variant
    headers = Array("Unit", "Name", "Floor Plan", "Lease Expiry", "Current Rent", _
                    "Last Increase", "Market Rent", "Next Increase", "Status", "Notes", "Select")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i

    With ws.Range("A2:K2")
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = BAR_GREY
    End With
End Sub

' ----------------------------------------------------------------
'  DoRefreshMTM  -  update existing rows, flag gone units, add new.
'
'  Cols F (Last Increase), J (Notes), K (Select) are NEVER
'  overwritten — they are preserved across every refresh.
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then lastRow = DATA_START - 1

    Dim sheetRows As Object: Set sheetRows = CreateObject("Scripting.Dictionary")
    sheetRows.CompareMode = 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u <> "" Then sheetRows(u) = r
    Next r

    Dim newUnits As Object: Set newUnits = CreateObject("Scripting.Dictionary")
    newUnits.CompareMode = 1
    Dim key As Variant
    For Each key In mtmDict.Keys
        Dim arr As Variant: arr = mtmDict(key)
        If sheetRows.Exists(key) Then
            r = sheetRows(key)
            ws.Cells(r, 2).Value = CStr(arr(0))
            ws.Cells(r, 3).Value = CStr(arr(1))
            If IsDate(arr(4)) Then ws.Cells(r, 4).Value = CDate(arr(4))
            ws.Cells(r, 5).Value = CDbl(arr(3))
            ' col 6 (F Last Increase) — manual, never touched
            If IsNumeric(arr(2)) Then ws.Cells(r, 7).Value = CDbl(arr(2))
            Dim liVal As Variant: liVal = ws.Cells(r, 6).Value
            If IsDate(liVal) Then ws.Cells(r, 8).Value = NextIncreaseDate(CDate(liVal))
            ' cols 9 (I Status), 10 (J Notes), 11 (K Select) — preserved
        Else
            newUnits(key) = arr
        End If
    Next key

    For Each key In sheetRows.Keys
        If Not mtmDict.Exists(key) Then
            If Trim(CStr(ws.Cells(sheetRows(key), 9).Value)) = "Active MTM" Then
                ws.Cells(sheetRows(key), 9).Value = ChrW(9888) & " Review - may have renewed"
            End If
        End If
    Next key

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < DATA_START Then nextRow = DATA_START
    For Each key In newUnits.Keys
        arr = newUnits(key)
        ws.Cells(nextRow, 1).Value  = CStr(key)
        ws.Cells(nextRow, 2).Value  = CStr(arr(0))
        ws.Cells(nextRow, 3).Value  = CStr(arr(1))
        If IsDate(arr(4)) Then ws.Cells(nextRow, 4).Value = CDate(arr(4))
        ws.Cells(nextRow, 5).Value  = CDbl(arr(3))
        If IsNumeric(arr(2)) Then ws.Cells(nextRow, 7).Value = CDbl(arr(2))
        ws.Cells(nextRow, 9).Value  = "Active MTM"
        ws.Cells(nextRow, 11).Value = CHK_OFF
        nextRow = nextRow + 1
    Next key

    Dim lastRowAfter As Long
    lastRowAfter = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowAfter >= DATA_START Then
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRowAfter, COL_COUNT)).Sort _
            Key1:=ws.Cells(DATA_START, 8), Order1:=xlAscending, Header:=xlNo
    End If
End Sub

' ----------------------------------------------------------------
'  PlaceUnitInSection  -  finds the matching floor plan section on
'                         the month sheet and writes the apt# in the
'                         first blank col-B row (inserts if needed).
'                         Returns False if unit was already present.
' ----------------------------------------------------------------
Private Function PlaceUnitInSection(mws As Worksheet, unitNum As String, grp As String) As Boolean
    PlaceUnitInSection = False
    Dim lastUsed As Long: lastUsed = mws.UsedRange.Row + mws.UsedRange.Rows.Count - 1

    Dim secFirst As Long: secFirst = 0
    Dim secLast  As Long: secLast  = 0
    Dim inTarget As Boolean: inTarget = False
    Dim r As Long

    For r = 3 To lastUsed
        Dim aVal As String: aVal = Trim(CStr(mws.Cells(r, 1).Value))
        Dim dVal As String: dVal = Trim(CStr(mws.Cells(r, 4).Value))

        If InStr(1, dVal, "Total", vbTextCompare) > 0 Or _
           InStr(1, aVal, "Total", vbTextCompare) > 0 Then
            If inTarget Then secLast = r - 1
            Exit For
        End If

        If IsSectionBar(mws, r) Then
            If inTarget Then
                secLast = r - 1
                Exit For
            End If
            If LCase(aVal) = LCase(grp) Then
                inTarget = True
                secFirst = r + 1
            End If
        End If
    Next r
    If inTarget And secLast = 0 Then secLast = lastUsed

    If secFirst = 0 Then Exit Function  ' section not found

    ' Check if unit is already in the section
    For r = secFirst To secLast
        If Trim(CStr(mws.Cells(r, 2).Value)) = unitNum Then Exit Function
    Next r

    ' Find first blank col-B row in section
    Dim blankRow As Long: blankRow = 0
    For r = secFirst To secLast
        If Trim(CStr(mws.Cells(r, 2).Value)) = "" Then blankRow = r: Exit For
    Next r

    ' No blank row — insert one at bottom of section
    If blankRow = 0 Then
        blankRow = secLast + 1
        mws.Rows(blankRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    mws.Cells(blankRow, 2).Value = unitNum
    If Trim(CStr(mws.Cells(blankRow, 20).Value)) = "" Then _
        mws.Cells(blankRow, 20).Value = "MTM"

    PlaceUnitInSection = True
End Function

' ----------------------------------------------------------------
'  NextIncreaseDate  -  Last Increase + 12 months, rounded up to
'                       the 1st of the month if not already on the 1st.
' ----------------------------------------------------------------
Private Function NextIncreaseDate(lastIncrease As Date) As Date
    Dim d As Date: d = DateAdd("m", 12, lastIncrease)
    If Day(d) <> 1 Then d = DateSerial(Year(d), Month(d) + 1, 1)
    NextIncreaseDate = d
End Function
