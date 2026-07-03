Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler.
'
'  Reads: modConfig (PropConfig, LoadConfig, BAR_GREY)
'         modReaders (PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists)
'
'  Version 2.1.0
' ================================================================

Private Const MTM_SHEET As String = "MTM"
Private Const DATA_START As Long = 3    ' row 1 = title, row 2 = headers, row 3+ = data

' ----------------------------------------------------------------
'  RefreshMTMSheet  -  button handler. Prompts for Yardi Rent Roll,
'                      refreshes the MTM tracker sheet.
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
'  FormatMTMSheet  -  applies monthly-sheet-style formatting:
'                     font, column widths, row heights, borders,
'                     number formats, freeze panes, conditional
'                     formatting, no gridlines.
' ----------------------------------------------------------------
Private Sub FormatMTMSheet(ws As Worksheet, cfg As PropConfig)
    ' Sheet-wide font
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    ' Column widths
    ws.Columns("A").ColumnWidth = 9.14
    ws.Columns("B").ColumnWidth = 21.57
    ws.Columns("C").ColumnWidth = 12.29
    ws.Columns("D").ColumnWidth = 13#
    ws.Columns("E").ColumnWidth = 13.86
    ws.Columns("F").ColumnWidth = 11.43
    ws.Columns("G").ColumnWidth = 13#
    ws.Columns("H").ColumnWidth = 20#
    ws.Columns("I").ColumnWidth = 48.71

    ' Row heights
    ws.Rows(1).RowHeight = 32.25
    ws.Rows(2).RowHeight = 90.75

    ' Number formats on data columns (applied to a large range; safe on empty cells)
    ws.Range("D3:D2000").NumberFormat = "mm/dd/yy;@"
    ws.Range("E3:E2000").NumberFormat = "$#,##0"
    ws.Range("F3:F2000").NumberFormat = "$#,##0"
    ws.Range("G3:G2000").NumberFormat = "mm/dd/yy;@"

    ' Center alignment on data area
    With ws.Range("A3:I2000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ' Notes column left-aligned
    ws.Range("I3:I2000").HorizontalAlignment = xlLeft

    ' Borders on header row (medium bottom)
    With ws.Range("A2:I2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ' Conditional formatting on data block — clear first, then add
    ws.Range("A3:I2000").FormatConditions.Delete

    ' Active MTM — green
    Dim fcG As Object
    Set fcG = ws.Range("A3:I2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$H3=""Active MTM""")
    fcG.Interior.Color = RGB(226, 239, 218)

    ' Review / flagged — pink
    Dim fcR As Object
    Set fcR = ws.Range("A3:I2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=LEFT($H3,1)=""" & ChrW(9888) & """")
    fcR.Interior.Color = RGB(252, 220, 220)

    ' Freeze panes: rows 1-2 and col A
    With ws.Parent.Windows(1)
        .FreezePanes = False
    End With
    ws.Cells(DATA_START, 2).Select
    ws.Parent.Windows(1).FreezePanes = True

    ' Hide gridlines
    ws.Parent.Windows(1).DisplayGridlines = False
End Sub

' ----------------------------------------------------------------
'  WriteHeaders  -  writes title bar (row 1) and column headers
'                   (row 2) matching monthly sheet style.
' ----------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet, cfg As PropConfig)
    ' Row 1 — title bar
    ws.Range("A1:I1").Merge
    With ws.Range("A1")
        .Value = cfg.ShortName & " - MTM Tracker"
        .Font.Name = "Garamond"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Row 2 — column headers
    Dim headers As Variant
    headers = Array("Unit", "Name", "Floor Plan", "Lease Expiry", _
                    "Current Rent", "Market Rent", "Next Increase", _
                    "Status", "Notes")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i

    With ws.Range("A2:I2")
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = BAR_GREY
    End With
End Sub

' ----------------------------------------------------------------
'  DoRefreshMTM  -  core merge logic: update existing rows,
'                   flag gone units, add new units, sort.
'                   Data starts at row 3 (rows 1-2 are title/headers).
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then lastRow = DATA_START - 1

    ' Build unit -> row index map from existing data rows
    Dim sheetRows As Object: Set sheetRows = CreateObject("Scripting.Dictionary")
    sheetRows.CompareMode = 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u <> "" Then sheetRows(u) = r
    Next r

    ' Pass 1 — update existing rows; collect new units
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
            If IsNumeric(arr(2)) Then ws.Cells(r, 6).Value = CDbl(arr(2))
            If IsDate(arr(4)) Then
                Dim nd As Variant: nd = NextIncreaseDate(CDate(arr(4)))
                If Not IsNull(nd) Then ws.Cells(r, 7).Value = CDate(nd)
            End If
            ' Preserve col H (Status) and col I (Notes)
        Else
            newUnits(key) = arr
        End If
    Next key

    ' Pass 2 — flag units on sheet no longer in Yardi as MTM
    For Each key In sheetRows.Keys
        If Not mtmDict.Exists(key) Then
            If Trim(CStr(ws.Cells(sheetRows(key), 8).Value)) = "Active MTM" Then
                ws.Cells(sheetRows(key), 8).Value = ChrW(9888) & " Review - may have renewed"
            End If
        End If
    Next key

    ' Pass 3 — append new units
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < DATA_START Then nextRow = DATA_START
    For Each key In newUnits.Keys
        arr = newUnits(key)
        ws.Cells(nextRow, 1).Value = CStr(key)
        ws.Cells(nextRow, 2).Value = CStr(arr(0))
        ws.Cells(nextRow, 3).Value = CStr(arr(1))
        If IsDate(arr(4)) Then ws.Cells(nextRow, 4).Value = CDate(arr(4))
        ws.Cells(nextRow, 5).Value = CDbl(arr(3))
        If IsNumeric(arr(2)) Then ws.Cells(nextRow, 6).Value = CDbl(arr(2))
        If IsDate(arr(4)) Then
            Dim nd2 As Variant: nd2 = NextIncreaseDate(CDate(arr(4)))
            If Not IsNull(nd2) Then ws.Cells(nextRow, 7).Value = CDate(nd2)
        End If
        ws.Cells(nextRow, 8).Value = "Active MTM"
        nextRow = nextRow + 1
    Next key

    ' Sort data rows by col G (Next Increase) ascending; blanks go to bottom
    Dim lastRowAfter As Long
    lastRowAfter = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowAfter >= DATA_START Then
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRowAfter, 9)).Sort _
            Key1:=ws.Cells(DATA_START, 7), Order1:=xlAscending, Header:=xlNo
    End If
End Sub

' ----------------------------------------------------------------
'  NextIncreaseDate  -  next 12-month anniversary of leaseExpiry
'                       that falls after today.
' ----------------------------------------------------------------
Private Function NextIncreaseDate(leaseExpiry As Date) As Variant
    Dim n As Long: n = 1
    Dim d As Date
    Do While n <= 100
        d = DateAdd("m", n * 12, leaseExpiry)
        If d > Date Then NextIncreaseDate = d: Exit Function
        n = n + 1
    Loop
    NextIncreaseDate = Null
End Function
