Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler.
'
'  Reads: modConfig (PropConfig, LoadConfig, BAR_GREY)
'         modReaders (PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists)
'
'  Column layout (A-J):
'    A Unit | B Name | C Floor Plan | D Lease Expiry | E Current Rent
'    F Last Increase (manual) | G Market Rent | H Next Increase (calc)
'    I Status | J Notes
'
'  Version 2.1.0
' ================================================================

Private Const MTM_SHEET  As String = "MTM"
Private Const DATA_START As Long = 3    ' row 1=title, row 2=headers, row 3+=data
Private Const COL_COUNT  As Long = 10   ' A through J

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
'  FormatMTMSheet  -  applies monthly-sheet-style formatting.
' ----------------------------------------------------------------
Private Sub FormatMTMSheet(ws As Worksheet, cfg As PropConfig)
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    ' Column widths
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

    ' Row heights
    ws.Rows(1).RowHeight = 32.25
    ws.Rows(2).RowHeight = 90.75

    ' Number formats
    ws.Range("D3:D2000").NumberFormat = "mm/dd/yy;@"
    ws.Range("E3:E2000").NumberFormat = "$#,##0"
    ws.Range("F3:F2000").NumberFormat = "mm/dd/yy;@"
    ws.Range("G3:G2000").NumberFormat = "$#,##0"
    ws.Range("H3:H2000").NumberFormat = "mm/dd/yy;@"

    ' Alignment
    With ws.Range("A3:J2000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("J3:J2000").HorizontalAlignment = xlLeft

    ' Header row border
    With ws.Range("A2:J2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ' Conditional formatting
    ws.Range("A3:J2000").FormatConditions.Delete

    Dim fcG As Object
    Set fcG = ws.Range("A3:J2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$I3=""Active MTM""")
    fcG.Interior.Color = RGB(226, 239, 218)

    Dim fcR As Object
    Set fcR = ws.Range("A3:J2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=LEFT($I3,1)=""" & ChrW(9888) & """")
    fcR.Interior.Color = RGB(252, 220, 220)

    ' Freeze rows 1-2 and col A
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
    ws.Range("A1:J1").Merge
    With ws.Range("A1")
        .Value = cfg.ShortName & " - MTM Tracker"
        .Font.Name = "Garamond"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim headers As Variant
    headers = Array("Unit", "Name", "Floor Plan", "Lease Expiry", _
                    "Current Rent", "Last Increase", "Market Rent", _
                    "Next Increase", "Status", "Notes")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i

    With ws.Range("A2:J2")
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
'  Col F (Last Increase) is MANUAL — never overwritten by refresh.
'  Col H (Next Increase) is recalculated from col F on every refresh.
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then lastRow = DATA_START - 1

    ' Build unit -> row map from existing data
    Dim sheetRows As Object: Set sheetRows = CreateObject("Scripting.Dictionary")
    sheetRows.CompareMode = 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u <> "" Then sheetRows(u) = r
    Next r

    ' Pass 1 — update Yardi fields; recalculate Next Increase from Last Increase
    Dim newUnits As Object: Set newUnits = CreateObject("Scripting.Dictionary")
    newUnits.CompareMode = 1
    Dim key As Variant
    For Each key In mtmDict.Keys
        Dim arr As Variant: arr = mtmDict(key)
        ' arr: (0)name (1)fpCode (2)marketRent (3)actualRent (4)expiryVal
        If sheetRows.Exists(key) Then
            r = sheetRows(key)
            ws.Cells(r, 2).Value = CStr(arr(0))          ' B Name
            ws.Cells(r, 3).Value = CStr(arr(1))          ' C Floor Plan
            If IsDate(arr(4)) Then _
                ws.Cells(r, 4).Value = CDate(arr(4))     ' D Lease Expiry
            ws.Cells(r, 5).Value = CDbl(arr(3))          ' E Current Rent
            ' F Last Increase — manual, never touched
            If IsNumeric(arr(2)) Then _
                ws.Cells(r, 7).Value = CDbl(arr(2))      ' G Market Rent
            ' H Next Increase — recalculate from F (Last Increase)
            Dim liVal As Variant: liVal = ws.Cells(r, 6).Value
            If IsDate(liVal) Then _
                ws.Cells(r, 8).Value = NextIncreaseDate(CDate(liVal))
            ' I Status and J Notes — preserved
        Else
            newUnits(key) = arr
        End If
    Next key

    ' Pass 2 — flag units no longer in Yardi as MTM (only if currently "Active MTM")
    For Each key In sheetRows.Keys
        If Not mtmDict.Exists(key) Then
            If Trim(CStr(ws.Cells(sheetRows(key), 9).Value)) = "Active MTM" Then
                ws.Cells(sheetRows(key), 9).Value = ChrW(9888) & " Review - may have renewed"
            End If
        End If
    Next key

    ' Pass 3 — append new units (Last Increase and Next Increase left blank for user to fill)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < DATA_START Then nextRow = DATA_START
    For Each key In newUnits.Keys
        arr = newUnits(key)
        ws.Cells(nextRow, 1).Value = CStr(key)           ' A Unit
        ws.Cells(nextRow, 2).Value = CStr(arr(0))        ' B Name
        ws.Cells(nextRow, 3).Value = CStr(arr(1))        ' C Floor Plan
        If IsDate(arr(4)) Then _
            ws.Cells(nextRow, 4).Value = CDate(arr(4))   ' D Lease Expiry
        ws.Cells(nextRow, 5).Value = CDbl(arr(3))        ' E Current Rent
        ' F Last Increase — blank, user fills in
        If IsNumeric(arr(2)) Then _
            ws.Cells(nextRow, 7).Value = CDbl(arr(2))    ' G Market Rent
        ' H Next Increase — blank until user fills in Last Increase and re-refreshes
        ws.Cells(nextRow, 9).Value = "Active MTM"        ' I Status
        nextRow = nextRow + 1
    Next key

    ' Sort by col H (Next Increase) ascending; blanks go to bottom
    Dim lastRowAfter As Long
    lastRowAfter = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowAfter >= DATA_START Then
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRowAfter, COL_COUNT)).Sort _
            Key1:=ws.Cells(DATA_START, 8), Order1:=xlAscending, Header:=xlNo
    End If
End Sub

' ----------------------------------------------------------------
'  NextIncreaseDate  -  Last Increase + 12 months, rounded up to
'                       the 1st of the month if not already on the 1st.
' ----------------------------------------------------------------
Private Function NextIncreaseDate(lastIncrease As Date) As Date
    Dim d As Date: d = DateAdd("m", 12, lastIncrease)
    If Day(d) <> 1 Then d = DateSerial(Year(d), Month(d) + 1, 1)
    NextIncreaseDate = d
End Function
