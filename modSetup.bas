Attribute VB_Name = "modSetup"
Option Explicit

' ================================================================
'  modSetup  -  Property Setup sheet creation and month sheet
'               generation.
'
'  Version 2.1.0 - carved from modPropertySetup v1.2.1
' ================================================================

' ================================================================
'  CREATE SETUP SHEET
' ================================================================
Public Sub CreateSetupSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SETUP_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If MsgBox("A '" & SETUP_SHEET & "' sheet already exists." & vbCrLf & _
                  "Replace it? (Your current setup values will be lost.)", _
                  vbYesNo + vbExclamation, "Property Setup") <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SETUP_SHEET

    With ws.Range("A1:H1")
        .Merge
        .Value = "PROPERTY SETUP"
        .Font.Name = "Garamond": .Font.Size = 16: .Font.Bold = True
        .Interior.Color = BAR_GREY
        .HorizontalAlignment = xlCenter
    End With

    Dim lbls As Variant, vals As Variant
    lbls = Array("Property Full Name (title bar)", "Property Short Name (col F header)", _
                 "Workbook Year", "Unit Number Pattern(s)  -  N=digit, A=letter, comma-separate multiples", _
                 "MTM Cap %", "MTM Cap Through Date", "Buffer Rows per Section", _
                 "# of Floor Plan Groups (info - the list below is the source of truth)")
    vals = Array("FountainGlen Laguna Niguel", "FG Laguna Niguel", 2027, "NN-NNN", _
                 0.087, DateSerial(2027, 7, 31), 2, 6)
    Dim i As Long
    For i = 0 To UBound(lbls)
        ws.Cells(3 + i, 1).Value = lbls(i)
        ws.Cells(3 + i, 1).Font.Bold = True
        ws.Cells(3 + i, 2).Value = vals(i)
        ws.Cells(3 + i, 2).Interior.Color = INPUT_FILL
    Next i
    ws.Range("B7").NumberFormat = "0.0%"
    ws.Range("B8").NumberFormat = "mm/dd/yyyy"

    ws.Range("A12").Value = "FLOOR PLAN GROUPS  (one per row, top-to-bottom = section order on the sheet)"
    ws.Range("A12").Font.Bold = True
    ws.Range("A13").Value = "#": ws.Range("B13").Value = "Group Name (becomes the grey section bar)"
    ws.Range("A13:B13").Font.Bold = True
    ws.Range("A13:B13").Interior.Color = BAR_GREY
    Dim grps As Variant
    grps = Array("1x1 Mission 613 sq. ft.", "1x1 Green 679 sq.ft.", "1x1 Villa 896 sq.ft.", _
                 "2x2 Stickley 853 sq. ft.", "2x1 Glen 1024 sq. ft.", "2x1 McIntosh 783-845 sq. ft.")
    For i = 0 To UBound(grps)
        ws.Cells(14 + i, 1).Value = i + 1
        ws.Cells(14 + i, 2).Value = grps(i)
        ws.Cells(14 + i, 2).Interior.Color = INPUT_FILL
    Next i

    ws.Range("D12").Value = "YARDI CODE MAP  (one code per row - multiple codes may point to the same group)"
    ws.Range("D12").Font.Bold = True
    ws.Range("D13").Value = "Yardi Code": ws.Range("E13").Value = "Floor Plan Group (must match list at left exactly)"
    ws.Range("D13:E13").Font.Bold = True
    ws.Range("D13:E13").Interior.Color = BAR_GREY
    Dim codes As Variant, cgrps As Variant
    codes = Array("lna1a", "lna1a1r", "lna1a1", "lna1a2", "lna1a2r", "lna1a3", "lna1a3r", _
                  "lnb2c", "lnb2cr", "lnb1b2r", "lnb1b", "lnb1br", "lnb1b1r")
    cgrps = Array(grps(0), grps(0), grps(0), grps(1), grps(1), grps(2), grps(2), _
                  grps(3), grps(3), grps(4), grps(5), grps(5), grps(5))
    For i = 0 To UBound(codes)
        ws.Cells(14 + i, 4).Value = codes(i)
        ws.Cells(14 + i, 5).Value = cgrps(i)
        ws.Range(ws.Cells(14 + i, 4), ws.Cells(14 + i, 5)).Interior.Color = INPUT_FILL
    Next i

    ws.Range("G12").Value = "SOURCE FALLBACK COLUMNS  (used only when header search fails - column numbers)"
    ws.Range("G12").Font.Bold = True
    Dim fl As Variant, fv As Variant
    fl = Array("Rent Roll: Unit Col", "Rent Roll: Unit Type Col", "Rent Roll: Resident Col", _
               "Rent Roll: Market Rent Col", "Rent Roll: Actual Rent Col", "Rent Roll: Lease Expiry Col", _
               "Rents Grid: Unit Col", "Rents Grid: Cur Eff Rent Col", "Rents Grid: Best Offer Col", _
               "Rents Grid: Best Term Col", "Rents Grid: New Lease Col", _
               "Box Score: Unit Col", "Box Score: Unit Type Col", "Box Score: Rent Col", "Box Score: Move-In Col")
    fv = Array(1, 2, 5, 6, 7, 11, 5, 8, 10, 11, 14, 2, 3, 7, 9)
    For i = 0 To UBound(fl)
        ws.Cells(14 + i, 7).Value = fl(i)
        ws.Cells(14 + i, 8).Value = fv(i)
        ws.Cells(14 + i, 8).Interior.Color = INPUT_FILL
    Next i

    ws.Columns("A").ColumnWidth = 9:  ws.Columns("B").ColumnWidth = 42
    ws.Columns("D").ColumnWidth = 14: ws.Columns("E").ColumnWidth = 42
    ws.Columns("G").ColumnWidth = 30: ws.Columns("H").ColumnWidth = 8

    AddName "PS.PropertyName", ws.Range("B3")
    AddName "PS.ShortName", ws.Range("B4")
    AddName "PS.Year", ws.Range("B5")
    AddName "PS.UnitPatterns", ws.Range("B6")
    AddName "PS.MTMCap", ws.Range("B7")
    AddName "PS.MTMThrough", ws.Range("B8")
    AddName "PS.BufferRows", ws.Range("B9")
    AddName "PS.GroupCountCell", ws.Range("B10")
    AddName "PS.GroupsTop", ws.Range("B14")
    AddName "PS.CodesTop", ws.Range("D14")
    AddName "PS.FallbacksTop", ws.Range("G14")

    With ws.Range("A31")
        .Value = "Property Renewal Workbook System    -    version " & VER & _
                 "    -    created by Christopher Sung"
        .Font.Size = 8: .Font.Italic = True: .Font.Color = RGB(150, 150, 150)
    End With

    MsgBox "Property Setup sheet created (pre-filled with FGLN as the example)." & vbCrLf & vbCrLf & _
           "Edit the values for your property, then run GenerateMonthSheets.", _
           vbInformation, "Property Setup"
End Sub

Private Sub AddName(nm As String, rng As Range)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & rng.Worksheet.Name & "'!" & rng.Address
End Sub

' ================================================================
'  GENERATE MONTH SHEETS
' ================================================================
Public Sub GenerateMonthSheets()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    Dim s As String
    s = InputBox("Which month(s) to generate?" & vbCrLf & vbCrLf & _
                 "  3        one month" & vbCrLf & _
                 "  1-6      a range" & vbCrLf & _
                 "  1,4,7    a list" & vbCrLf & _
                 "  ALL      all twelve", "Generate Month Sheets", Month(Now))
    If Trim(s) = "" Then Exit Sub

    Dim months() As Long, mCount As Long
    If Not ParseMonths(s, months, mCount) Then
        MsgBox "Couldn't read '" & s & "'. Use a number 1-12, a range like 1-6," & vbCrLf & _
               "a list like 1,4,7, or ALL.", vbExclamation, "Generate Month Sheets"
        Exit Sub
    End If

    Dim yStr As String
    yStr = InputBox("Which year?", "Generate Month Sheets", cfg.yr)
    If Trim(yStr) = "" Then Exit Sub
    If Not IsNumeric(yStr) Then MsgBox "Enter a 4-digit year.", vbExclamation: Exit Sub
    Dim yr As Long: yr = CLng(yStr)
    If yr < 2000 Or yr > 2100 Then MsgBox "Enter a 4-digit year.", vbExclamation: Exit Sub

    Dim rowsPerSection As Long: rowsPerSection = cfg.BufferRows
    Dim rStr As String
    rStr = InputBox("Empty rows per section?" & vbCrLf & vbCrLf & _
                    "Leave blank for the normal " & cfg.BufferRows & "." & vbCrLf & _
                    "For PAST months you'll paste a full roster into, enter a" & vbCrLf & _
                    "bigger number (e.g. 50) so the whole month fits in one paste.", _
                    "Generate Month Sheets", "")
    If Trim(rStr) <> "" Then
        If Not IsNumeric(rStr) Then MsgBox "Enter a whole number of rows.", vbExclamation: Exit Sub
        rowsPerSection = CLng(rStr)
        If rowsPerSection < 1 Then rowsPerSection = cfg.BufferRows
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrHandler

    Dim built As String, skipped As String
    Dim i As Long
    For i = 0 To mCount - 1
        Dim shName As String: shName = MonthSheetName(months(i), yr)
        If SheetExists(shName) Then
            Application.ScreenUpdating = True
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Sheet '" & shName & "' already exists." & vbCrLf & _
                         "Replace it? ALL DATA ON IT WILL BE LOST.", _
                         vbYesNo + vbExclamation, "Generate Month Sheets")
            Application.ScreenUpdating = False
            If ans <> vbYes Then skipped = skipped & "  " & shName & " (kept existing)" & vbCrLf: GoTo NextMonth
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(shName).Delete
            Application.DisplayAlerts = True
        End If
        BuildMonthSheet cfg, months(i), yr, rowsPerSection
        built = built & "  " & shName & vbCrLf
NextMonth:
    Next i

    RefreshOverview cfg

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Dim msg As String
    If built <> "" Then msg = "Generated:" & vbCrLf & built
    If skipped <> "" Then msg = msg & vbCrLf & "Skipped:" & vbCrLf & skipped
    If msg = "" Then msg = "Nothing generated."
    MsgBox msg, vbInformation, "Generate Month Sheets"
    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Generate Month Sheets"
End Sub

Private Function ParseMonths(s As String, months() As Long, mCount As Long) As Boolean
    ParseMonths = False: mCount = 0
    ReDim months(11)
    s = UCase(Trim(s))
    Dim i As Long
    If s = "ALL" Then
        For i = 1 To 12: months(mCount) = i: mCount = mCount + 1: Next i
        ParseMonths = True: Exit Function
    End If
    If InStr(s, "-") > 0 Then
        Dim p() As String: p = Split(s, "-")
        If UBound(p) <> 1 Then Exit Function
        If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then Exit Function
        Dim a As Long, b As Long: a = CLng(p(0)): b = CLng(p(1))
        If a < 1 Or b > 12 Or a > b Then Exit Function
        For i = a To b: months(mCount) = i: mCount = mCount + 1: Next i
        ParseMonths = True: Exit Function
    End If
    If InStr(s, ",") > 0 Then
        Dim q() As String: q = Split(s, ",")
        For i = 0 To UBound(q)
            If Not IsNumeric(Trim(q(i))) Then Exit Function
            Dim m As Long: m = CLng(Trim(q(i)))
            If m < 1 Or m > 12 Then Exit Function
            months(mCount) = m: mCount = mCount + 1
        Next i
        ParseMonths = (mCount > 0): Exit Function
    End If
    If Not IsNumeric(s) Then Exit Function
    Dim one As Long: one = CLng(s)
    If one < 1 Or one > 12 Then Exit Function
    months(0) = one: mCount = 1
    ParseMonths = True
End Function

' ================================================================
'  BUILD ONE MONTH SHEET
' ================================================================
Private Sub BuildMonthSheet(cfg As PropConfig, mNum As Long, yr As Long, rowsPerSection As Long)
    Dim shName As String: shName = MonthSheetName(mNum, yr)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = shName
    PlaceMonthSheet ws, mNum, yr

    Dim g As Long, r As Long, i As Long

    Dim barRow() As Long, dFirst() As Long, dLast() As Long
    ReDim barRow(cfg.GroupCount - 1)
    ReDim dFirst(cfg.GroupCount - 1)
    ReDim dLast(cfg.GroupCount - 1)
    r = 3
    For g = 0 To cfg.GroupCount - 1
        barRow(g) = r
        dFirst(g) = r + 1
        dLast(g) = r + rowsPerSection
        r = dLast(g) + 1
    Next g
    Dim totalRow As Long:  totalRow = r
    Dim statsHdr As Long:  statsHdr = totalRow + 3
    Dim firstStat As Long: firstStat = statsHdr + 1
    Dim rRenew As Long:  rRenew = firstStat
    Dim rPct As Long:    rPct = firstStat + 1
    Dim rInc As Long:    rInc = firstStat + 2
    Dim rAvg As Long:    rAvg = firstStat + 3
    Dim rDol As Long:    rDol = firstStat + 4
    Dim statsEnd As Long
    statsEnd = rDol + 2
    If rPct + cfg.GroupCount - 1 > statsEnd Then statsEnd = rPct + cfg.GroupCount - 1

    Dim capStr As String: capStr = Trim(Str(cfg.MTMCap))

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    ' Title row
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 21))
        .Merge
        .Value = cfg.FullName & " Lease Renewal Spreadsheet " & MonthName(mNum) & " " & yr
        .Font.Name = "Garamond": .Font.Size = 14: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 32.25

    ' Header row
    ws.Cells(2, 1).Value = "Renewal Status"
    Dim hdr As Variant
    hdr = Array("Apt #", "Resident", "Floor plan", "Current Rent", "Yieldstar Rec. Increase", _
                cfg.ShortName & " Rec. Increase", "% Of Increase", "Pet Fees", _
                "Renewal Rate (not including Pet Fee)", _
                "Rent Roll Market Rent as of (xx/xx/xx)", "Occupied Rent in place blended Avg.", _
                "Recent Avg Effective Rent - Blended", "Yieldstar New Lease Rent", _
                "% Difference from New Lease Rent", "Lease End Date", "MTM $ Increase", _
                "MTM % Increase (Max of " & Format(cfg.MTMCap, "0.0%") & " through " & _
                Format(cfg.MTMThrough, "mm/dd/yyyy") & ")", "MTM Rate", "Current Term", "Notes")
    For i = 0 To 19
        ws.Cells(2, i + 2).Value = hdr(i)
    Next i
    Dim hdr2 As Variant
    hdr2 = Array("Renewal Rate", "Inplace Lease Avg", "$ Over Inplace Avg.", _
                 "% Over Inplace Avg.", "% Over Prior Rent")
    For i = 0 To 4
        ws.Cells(2, 23 + i).Value = hdr2(i)
    Next i
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 21))
        .Font.Bold = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = BAR_GREY
    End With
    With ws.Range(ws.Cells(2, 23), ws.Cells(2, 27))
        .Font.Bold = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = BAR_GREY
    End With
    ws.Rows(2).RowHeight = 90.75

    ' Number formats
    Dim dataTop As Long: dataTop = 3
    Dim dataBot As Long: dataBot = totalRow
    SetColFmt ws, "E", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "F", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "G", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "H", dataTop, dataBot, "0.00%"
    SetColFmt ws, "I", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "J", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "K", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "L", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "M", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "N", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "O", dataTop, dataBot, "0.00%"
    SetColFmt ws, "P", dataTop, dataBot, "mm/dd/yy;@"
    SetColFmt ws, "Q", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "R", dataTop, dataBot, "0.00%"
    SetColFmt ws, "S", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "T", dataTop, dataBot, "General"
    SetColFmt ws, "U", dataTop, dataBot, "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    SetColFmt ws, "W", dataTop, dataBot, "$#,##0"
    SetColFmt ws, "X", dataTop, dataBot, "$#,##0_);[Red]($#,##0)"
    SetColFmt ws, "Y", dataTop, dataBot, "$#,##0_);[Red]($#,##0)"
    SetColFmt ws, "Z", dataTop, dataBot, "0.00%"
    SetColFmt ws, "AA", dataTop, dataBot, "0.00%"

    ' Section bars and data rows
    For g = 0 To cfg.GroupCount - 1
        Dim br As Long: br = barRow(g)
        With ws.Range(ws.Cells(br, 1), ws.Cells(br, 21))
            .Merge
            .Value = cfg.GroupNames(g)
            .Font.Name = "Garamond": .Font.Size = 14: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Interior.Color = BAR_GREY
        End With
        With ws.Range(ws.Cells(br, 23), ws.Cells(br, 27))
            .Merge: .Interior.Color = BAR_GREY
        End With
        ws.Cells(br, 22).Interior.Color = BAR_GREY
        ws.Rows(br).RowHeight = 18.75

        For r = dFirst(g) To dLast(g)
            ws.Cells(r, 8).Formula = "=IF(E" & r & "="""","""",SUM(G" & r & "/E" & r & "))"
            ws.Cells(r, 9).Value = 0
            ws.Cells(r, 10).Formula = "=E" & r & "+G" & r
            ws.Cells(r, 15).Formula = "=IF(N" & r & "="""","""",SUM((N" & r & "-J" & r & ")/J" & r & "))"
            ws.Cells(r, 17).Formula = "=SUM(S" & r & "-E" & r & ")"
            ws.Cells(r, 18).Formula = "=IF(E" & r & "="""","""",SUM(Q" & r & "/E" & r & "))"
            ws.Cells(r, 19).Formula = "=SUM((E" & r & "*" & capStr & ")+E" & r & ")"
            ws.Cells(r, 23).Formula = "=J" & r
            ws.Cells(r, 25).Formula = "=IF(X" & r & "="""","""",W" & r & "-X" & r & ")"
            ws.Cells(r, 26).Formula = "=IF(Y" & r & "="""","""",SUM(Y" & r & "/W" & r & "))"
            ws.Cells(r, 27).Formula = "=H" & r
            ws.Rows(r).RowHeight = 20.1
        Next r
    Next g

    With ws.Range(ws.Cells(3, 1), ws.Cells(totalRow, 27))
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Total row
    Dim uE As String: uE = UnionRef("E", dFirst, dLast, cfg.GroupCount)
    Dim uF As String: uF = UnionRef("F", dFirst, dLast, cfg.GroupCount)
    Dim uG As String: uG = UnionRef("G", dFirst, dLast, cfg.GroupCount)
    Dim uH As String: uH = UnionRef("H", dFirst, dLast, cfg.GroupCount)
    Dim uI As String: uI = UnionRef("I", dFirst, dLast, cfg.GroupCount)
    Dim uJ As String: uJ = UnionRef("J", dFirst, dLast, cfg.GroupCount)

    ws.Cells(totalRow, 4).Value = "Total Rent:": ws.Cells(totalRow, 4).Font.Bold = True
    ws.Cells(totalRow, 5).Formula = "=SUM(" & uE & ")"
    ws.Cells(totalRow, 6).Formula = "=SUM(" & uF & ")"
    ws.Cells(totalRow, 7).Formula = "=SUM(" & uG & ")"
    ws.Cells(totalRow, 8).Formula = "=IFERROR(AVERAGE(" & uH & "),"""")"
    ws.Cells(totalRow, 9).Formula = "=SUM(" & uI & ")"
    ws.Cells(totalRow, 10).Formula = "=SUM(" & uJ & ")"
    With ws.Range(ws.Cells(totalRow, 5), ws.Cells(totalRow, 10))
        .Font.Bold = True: .Interior.Color = BAR_GREY
    End With
    ws.Cells(totalRow, 8).NumberFormat = "0.0%"
    ws.Cells(totalRow, 10).NumberFormat = "$#,##0.00"
    ws.Rows(totalRow).RowHeight = 20.1

    ' Legends
    With ws.Range(ws.Cells(totalRow, 19), ws.Cells(totalRow, 21))
        .Merge: .Value = "Orange = Difficult Location or History of Vacancy Loss"
        .Interior.Color = RGB(255, 192, 0): .Font.Bold = True
        .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(totalRow + 1, 19), ws.Cells(totalRow + 1, 21))
        .Merge: .Value = "Yellow = Property Adjustment"
        .Interior.Color = RGB(255, 255, 0): .Font.Bold = True
        .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Stats header
    With ws.Range(ws.Cells(statsHdr, 2), ws.Cells(statsHdr, 12))
        .Merge: .Value = "Monthly Summary    (grey = calculated automatically    -    gold = enter manually)"
        .Font.Bold = True: .Interior.Color = BAR_GREY
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    ws.Rows(statsHdr).RowHeight = 20.1

    ' Stats left side
    StatLabel ws, rRenew, "Total # of Renewals for Mth:"
    StatLabel ws, rPct, "% of Increases:"
    StatLabel ws, rInc, "# of Increases"
    StatLabel ws, rAvg, "Average Increase:"
    StatLabel ws, rDol, "Total $ of Increases:"

    ws.Cells(rRenew, 5).NumberFormat = "0"
    ws.Cells(rPct, 5).Formula = "=H" & totalRow
    ws.Cells(rPct, 5).NumberFormat = "0.00%"
    ws.Cells(rInc, 5).NumberFormat = "0"
    ws.Cells(rAvg, 5).Formula = "=IFERROR(G" & totalRow & "/E" & rRenew & ","""")"
    ws.Cells(rAvg, 5).NumberFormat = "$#,##0"
    ws.Cells(rDol, 5).Formula = "=G" & totalRow
    ws.Cells(rDol, 5).NumberFormat = "$#,##0"

    Dim sr As Long
    For sr = rRenew To rDol
        With ws.Cells(sr, 5)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            If sr = rRenew Or sr = rInc Then
                .Interior.Color = INPUT_FILL
            Else
                .Interior.Color = BAR_GREY
            End If
        End With
    Next sr

    ' Per-group avg increase
    For g = 0 To cfg.GroupCount - 1
        r = rPct + g
        ws.Cells(r, 6).Value = cfg.GroupNames(g) & " Avg Increase"
        ws.Cells(r, 7).Formula = "=IFERROR(AVERAGE(G" & dFirst(g) & ":G" & dLast(g) & "),"""")"
        ws.Cells(r, 7).NumberFormat = "$#,##0"
        ws.Cells(r, 7).Font.Bold = True
        ws.Cells(r, 7).Interior.Color = BAR_GREY
    Next g

    ' Stats right side
    Dim uBlock As String:     uBlock = "A" & dFirst(0) & ":A" & (totalRow - 1)
    Dim fBlock As String:     fBlock = "G" & dFirst(0) & ":G" & (totalRow - 1)
    Dim qBlock As String:     qBlock = "Q" & dFirst(0) & ":Q" & (totalRow - 1)
    Dim rMTM As Long:         rMTM = rDol
    Dim rTotalCap As Long:    rTotalCap = rDol + 1
    Dim rPotential As Long:   rPotential = rDol + 2

    StatRightLabel ws, rPct, "Signed (Renewed):"
    StatRightLabel ws, rInc, "Capture Ratio %"
    StatRightLabel ws, rAvg, "Renewal $ Increase (12-Mo Signed):"
    StatRightLabel ws, rMTM, "MTM $ Increase (Premium):"
    StatRightLabel ws, rTotalCap, "Total $ Captured (Renewals + MTM):"
    StatRightLabel ws, rPotential, "Potential $ Increase (All Units):"

    ws.Cells(rPct, 12).Formula = "=COUNTIF(" & uBlock & ",""Renewed"")"
    ws.Cells(rPct, 12).NumberFormat = "0"
    ws.Cells(rInc, 12).Formula = "=IFERROR(L" & rPct & "/E" & rRenew & ","""")"
    ws.Cells(rInc, 12).NumberFormat = "0.0%"
    ws.Cells(rAvg, 12).Formula = "=IFERROR(SUMIF(" & uBlock & ",""Renewed""," & fBlock & "),"""")"
    ws.Cells(rAvg, 12).NumberFormat = "$#,##0"
    ws.Cells(rMTM, 12).Formula = "=IFERROR(SUMIF(" & uBlock & ",""MTM""," & qBlock & "),"""")"
    ws.Cells(rMTM, 12).NumberFormat = "$#,##0"
    ws.Cells(rTotalCap, 12).Formula = "=L" & rAvg & "+L" & rMTM
    ws.Cells(rTotalCap, 12).NumberFormat = "$#,##0"
    ws.Cells(rPotential, 12).Formula = "=G" & totalRow
    ws.Cells(rPotential, 12).NumberFormat = "$#,##0"

    Dim kr As Long
    For kr = rPct To rPotential
        With ws.Cells(kr, 12)
            .Font.Bold = True: .Interior.Color = BAR_GREY
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
        End With
    Next kr

    ws.Cells(rDol, 19).Formula = "=IFERROR(G" & totalRow & "/E" & rInc & ","""")"
    ws.Cells(rDol, 19).NumberFormat = "$#,##0.00"

    For sr = rRenew To statsEnd
        ws.Rows(sr).RowHeight = 29.25
    Next sr

    ' Borders
    With ws.Range(ws.Cells(2, 1), ws.Cells(totalRow, 21)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin
    End With
    With ws.Range(ws.Cells(2, 23), ws.Cells(totalRow, 27)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin
    End With
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 21)).Borders(xlEdgeBottom).Weight = xlMedium
    For g = 0 To cfg.GroupCount - 1
        ws.Range(ws.Cells(barRow(g), 1), ws.Cells(barRow(g), 21)).Borders(xlEdgeTop).Weight = xlMedium
    Next g

    ' Column widths
    Dim wCols As Variant, wVals As Variant
    wCols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", _
                  "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA")
    wVals = Array(12#, 9.14, 21.57, 12.29, 13.86, 10.71, 14.57, 10.71, 10.71, 13.57, 11.43, _
                  10.71, 10.71, 10.71, 11.57, 13#, 13#, 13.14, 15.86, 9.14, 48.71, _
                  4#, 10.71, 10.71, 9.86, 9.14, 13.86)
    For i = 0 To UBound(wCols)
        ws.Columns(wCols(i)).ColumnWidth = wVals(i)
    Next i

    ' Renewal Status dropdown + row coloring
    Dim firstData As Long: firstData = dFirst(0)
    Dim lastData As Long:  lastData = totalRow - 1
    Dim gv As Long
    For gv = 0 To cfg.GroupCount - 1
        With ws.Range(ws.Cells(dFirst(gv), 1), ws.Cells(dLast(gv), 1)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="Renewed,MTM,NTV,Pending"
            .IgnoreBlank = True: .InCellDropdown = True
        End With
    Next gv

    Dim cfBlock As Range
    Set cfBlock = ws.Range(ws.Cells(firstData, 1), ws.Cells(lastData, 27))
    Dim fcG As FormatCondition, fcO As FormatCondition, fcP As FormatCondition
    Set fcG = cfBlock.FormatConditions.Add(Type:=xlExpression, _
              Formula1:="=$A" & firstData & "=""Renewed""")
    fcG.Interior.Color = RGB(226, 239, 218)
    Set fcO = cfBlock.FormatConditions.Add(Type:=xlExpression, _
              Formula1:="=OR($A" & firstData & "=""MTM"",$A" & firstData & "=""NTV"")")
    fcO.Interior.Color = RGB(252, 220, 220)
    Set fcP = cfBlock.FormatConditions.Add(Type:=xlExpression, _
              Formula1:="=$A" & firstData & "=""Pending""")
    fcP.Interior.Color = RGB(189, 215, 238)

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(3, 4).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    Dim p As String: p = MonthYearPrefix(mNum, yr)
    DefineMonthNames p, shName, rRenew, rPct, rInc, rAvg, rDol
End Sub

' ----------------------------------------------------------------
'  LAYOUT HELPERS (private)
' ----------------------------------------------------------------
Private Sub SetColFmt(ws As Worksheet, col As String, r1 As Long, r2 As Long, fmt As String)
    ws.Range(col & r1 & ":" & col & r2).NumberFormat = fmt
End Sub

Private Sub StatLabel(ws As Worksheet, r As Long, txt As String)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Merge: .Value = txt
        .Font.Bold = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StatRightLabel(ws As Worksheet, r As Long, txt As String)
    With ws.Range(ws.Cells(r, 8), ws.Cells(r, 11))
        .Merge: .Value = txt
        .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub

Private Function UnionRef(col As String, dFirst() As Long, dLast() As Long, n As Long) As String
    Dim g As Long, s As String
    For g = 0 To n - 1
        If s <> "" Then s = s & ","
        s = s & col & dFirst(g) & ":" & col & dLast(g)
    Next g
    UnionRef = s
End Function

Private Sub PlaceMonthSheet(ws As Worksheet, mNum As Long, yr As Long)
    Dim thisKey As Long: thisKey = yr * 100 + mNum
    Dim bestKey As Long: bestKey = -1
    Dim bestSheet As Worksheet
    Dim Sh As Worksheet, sm As Long, sy As Long
    For Each Sh In ThisWorkbook.Sheets
        If Sh.Name <> ws.Name Then
            If ParseMonthSheet(Sh.Name, sm, sy) Then
                Dim k As Long: k = sy * 100 + sm
                If k < thisKey And k > bestKey Then bestKey = k: Set bestSheet = Sh
            End If
        End If
    Next Sh
    If Not bestSheet Is Nothing Then ws.Move After:=bestSheet: Exit Sub
    Dim ov As Worksheet
    For Each ov In ThisWorkbook.Sheets
        If InStr(1, ov.Name, "Overview", vbTextCompare) > 0 Then
            ws.Move After:=ov: Exit Sub
        End If
    Next ov
    ws.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
End Sub

Private Sub DefineMonthNames(p As String, shName As String, _
                              rRenew As Long, rPct As Long, rInc As Long, _
                              rAvg As Long, rDol As Long)
    Dim i As Long, nm As Object
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        Dim kill As Boolean: kill = False
        If Left(nm.Name, Len(p) + 1) = p & "." Then kill = True
        If Not kill Then
            If InStr(nm.RefersTo, "#REF") > 0 And IsMonthPrefixedName(nm.Name) Then kill = True
        End If
        If kill Then
            On Error Resume Next: nm.Delete: On Error GoTo 0
        End If
    Next i

    Dim pre As String: pre = "='" & shName & "'!"
    ThisWorkbook.Names.Add Name:=p & ".Renewals", RefersTo:=pre & "$E$" & rRenew
    ThisWorkbook.Names.Add Name:=p & ".Increases", RefersTo:=pre & "$E$" & rInc
    ThisWorkbook.Names.Add Name:=p & ".AvgDollar", RefersTo:=pre & "$E$" & rAvg
    ThisWorkbook.Names.Add Name:=p & ".AvgPercent", RefersTo:=pre & "$E$" & rPct
    ThisWorkbook.Names.Add Name:=p & ".TotalDollar", RefersTo:=pre & "$E$" & rDol
    ThisWorkbook.Names.Add Name:=p & ".CurrentRenewed", RefersTo:=pre & "$L$" & rPct
    ThisWorkbook.Names.Add Name:=p & ".CaptureRatio", RefersTo:=pre & "$L$" & rInc
    ThisWorkbook.Names.Add Name:=p & ".SignedDollar", RefersTo:=pre & "$L$" & rAvg
    ThisWorkbook.Names.Add Name:=p & ".MTMDollar", RefersTo:=pre & "$L$" & rDol
    ThisWorkbook.Names.Add Name:=p & ".TotalCapturedDollar", RefersTo:=pre & "$L$" & (rDol + 1)
End Sub

Private Function IsMonthPrefixedName(nm As String) As Boolean
    Dim dotPos As Long: dotPos = InStr(nm, ".")
    If dotPos < 4 Then Exit Function
    Dim head As String: head = Left(nm, dotPos - 1)
    Dim mab As String:  mab = LCase(Left(head, 3))
    Dim ok As Boolean:  ok = False
    Dim a As Variant
    For Each a In Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
        If mab = a Then ok = True: Exit For
    Next a
    If Not ok Then Exit Function
    Dim rest As String: rest = Mid(head, 4)
    If rest <> "" Then
        If Not (IsNumeric(rest) And Len(rest) = 4) Then Exit Function
    End If
    IsMonthPrefixedName = True
End Function
