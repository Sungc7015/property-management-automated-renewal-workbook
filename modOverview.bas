Option Explicit

' ================================================================
'  modOverview  -  multi-year renewal summary sheet builder.
'
'  Version 2.1.0 - carved from modPropertySetup v1.2.1
' ================================================================

Public Sub CreateOverviewSheet()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    Dim existing As String: existing = FindOverviewName()
    If existing <> "" Then
        If MsgBox("An overview sheet ('" & existing & "') already exists." & vbCrLf & _
                  "Rebuild it as the multi-year summary? The monthly stats it shows" & vbCrLf & _
                  "live on the month sheets, so nothing is lost.", _
                  vbYesNo + vbExclamation, "Overview") <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    RefreshOverview cfg
    Application.ScreenUpdating = True
    MsgBox "Multi-year summary built." & vbCrLf & _
           "A new column appears for each year you generate.", _
           vbInformation, "Overview"
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Overview"
End Sub

Public Sub RefreshOverview(cfg As PropConfig)
    Dim targetName As String: targetName = FindOverviewName()
    Dim evState As Boolean: evState = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Dim exMarks As Object
    If targetName = "" Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
        ws.Name = "Renewal Overview"
        Set exMarks = CreateObject("Scripting.Dictionary")
    Else
        Set ws = ThisWorkbook.Sheets(targetName)
        Set exMarks = CaptureExcludes(ws)
        ws.Cells.UnMerge
        ws.Cells.Clear
        On Error Resume Next
        ws.Cells.FormatConditions.Delete
        On Error GoTo 0
    End If

    BuildOverview cfg, ws, exMarks
    Application.EnableEvents = evState
End Sub

Public Sub EnsureOverview(cfg As PropConfig)
    RefreshOverview cfg
End Sub

Public Function FindOverviewName() As String
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Overview", vbTextCompare) > 0 Then
            FindOverviewName = ws.Name: Exit Function
        End If
    Next ws
    FindOverviewName = ""
End Function

' ----------------------------------------------------------------
'  INTERNAL HELPERS
' ----------------------------------------------------------------
Private Function DetectYears(cfg As PropConfig) As Variant
    Dim found(2000 To 2100) As Boolean
    Dim Sh As Object, mm As Long, yy As Long
    For Each Sh In ThisWorkbook.Sheets
        If TypeOf Sh Is Worksheet Then
            If ParseMonthSheet(Sh.Name, mm, yy) Then
                If yy >= 2000 And yy <= 2100 Then found(yy) = True
            End If
        End If
    Next Sh

    Dim years() As Long, n As Long: n = 0
    Dim y As Long
    For y = 2000 To 2100
        If found(y) Then
            ReDim Preserve years(n)
            years(n) = y
            n = n + 1
        End If
    Next y
    If n = 0 Then ReDim years(0): years(0) = cfg.yr
    DetectYears = years
End Function

Private Function MetricTitles() As Variant
    MetricTitles = Array("# of Renewals", "# of Increases", "Average $ Increase", _
                         "Average % Increase", "Potential Revenue", "Signed Revenue", _
                         "# Signed", "Capture Ratio %", "MTM $ Increase", "Total $ Captured")
End Function

Private Function MetricSuffix() As Variant
    MetricSuffix = Array("Renewals", "Increases", "AvgDollar", "AvgPercent", _
                         "TotalDollar", "SignedDollar", "CurrentRenewed", "CaptureRatio", _
                         "MTMDollar", "TotalCapturedDollar")
End Function

Private Function MetricFormat() As Variant
    MetricFormat = Array("0", "0", "$#,##0", "0.0%", "$#,##0", "$#,##0", "0", "0.0%", _
                         "$#,##0", "$#,##0")
End Function

Private Function SuffixCol(suffixes As Variant, nm As String) As Long
    Dim i As Long
    For i = 0 To UBound(suffixes)
        If suffixes(i) = nm Then SuffixCol = 2 + i: Exit Function
    Next i
    SuffixCol = 0
End Function

Private Function ColRange(ws As Worksheet, col As Long, r1 As Long, r2 As Long) As String
    ColRange = ws.Cells(r1, col).Address(False, False) & ":" & _
               ws.Cells(r2, col).Address(False, False)
End Function

Private Function CaptureExcludes(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Set CaptureExcludes = d: Exit Function

    Dim exCol As Long: exCol = 0
    Dim c As Long
    For c = 1 To 40
        If LCase(Trim(CStr(ws.Cells(2, c).Value))) = "exclude" Then exCol = c: Exit For
    Next c
    If exCol = 0 Then Set CaptureExcludes = d: Exit Function

    Dim curYear As Long: curYear = 0
    Dim r As Long
    For r = 3 To lastR
        Dim aVal As String: aVal = Trim(CStr(ws.Cells(r, 1).Value))
        If aVal = "" Then GoTo NextR
        If IsNumeric(aVal) Then
            Dim yv As Long: yv = CLng(Val(aVal))
            If yv >= 2000 And yv <= 2100 Then curYear = yv
            GoTo NextR
        End If
        Dim mNum As Long: mNum = MonthNumberFromName(aVal)
        If mNum >= 1 And mNum <= 12 And curYear > 0 Then
            If LCase(Trim(CStr(ws.Cells(r, exCol).Value))) = "x" Then
                d(MonthYearPrefix(mNum, curYear)) = "x"
            End If
        End If
NextR:
    Next r
    Set CaptureExcludes = d
End Function

Private Sub BuildOverview(cfg As PropConfig, ws As Worksheet, exMarks As Object)
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    Dim years As Variant: years = DetectYears(cfg)
    Dim yCount As Long: yCount = UBound(years) - LBound(years) + 1

    Dim titles As Variant:   titles = MetricTitles()
    Dim suffixes As Variant: suffixes = MetricSuffix()
    Dim fmts As Variant:     fmts = MetricFormat()
    Dim nMetric As Long: nMetric = UBound(titles) + 1
    Dim lastCol As Long: lastCol = 1 + nMetric
    Dim exCol As Long:   exCol = lastCol + 1
    Dim fullCol As Long: fullCol = exCol

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, fullCol))
        .Merge
        .Value = cfg.FullName & " Renewal Summary"
        .Font.Name = "Garamond": .Font.Size = 16: .Font.Bold = True
        .Interior.Color = BAR_GREY
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 28

    Dim k As Long
    ws.Cells(2, 1).Value = "Month"
    For k = 0 To nMetric - 1
        Dim hd As String: hd = titles(k)
        Select Case suffixes(k)
            Case "AvgDollar", "AvgPercent", "CaptureRatio"
                hd = hd & vbLf & "(Total = weighted avg)"
        End Select
        ws.Cells(2, 2 + k).Value = hd
    Next k
    ws.Cells(2, exCol).Value = "Exclude"
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, fullCol))
        .Font.Bold = True: .WrapText = True
        .Interior.Color = BAR_GREY
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(2).RowHeight = 52

    Dim r As Long: r = 3
    Dim yi As Long, m As Long

    For yi = 0 To yCount - 1
        Dim yr As Long: yr = years(LBound(years) + yi)

        With ws.Range(ws.Cells(r, 1), ws.Cells(r, fullCol))
            .Merge
            .Value = yr
            .Font.Name = "Garamond": .Font.Size = 13: .Font.Bold = True
            .Interior.Color = BAR_GREY
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeTop).Weight = xlMedium
        End With
        r = r + 1

        Dim dataTop As Long: dataTop = r

        For m = 1 To 12
            ws.Cells(r, 1).Value = MonthName(m)
            ws.Cells(r, 1).Font.Bold = True
            Dim p As String: p = MonthYearPrefix(m, yr)
            For k = 0 To nMetric - 1
                If NameExists(p & "." & suffixes(k)) Then
                    ws.Cells(r, 2 + k).Formula = "=IFERROR(" & p & "." & suffixes(k) & ","""")"
                End If
                ws.Cells(r, 2 + k).NumberFormat = fmts(k)
            Next k
            With ws.Cells(r, exCol)
                .Interior.Color = INPUT_FILL
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
                If Not exMarks Is Nothing Then
                    If exMarks.exists(p) Then .Value = "x"
                End If
            End With
            r = r + 1
        Next m
        Dim dataBot As Long: dataBot = r - 1

        Dim xRng As String: xRng = ColRange(ws, exCol, dataTop, dataBot)
        Dim bRng As String: bRng = ColRange(ws, SuffixCol(suffixes, "Renewals"), dataTop, dataBot)
        Dim fRng As String: fRng = ColRange(ws, SuffixCol(suffixes, "TotalDollar"), dataTop, dataBot)
        Dim gRng As String: gRng = ColRange(ws, SuffixCol(suffixes, "CurrentRenewed"), dataTop, dataBot)

        ws.Cells(r, 1).Value = "YTD Total (weighted)"
        For k = 0 To nMetric - 1
            Dim f As String
            Select Case suffixes(k)
                Case "AvgDollar"
                    f = "=IFERROR(SUMIF(" & xRng & ",""<>x""," & fRng & ")/SUMIF(" & xRng & ",""<>x""," & bRng & "),"""")"
                Case "CaptureRatio"
                    f = "=IFERROR(SUMIF(" & xRng & ",""<>x""," & gRng & ")/SUMIF(" & xRng & ",""<>x""," & bRng & "),"""")"
                Case "AvgPercent"
                    Dim num As String: num = ""
                    Dim mr As Long
                    Dim colPct As Long: colPct = SuffixCol(suffixes, "AvgPercent")
                    Dim colRen As Long: colRen = SuffixCol(suffixes, "Renewals")
                    For mr = dataTop To dataBot
                        Dim xA As String: xA = ws.Cells(mr, exCol).Address(False, False)
                        Dim eA As String: eA = ws.Cells(mr, colPct).Address(False, False)
                        Dim bA As String: bA = ws.Cells(mr, colRen).Address(False, False)
                        If num <> "" Then num = num & "+"
                        num = num & "IF(" & xA & "=""x"",0,IFERROR(" & eA & "*" & bA & ",0))"
                    Next mr
                    f = "=IFERROR((" & num & ")/SUMIF(" & xRng & ",""<>x""," & bRng & "),"""")"
                Case Else
                    f = "=IFERROR(SUMIF(" & xRng & ",""<>x""," & ColRange(ws, 2 + k, dataTop, dataBot) & "),"""")"
            End Select
            ws.Cells(r, 2 + k).Formula = f
            ws.Cells(r, 2 + k).NumberFormat = fmts(k)
        Next k
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, fullCol))
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)
        End With

        Dim cfRange As Range
        Set cfRange = ws.Range(ws.Cells(dataTop, 1), ws.Cells(dataBot, lastCol))
        Dim fc As FormatCondition
        Set fc = cfRange.FormatConditions.Add(Type:=xlExpression, _
                 Formula1:="=" & ws.Cells(dataTop, exCol).Address(False, True) & "=""x""")
        fc.Font.Strikethrough = True
        fc.Font.Color = RGB(150, 150, 150)
        fc.Interior.Color = RGB(228, 228, 228)

        With ws.Range(ws.Cells(dataTop, 1), ws.Cells(r, fullCol)).Borders
            .LineStyle = xlContinuous: .Weight = xlThin
        End With

        r = r + 2
    Next yi

    ws.Range(ws.Cells(3, 2), ws.Cells(r, fullCol)).HorizontalAlignment = xlCenter
    ws.Columns(1).ColumnWidth = 16
    For k = 0 To nMetric - 1
        ws.Columns(2 + k).ColumnWidth = 15
    Next k
    ws.Columns(exCol).ColumnWidth = 9
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    With ws.Cells(r + 1, 1)
        .Value = "YTD Total rows use renewal-weighted averages for Average $, " & _
                 "Average %, and Capture Ratio - larger months count proportionally, " & _
                 "not a simple average of the monthly figures.  " & _
                 "Potential Revenue counts every renewal's increase; Signed Revenue counts " & _
                 "only units marked 'Renewed' on the month sheet (the gap is increase lost " & _
                 "to MTM / NTV).  " & _
                 "MTM $ Increase is the premium captured from units that went month-to-month " & _
                 "instead of renewing; Total $ Captured is Signed Revenue plus MTM $ Increase - " & _
                 "what the property actually picked up for the month across both paths.  " & _
                 "Type 'x' in the Exclude column to leave an in-process month out of " & _
                 "all YTD totals (the row greys out); clear it to include the month again."
        .Font.Size = 9: .Font.Italic = True: .Font.Color = RGB(120, 120, 120)
        .HorizontalAlignment = xlLeft
    End With

    With ws.Cells(r + 2, 1)
        .Value = "Workbook created by Christopher Sung"
        .Font.Size = 8: .Font.Italic = True: .Font.Color = RGB(166, 166, 166)
        .HorizontalAlignment = xlLeft
    End With
End Sub
