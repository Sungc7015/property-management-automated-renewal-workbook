Option Explicit

' ================================================================
'  modSheetUtils  -  worksheet and range utility helpers shared
'                    by modImport, modDynamic, modSetup, modOverview.
'
'  KEY FIX (v2.0.0): The 4-step row-insert sequence (unmerge source,
'  insert, copy formats, unmerge dest, copy formulas) is consolidated
'  in InsertRowCopyFromSource. Called by modImport.FillSheet and
'  modDynamic.InsertBufferRow.
'
'  Version 2.1.0 - carved from modRenewalImporter + modRenewalDynamic
'                  + modPropertySetup v1.2.1
' ================================================================

' ----------------------------------------------------------------
'  SHEET EXISTENCE
' ----------------------------------------------------------------
Public Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Public Function NameExists(nm As String) As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ThisWorkbook.Names(nm)
    NameExists = Not (x Is Nothing)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  MONTH SHEET NAME PARSING + GENERATION
' ----------------------------------------------------------------
Public Function ParseMonthSheet(nm As String, ByRef mNum As Long, ByRef yr As Long) As Boolean
    ParseMonthSheet = False
    mNum = 0: yr = 0
    Dim s As String: s = Trim(nm)
    Dim monthPart As String: monthPart = s

    Dim sp As Long: sp = InStrRev(s, " ")
    If sp > 0 Then
        Dim tail As String: tail = Mid(s, sp + 1)
        If Len(tail) = 4 And IsNumeric(tail) Then
            yr = CLng(tail)
            monthPart = Trim(Left(s, sp - 1))
        ElseIf Len(tail) = 2 And IsNumeric(tail) Then
            yr = 2000 + CLng(tail)
            monthPart = Trim(Left(s, sp - 1))
        End If
    End If

    mNum = MonthNumberFromName(monthPart)
    ParseMonthSheet = (mNum > 0)
    If Not ParseMonthSheet Then yr = 0
End Function

Public Function MonthNumberFromName(s As String) As Long
    Dim full As Variant, abbr As Variant
    full = Array("january", "february", "march", "april", "may", "june", _
                 "july", "august", "september", "october", "november", "december")
    abbr = Array("jan", "feb", "mar", "apr", "may", "jun", _
                 "jul", "aug", "sep", "oct", "nov", "dec")
    Dim t As String: t = LCase(Trim(s))
    Dim i As Long
    For i = 0 To 11
        If t = full(i) Or t = abbr(i) Then MonthNumberFromName = i + 1: Exit Function
    Next i
    MonthNumberFromName = 0
End Function

Public Function MonthPrefix(ByVal mNum As Long) As String
    Dim p As Variant
    p = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    MonthPrefix = p(mNum - 1)
End Function

Public Function MonthYearPrefix(ByVal mNum As Long, ByVal yr As Long) As String
    MonthYearPrefix = MonthPrefix(mNum) & yr
End Function

Public Function MonthSheetName(ByVal mNum As Long, ByVal yr As Long) As String
    MonthSheetName = MonthPrefix(mNum) & " " & Format(yr Mod 100, "00")
End Function

' ----------------------------------------------------------------
'  SECTION / HEADER DETECTION  (used by modImport and modDynamic)
' ----------------------------------------------------------------

' Monthly sheet name check (used by modDynamic to filter change events)
Public Function IsMonthlySheet(nm As String) As Boolean
    Dim monthPart As String: monthPart = Trim(nm)
    Dim sp As Long: sp = InStrRev(monthPart, " ")
    If sp > 0 Then
        Dim tail As String: tail = Mid(monthPart, sp + 1)
        If (Len(tail) = 4 Or Len(tail) = 2) And IsNumeric(tail) Then
            monthPart = Trim(Left(monthPart, sp - 1))
        End If
    End If
    IsMonthlySheet = (MonthNumberFromName(monthPart) > 0)
End Function

' Section bar: grey RGB(217,217,217) + bold in col A
Public Function IsSectionHeader(ws As Worksheet, r As Long) As Boolean
    Dim cel As Range: Set cel = ws.Cells(r, 1)
    IsSectionHeader = (cel.Interior.Color = RGB(217, 217, 217) And cel.Font.Bold = True)
End Function

' Alias used by modImport (matches the original name in modRenewalImporter)
Public Function IsSectionBar(ws As Worksheet, r As Long) As Boolean
    IsSectionBar = IsSectionHeader(ws, r)
End Function

' Scan forward from fromRow for next section bar or Total row
Public Function FindNextBoundary(ws As Worksheet, fromRow As Long) As Long
    Const MAX_SCAN As Long = 600
    Dim lastRow As Long: lastRow = fromRow + MAX_SCAN
    Dim r As Long
    For r = fromRow + 1 To lastRow
        If IsSectionHeader(ws, r) Then FindNextBoundary = r: Exit Function

        ' Total row: "Total" text in col D (v1.1.0+ layout)
        If InStr(1, CStr(ws.Cells(r, 4).Value), "Total", vbTextCompare) > 0 Then
            FindNextBoundary = r: Exit Function
        End If

        ' Legacy fallback (pre-v1.1.0 sheets)
        If ws.Cells(r, 1).Font.Bold = True And _
           CStr(ws.Cells(r, 1).Value) <> "" And _
           CStr(ws.Cells(r, 5).Value) = "" And _
           Trim(ws.Cells(r, 5).Formula) = "" Then
            FindNextBoundary = r: Exit Function
        End If

        If r > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 5 Then Exit For
    Next r
    FindNextBoundary = 0
End Function

' ----------------------------------------------------------------
'  MERGE SAFETY  (the fix lives here)
' ----------------------------------------------------------------

' Unmerge every merge on row r (cols A through AB).
' Called TWICE during row insertion:
'   1. Before Insert  - to clear multi-row merges on the source
'   2. After PasteFormats, Before PasteFormulas - to clear any
'      single-row horizontal merges that PasteFormats copied onto
'      the destination row, preventing the "can't do that to a
'      merged cell" crash on PasteFormulas.
Public Sub UnmergeMultiRowMerges(ws As Worksheet, r As Long)
    Dim c As Long
    For c = 1 To 28
        Dim cel As Range: Set cel = ws.Cells(r, c)
        If cel.MergeCells Then cel.MergeArea.UnMerge
    Next c
End Sub

' Shared 4-step row-insert helper used by modImport.FillSheet and
' modDynamic.InsertBufferRow. Unmerges sourceRow, inserts at insertAt,
' copies formats then formulas from sourceRow. If sourceRow >= insertAt
' it is auto-adjusted for the downward shift caused by the Insert.
Public Sub InsertRowCopyFromSource(ws As Worksheet, insertAt As Long, sourceRow As Long)
    UnmergeMultiRowMerges ws, sourceRow
    ws.Rows(insertAt).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Dim src As Long: src = sourceRow
    If src >= insertAt Then src = src + 1
    ws.Rows(src).Copy
    ws.Rows(insertAt).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    UnmergeMultiRowMerges ws, insertAt
    ws.Rows(src).Copy
    ws.Rows(insertAt).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Sub

' ----------------------------------------------------------------
'  DATA CELL CLEARING  (used by modImport and modDynamic)
' ----------------------------------------------------------------

' Clear import-written cells on a newly inserted row; preserve formulas.
' v1.2.0 column mapping:
'   A(1) B(2) C(3) D(4) E(5) F(6) I(9) K(11) L(12) M(13)
'   N(14) P(16) T(20) U(21) X(24)
Public Sub ClearDataCells(ws As Worksheet, r As Long)
    Dim dataCols As Variant
    dataCols = Array(1, 2, 3, 4, 5, 6, 9, 11, 12, 13, 14, 16, 20, 21, 24)
    Dim c As Variant
    For Each c In dataCols
        ws.Cells(r, c).ClearContents
    Next c
    ws.Cells(r, 9).Value = 0   ' I: Pet Fees default
End Sub
