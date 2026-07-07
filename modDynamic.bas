Attribute VB_Name = "modDynamic"
Option Explicit

' ================================================================
'  modDynamic  -  live dynamic row insertion driven by SheetChange.
'
'  When the user fills the Apt# (col B) of the second-to-last
'  buffer row in a section, a new buffer row is inserted automatically.
'
'  ThisWorkbook must contain:
'    Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'        If Not TypeOf Sh Is Worksheet Then Exit Sub
'        modDynamic.HandleSheetChange Sh, Target
'    End Sub
'
'  BUG FIX (v2.0.0): Row insertion uses InsertRowCopyFromSource
'  (modSheetUtils), which handles the unmerge/insert/format/formula
'  sequence and prevents the "merged cell" crash.
'
'  Version 2.6.0 - removed the module-level Buffer Rows cache:
'  GetBufferRows now reads PS.BufferRows fresh on every qualifying
'  change event, so Setup-sheet edits take effect immediately
'  (matching the documented behavior). The read only happens after
'  the cheap short-circuit checks in HandleSheetChange pass.
'
'  Added: live "Pending (Manual)" MTM pickup - the instant col A
'  (Renewal Status) is set to "MTM" on a month sheet inside the
'  current rolling 3-month window, HandlePendingStatusChange copies
'  that unit onto the MTM tracker's Pending section immediately. This
'  module now also depends on modMTM (GetMTMAnchorDate, AddPendingUnit)
'  in addition to modSheetUtils. Window membership is decided by
'  parsing the sheet's own name via modSheetUtils.ParseMonthSheet and
'  comparing (month, year) against the anchor window, rather than by
'  exact-string-matching canonical sheet names, so legacy-named month
'  sheets (e.g. "March 2026") are still correctly included/excluded.
' ================================================================

' ----------------------------------------------------------------
'  BUFFER ROWS  -  read fresh from the Setup sheet every time.
' ----------------------------------------------------------------
Private Function GetBufferRows() As Long
    Dim n As Long: n = 0
    On Error Resume Next
    n = CLng(ThisWorkbook.Names("PS.BufferRows").RefersToRange.Value)
    On Error GoTo 0
    If n < 1 Then n = 2
    GetBufferRows = n
End Function

' ----------------------------------------------------------------
'  CHANGE EVENT HANDLER  (called from ThisWorkbook)
' ----------------------------------------------------------------
Public Sub HandleSheetChange(Sh As Object, Target As Range)
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Dim ws As Worksheet: Set ws = Sh

    If Not IsMonthlySheet(ws.Name) Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub    ' ignore multi-cell paste

    If Target.Column = 1 Then                  ' watch col A (Renewal Status)
        HandlePendingStatusChange ws, Target   ' for the live MTM Pending pickup
        Exit Sub
    End If

    If Target.Column <> 2 Then Exit Sub        ' watch col B (Apt#) only
    If Target.Row <= 2 Then Exit Sub
    If IsEmpty(Target.Value) Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    Dim changedRow As Long: changedRow = Target.Row
    If IsSectionHeader(ws, changedRow) Then Exit Sub

    Dim nextBound As Long: nextBound = FindNextBoundary(ws, changedRow)
    If nextBound = 0 Then Exit Sub

    Dim bufN As Long: bufN = GetBufferRows()
    If changedRow = nextBound - bufN Then
        Application.EnableEvents = False
        On Error GoTo ReEnable
        InsertBufferRow ws, nextBound - 1
ReEnable:
        Application.EnableEvents = True
        If Err.Number <> 0 Then
            MsgBox "Row insert failed: " & Err.Description, vbExclamation, "Dynamic Row Error"
        End If
    End If
End Sub

' ----------------------------------------------------------------
'  ROW INSERTION
'  Delegates the 4-step insert/format/unmerge/formula sequence to
'  InsertRowCopyFromSource (modSheetUtils), then clears data cells.
' ----------------------------------------------------------------
Private Sub InsertBufferRow(ws As Worksheet, insertAt As Long)
    Dim copyFrom As Long: copyFrom = insertAt - 1
    InsertRowCopyFromSource ws, insertAt, copyFrom

    ' Clear unit-specific data. Cols L(12) M(13) X(24) are intentionally kept
    ' from the row above - new buffer rows inherit floor-plan averages from peers.
    ' v1.1.0 columns: A(1) B(2) C(3) D(4) E(5) F(6) K(11) N(14) P(16) T(20) U(21)
    Dim clearCols As Variant
    clearCols = Array(1, 2, 3, 4, 5, 6, 11, 14, 16, 20, 21)
    Dim col As Variant
    For Each col In clearCols
        ws.Cells(insertAt, col).ClearContents
    Next col
    ws.Cells(insertAt, 9).Value = 0    ' I: Pet Fees default
    ws.Rows(insertAt).RowHeight = 20.1
End Sub

' ----------------------------------------------------------------
'  HandlePendingStatusChange  -  live trigger: the instant col A
'  (Renewal Status) is set to "MTM" on a month sheet inside the
'  current 3-month Pending window (anchored to modMTM.GetMTMAnchorDate),
'  copy that unit onto the MTM tracker's Pending (Manual) section
'  immediately. Every other Renewal Status value (Renewed/NTV/Pending/
'  blank) is a cheap no-op. Window membership is decided by parsing
'  this sheet's own name via modSheetUtils.ParseMonthSheet and
'  comparing its (month, year) against the 3 (month, year) pairs the
'  anchor window represents - not by exact-string-matching canonical
'  sheet names - so legacy-named sheets are still correctly included/
'  excluded. Wrapped in the same EnableEvents/On Error pattern as
'  InsertBufferRow above (covering everything from resolving the
'  anchor through calling AddPendingUnit), since this writes to a
'  different sheet (the MTM tracker) than the one that raised the
'  event. Also carries the row's col S ("MTM Rate") over to
'  AddPendingUnit, which lands in Pending's New MTM Rate column.
' ----------------------------------------------------------------
Private Sub HandlePendingStatusChange(ws As Worksheet, Target As Range)
    If Target.Row <= 2 Then Exit Sub
    If IsSectionHeader(ws, Target.Row) Then Exit Sub
    If LCase(Trim(CStr(Target.Value))) <> "mtm" Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo ReEnable

    Dim curMonth As Long, curYear As Long
    If Not modSheetUtils.ParseMonthSheet(ws.Name, curMonth, curYear) Then GoTo ReEnable

    Dim anchorDate As Date: anchorDate = modMTM.GetMTMAnchorDate()
    Dim inWindow As Boolean: inWindow = False
    Dim offset As Long
    For offset = -2 To 0
        Dim wDate As Date: wDate = DateAdd("m", offset, anchorDate)
        If curMonth = Month(wDate) And curYear = Year(wDate) Then inWindow = True: Exit For
    Next offset
    If Not inWindow Then GoTo ReEnable

    Dim r As Long: r = Target.Row
    Dim unitNum As String: unitNum = Trim(CStr(ws.Cells(r, 2).Value))
    If unitNum = "" Then GoTo ReEnable

    Dim residentName As String: residentName = Trim(CStr(ws.Cells(r, 3).Value))
    Dim fpCode       As String: fpCode       = Trim(CStr(ws.Cells(r, 4).Value))
    Dim currentRent  As Variant: currentRent = ws.Cells(r, 5).Value
    Dim mtmRate      As Variant: mtmRate     = ws.Cells(r, 19).Value   ' col S: MTM Rate

    modMTM.AddPendingUnit unitNum, residentName, fpCode, currentRent, ws.Name, mtmRate
ReEnable:
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Pending MTM pickup failed: " & Err.Description, vbExclamation, "Dynamic Row Error"
    End If
End Sub
