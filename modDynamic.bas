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
'  Version 2.1.0 - refactored from modRenewalDynamic v1.2.1
' ================================================================

Private mBufCache As Long

Public Sub RefreshBufferCache()
    mBufCache = 0
End Sub

Private Function GetBufferRows() As Long
    If mBufCache < 1 Then
        On Error Resume Next
        mBufCache = CLng(ThisWorkbook.Names("PS.BufferRows").RefersToRange.Value)
        On Error GoTo 0
        If mBufCache < 1 Then mBufCache = 2
    End If
    GetBufferRows = mBufCache
End Function

' ----------------------------------------------------------------
'  CHANGE EVENT HANDLER  (called from ThisWorkbook)
' ----------------------------------------------------------------
Public Sub HandleSheetChange(Sh As Object, Target As Range)
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Dim ws As Worksheet: Set ws = Sh

    If Not IsMonthlySheet(ws.Name) Then Exit Sub
    If Target.Column <> 2 Then Exit Sub        ' watch col B (Apt#) only
    If Target.Cells.Count > 1 Then Exit Sub    ' ignore multi-cell paste
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
    ' from the row above â€” new buffer rows inherit floor-plan averages from peers.
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
