Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler and import tools.
'
'  Reads: modConfig (PropConfig, LoadConfig, BAR_GREY, GetGroupForCode)
'         modReaders (MTMUnitRec, PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists, IsSectionBar, MonthSheetName)
'         modAdmin (AddButton - shared button-drawing helper)
'
'  Column layout (A-K):
'    A Unit | B Name | C Floor Plan | D Lease Expiry | E Current Rent
'    F Last Increase (manual) | G Market Rent | H Next Increase (calc)
'    I Status | J Notes | K [checkbox - linked to cell, TRUE/FALSE]
'
'  The tracker is a single flat list, sorted by col H (Next Increase),
'  built entirely from the Yardi Rent Roll (no RealPage or Resident
'  Lease Expirations inputs). Status (col I) is a purely informational
'  "Eligible"/"Not Yet Eligible" label - it has no effect on checkbox
'  availability or import behavior. Next Increase = Last Increase + 1
'  year + 1 day. If Last Increase is blank and the unit's lease expiry
'  is more than 15 months past-due, a note is auto-written to col J
'  (never overwriting an existing manual note) prompting the CM to
'  verify the resident ledger.
'
'  Version 2.6.0 - EnsureMTMButtons added: RefreshMTMSheet now wires
'  the 4 MTM buttons automatically the first time it creates the
'  sheet, so no manual button setup is ever required (modAdmin
'  .SetupWorkbook also calls this directly if the sheet already
'  exists). WriteMTMDataRow now reads named fields off MTMUnitRec
'  (modReaders) instead of a positional Array(...).
' ================================================================

Private Const MTM_SHEET     As String = "MTM"
Private Const MTM_SHEET_OLD As String = "MTM & STL"
Private Const DATA_START    As Long = 3
Private Const COL_COUNT     As Long = 11
Private Const CB_PREFIX     As String = "mtmChk_"

' ----------------------------------------------------------------
'  MigrateMTMSheetName  -  self-healing rename: if a workbook still
'  has the sheet under its old name ("MTM & STL") and the current
'  name ("MTM") doesn't exist yet, rename the sheet object in place
'  (preserves all data/formatting - no copy). Called at the start of
'  every public entry point that references MTM_SHEET, before the
'  "does this sheet exist" checks.
' ----------------------------------------------------------------
Private Sub MigrateMTMSheetName()
    If Not SheetExists(MTM_SHEET) And SheetExists(MTM_SHEET_OLD) Then
        ThisWorkbook.Sheets(MTM_SHEET_OLD).Name = MTM_SHEET
    End If
End Sub

' ================================================================
'  PUBLIC - button handlers
' ================================================================

Public Sub RefreshMTMSheet()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    MigrateMTMSheetName

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

    Dim mtmRecs() As MTMUnitRec
    Dim mtmDict As Object
    Set mtmDict = ReadYardiMTM(cfg, yardiWB, mtmRecs)
    yardiWB.Close False
    Set yardiWB = Nothing

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    FormatMTMSheet ws, cfg
    WriteHeaders ws, cfg
    DoRefreshMTM ws, mtmDict, mtmRecs
    EnsureMTMButtons

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
'  ImportSelectedMTM  -  places checked units into the correct month
'                        sheet based on Next Increase date. Also carries
'                        over Name, Floor Plan, and Current Rent from the
'                        MTM tracker row, not just the apt#.
' ----------------------------------------------------------------
Public Sub ImportSelectedMTM()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    MigrateMTMSheetName

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

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim imported As Long: imported = 0
    Dim skipped  As String: skipped = ""
    Dim r As Long

    For r = DATA_START To lastRow
        If ws.Cells(r, 11).Value <> True Then GoTo NextImportRow

        Dim unitNum      As String:  unitNum      = Trim(CStr(ws.Cells(r, 1).Value))
        Dim residentName As String:  residentName = Trim(CStr(ws.Cells(r, 2).Value))
        Dim fpCode       As String:  fpCode       = Trim(CStr(ws.Cells(r, 3).Value))
        Dim currentRent  As Variant: currentRent  = ws.Cells(r, 5).Value
        Dim niCell  As Variant: niCell = ws.Cells(r, 8).Value

        If unitNum = "" Then GoTo NextImportRow

        If Not IsDate(niCell) Then
            skipped = skipped & "  " & unitNum & " - no Next Increase date" & vbCrLf
            GoTo NextImportRow
        End If

        Dim mNum   As Long: mNum = Month(CDate(niCell))
        Dim yr     As Long: yr   = Year(CDate(niCell))
        Dim shName As String: shName = MonthSheetName(mNum, yr)

        If Not SheetExists(shName) Then
            skipped = skipped & "  " & unitNum & " - sheet '" & shName & "' not found" & vbCrLf
            GoTo NextImportRow
        End If

        Dim grp As String: grp = GetGroupForCode(cfg, fpCode)
        If grp = "" Then
            skipped = skipped & "  " & unitNum & " - floor plan '" & fpCode & "' not in config" & vbCrLf
            GoTo NextImportRow
        End If

        Dim mws As Worksheet: Set mws = ThisWorkbook.Sheets(shName)
        PlaceUnitInSection mws, unitNum, grp, residentName, fpCode, currentRent

        ws.Cells(r, 11).Value = False   ' uncheck via linked cell
        imported = imported + 1

NextImportRow:
    Next r

    Application.ScreenUpdating = True

    Dim msg As String
    msg = imported & " unit(s) placed in month sheet(s)."
    If skipped <> "" Then msg = msg & vbCrLf & vbCrLf & "Skipped:" & vbCrLf & skipped
    msg = msg & vbCrLf & vbCrLf & "Run the monthly import to fill in rent and other data."
    MsgBox msg, vbInformation, "Import MTM"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Import MTM error: " & Err.Description, vbCritical, "Import MTM"
End Sub

Public Sub ClearSelectionMTM()
    MigrateMTMSheetName
    If Not SheetExists(MTM_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then Exit Sub

    ws.Range(ws.Cells(DATA_START, 11), ws.Cells(lastRow, 11)).Value = False
End Sub

' ----------------------------------------------------------------
'  SortMTMSheet  -  manually re-sorts the MTM tab by col H (Next
'  Increase), on demand, without running a full RefreshMTMSheet.
'  Before sorting, recomputes each row's Next Increase (H) and Status
'  (I) from its current Last Increase (F) value, so that manually
'  typing in a new Last Increase date and clicking Sort actually
'  produces the correct updated Next Increase/Status/row-position
'  (previously this only re-sorted stale existing H/I values).
' ----------------------------------------------------------------
Public Sub SortMTMSheet()
    MigrateMTMSheetName

    If Not SheetExists(MTM_SHEET) Then
        MsgBox "MTM tracker sheet not found.", vbExclamation, "Sort MTM"
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then
        MsgBox "No MTM units to sort.", vbInformation, "Sort MTM"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' Recompute H (Next Increase) and I (Status) from the current F value
    ' before sorting, so manual edits to Last Increase are reflected.
    ' targetEnd = last calendar date of the month containing today+90 days.
    Dim tDate As Date: tDate = Now + 90
    Dim targetEnd As Date: targetEnd = DateSerial(Year(tDate), Month(tDate) + 1, 1) - 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim fVal As Variant: fVal = ws.Cells(r, 6).Value
        If IsDate(fVal) Then
            ws.Cells(r, 8).Value = NextIncreaseDate(CDate(fVal))
        Else
            ws.Cells(r, 8).Value = ""
        End If
        ws.Cells(r, 9).Value = MTMStatusLabel(ws.Cells(r, 8).Value, targetEnd)
    Next r

    ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRow, COL_COUNT)).Sort _
        Key1:=ws.Cells(DATA_START, 8), Order1:=xlAscending, Header:=xlNo

    ' Re-sorting reshuffles which rows are odd/even - refresh the banding.
    ApplyRowBanding ws, lastRow

    ' Rebuild checkboxes at their new row positions
    SyncCheckboxes ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Sort MTM error: " & Err.Description, vbCritical, "Sort MTM"
End Sub

' ----------------------------------------------------------------
'  EnsureMTMButtons  -  adds the 4 MTM tracker buttons to the MTM
'  sheet if that sheet exists. Called from modAdmin.SetupWorkbook
'  (in case the MTM sheet was already created) and from the end of
'  RefreshMTMSheet (so first-run creation always wires the buttons
'  automatically, with no separate manual step ever required).
'  Returns True if the buttons were added, False if the MTM sheet
'  does not exist yet.
' ----------------------------------------------------------------
Public Function EnsureMTMButtons() As Boolean
    MigrateMTMSheetName
    If Not SheetExists(MTM_SHEET) Then
        EnsureMTMButtons = False
        Exit Function
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)

    Dim gridLeft As Single, gridTop As Single
    Dim btnW As Single, btnH As Single, gap As Single
    gridLeft = ws.Cells(2, "M").Left
    gridTop = ws.Cells(2, "M").Top
    btnW = 160: btnH = 24: gap = 10

    modAdmin.AddButton ws, "btn_RefreshMTM", "Refresh MTM Tracker", gridLeft, gridTop, "modMTM.RefreshMTMSheet"
    modAdmin.AddButton ws, "btn_ImportMTM", "Import Selected MTM", gridLeft + btnW + gap, gridTop, "modMTM.ImportSelectedMTM"
    modAdmin.AddButton ws, "btn_ClearMTM", "Clear Selection MTM", gridLeft, gridTop + btnH + gap, "modMTM.ClearSelectionMTM"
    modAdmin.AddButton ws, "btn_SortMTM", "Sort MTM Tracker", gridLeft + btnW + gap, gridTop + btnH + gap, "modMTM.SortMTMSheet"

    EnsureMTMButtons = True
End Function

' ================================================================
'  PRIVATE
' ================================================================

Private Sub FormatMTMSheet(ws As Worksheet, cfg As PropConfig)
    ws.Activate
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
    ws.Range("K3:K2000").NumberFormat = ";;;"   ' hide TRUE/FALSE - checkbox shows state

    With ws.Range("A3:K2000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("J3:J2000").HorizontalAlignment = xlLeft

    With ws.Range("A2:K2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ws.Range("A3:K2000").FormatConditions.Delete   ' clears any wide-range CF left over from a previously-shipped version of this sheet
    ws.Range("I3:I2000").FormatConditions.Delete   ' idempotency for this version's own rules

    Dim fcG As Object
    Set fcG = ws.Range("I3:I2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$I3=""Eligible""")
    fcG.Interior.Color = RGB(226, 239, 218)

    Dim fcS As Object
    Set fcS = ws.Range("I3:I2000").FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$I3=""Not Yet Eligible""")
    fcS.Interior.Color = RGB(255, 230, 153)

    With ws.Parent.Windows(1)
        .FreezePanes = False
    End With
    ws.Cells(DATA_START, 2).Select
    ws.Parent.Windows(1).FreezePanes = True
    ws.Parent.Windows(1).DisplayGridlines = False
End Sub

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
                    "Last Increase", "Market Rent", "Next Increase", "Status", "Notes", "Import")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i

    With ws.Range("A2:K2")
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = BAR_GREY
    End With
End Sub

' ----------------------------------------------------------------
'  DoRefreshMTM  -  rebuilds the sheet as a single flat list: every
'  mtmDict entry is written in one pass starting at DATA_START, then
'  the whole written range is sorted by col H (Next Increase). Units
'  that dropped out of the current scan are simply not written back
'  (no flag, no retained row).
'
'  mtmDict maps unit number -> index into mtmRecs() (see
'  modReaders.ReadYardiMTM).
'
'  Cols F (Last Increase), J (Notes), and K (checkbox) are manual/
'  carried and are pulled forward verbatim from the pre-refresh
'  snapshot if the unit existed in it.
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object, mtmRecs() As MTMUnitRec)
    ' Unwrap any live Table before the full-row delete below - deleting
    ' rows out from under a live ListObject down to zero data rows is
    ' an unsupported operation. .Unlist preserves all cell values/formatting.
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then lastRow = DATA_START - 1

    ' Resolved before the destructive delete below - if this were ever to
    ' fail, we haven't wiped any data yet.
    ' targetEnd = last calendar date of the month containing today+90 days.
    Dim tDate As Date: tDate = Now + 90
    Dim targetEnd As Date: targetEnd = DateSerial(Year(tDate), Month(tDate) + 1, 1) - 1

    ' Snapshot pass: capture F/J/K for existing tracked units before
    ' clearing anything. Skip blank rows.
    Dim snapshot As Object: Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u = "" Then GoTo NextSnapRow
        snapshot(u) = Array(ws.Cells(r, 6).Value, ws.Cells(r, 10).Value, ws.Cells(r, 11).Value)
NextSnapRow:
    Next r

    ' Clear (delete, not ClearContents) all data rows so any leftover
    ' divider row from a prior version is cleanly disposed of.
    If lastRow >= DATA_START Then
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRow, COL_COUNT)).Delete Shift:=xlUp
    End If

    ' --- Single pass: write every mtmDict entry starting at DATA_START ---
    Dim key As Variant
    Dim writeRow As Long: writeRow = DATA_START
    Dim mtmCount As Long: mtmCount = 0
    For Each key In mtmDict.Keys
        WriteMTMDataRow ws, writeRow, CStr(key), mtmRecs(CLng(mtmDict(key))), snapshot, targetEnd
        writeRow = writeRow + 1
        mtmCount = mtmCount + 1
    Next key
    If mtmCount > 0 Then
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(DATA_START + mtmCount - 1, COL_COUNT)).Sort _
            Key1:=ws.Cells(DATA_START, 8), Order1:=xlAscending, Header:=xlNo

        ' Rebuild the Table over the fresh, final range (row 2 header through
        ' last data row) - always recreated fresh so refreshes never
        ' accumulate orphaned Table1/Table2 auto-named leftovers.
        Dim lo2 As ListObject
        Set lo2 = ws.ListObjects.Add(xlSrcRange, _
            ws.Range(ws.Cells(DATA_START - 1, 1), ws.Cells(DATA_START + mtmCount - 1, COL_COUNT)), , xlYes)
        lo2.Name = "MTMTable"
        lo2.TableStyle = "TableStyleLight9"
        lo2.ShowTableStyleRowStripes = True
        lo2.ShowAutoFilter = False

        ' Table styling overrides the header row's own formatting - reapply it
        ' (fill/bold/wrap from WriteHeaders, borders from FormatMTMSheet).
        With ws.Range(ws.Cells(2, 1), ws.Cells(2, COL_COUNT))
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = BAR_GREY
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With

        ' Explicit alternating row fill (Table quick-style stripes are not
        ' reliably visible in all Excel renderings) - guaranteed banding.
        ApplyRowBanding ws, DATA_START + mtmCount - 1
    End If

    ' Rebuild checkboxes over the new full range
    SyncCheckboxes ws
End Sub

' ----------------------------------------------------------------
'  ApplyRowBanding  -  applies straightforward alternating light
'  row fills (white / light grey) across A:COL_COUNT for the data
'  rows DATA_START..lastRow, as a guaranteed fallback for readability
'  (the MTMTable's own TableStyleLight9 row stripes are not always
'  visually apparent). Called after DoRefreshMTM writes/sorts the
'  rows, and again from SortMTMSheet after it re-sorts (since sorting
'  reshuffles which rows land on odd/even positions).
' ----------------------------------------------------------------
Private Sub ApplyRowBanding(ws As Worksheet, lastRow As Long)
    Dim r As Long
    For r = DATA_START To lastRow
        If (r - DATA_START) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Interior.Color = RGB(255, 255, 255)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Interior.Color = RGB(242, 242, 242)
        End If
    Next r
End Sub

' ----------------------------------------------------------------
'  MTMStatusLabel  -  informational-only Status label. Has zero effect
'  on checkbox availability or import eligibility.
' ----------------------------------------------------------------
Private Function MTMStatusLabel(niCell As Variant, targetEnd As Date) As String
    If IsDate(niCell) And CDate(niCell) > targetEnd Then
        MTMStatusLabel = "Not Yet Eligible"
    Else
        MTMStatusLabel = "Eligible"
    End If
End Function

' ----------------------------------------------------------------
'  WriteMTMDataRow  -  writes one MTMUnitRec to row r, carrying
'  forward F/J/K from the pre-refresh snapshot (see DoRefreshMTM)
'  unconditionally if the unit existed in it.
' ----------------------------------------------------------------
Private Sub WriteMTMDataRow(ws As Worksheet, r As Long, unit As String, rec As MTMUnitRec, _
                             snapshot As Object, targetEnd As Date)
    ws.Cells(r, 1).Value = unit
    ws.Cells(r, 2).Value = rec.Name
    ws.Cells(r, 3).Value = rec.FloorPlanCode
    If IsDate(rec.ExpiryVal) Then ws.Cells(r, 4).Value = CDate(rec.ExpiryVal)
    ws.Cells(r, 5).Value = rec.ActualRent
    If IsNumeric(rec.MarketRent) Then ws.Cells(r, 7).Value = CDbl(rec.MarketRent)

    Dim hasSnap As Boolean: hasSnap = snapshot.Exists(unit)
    Dim snapArr As Variant
    If hasSnap Then snapArr = snapshot(unit)

    Dim liVal As Variant
    If hasSnap Then
        liVal = snapArr(0)
        If Not IsEmpty(liVal) And CStr(liVal) <> "" Then ws.Cells(r, 6).Value = liVal
        If IsDate(liVal) Then ws.Cells(r, 8).Value = NextIncreaseDate(CDate(liVal))
        Dim noteVal As Variant: noteVal = snapArr(1)
        If Not IsEmpty(noteVal) And CStr(noteVal) <> "" Then ws.Cells(r, 10).Value = noteVal
    End If

    ' Auto-note for stale MTM units with no Last Increase on record yet -
    ' never overwrites an existing manual note.
    If Trim(CStr(ws.Cells(r, 6).Value)) = "" And rec.StaleOut And Trim(CStr(ws.Cells(r, 10).Value)) = "" Then
        ws.Cells(r, 10).Value = "Lease expired 15+ months ago - verify resident ledger for actual increase history"
    End If

    ws.Cells(r, 9).Value = MTMStatusLabel(ws.Cells(r, 8).Value, targetEnd)

    If hasSnap Then
        ws.Cells(r, 11).Value = snapArr(2)
    Else
        ws.Cells(r, 11).Value = False
    End If
End Sub

' ----------------------------------------------------------------
'  SyncCheckboxes  -  deletes all mtmChk_* Form Control checkboxes
'                     and recreates one for every row with a non-blank
'                     unit number in col A, linked to col K. Status has
'                     no bearing on checkbox creation. Called after
'                     every sort so checkboxes stay aligned.
' ----------------------------------------------------------------
Private Sub SyncCheckboxes(ws As Worksheet)
    ' Collect names first (can't delete while iterating the collection)
    Dim names() As String
    Dim cnt As Long: cnt = 0
    Dim cb As CheckBox
    For Each cb In ws.CheckBoxes
        If Left(cb.Name, Len(CB_PREFIX)) = CB_PREFIX Then
            ReDim Preserve names(cnt)
            names(cnt) = cb.Name
            cnt = cnt + 1
        End If
    Next cb
    Dim i As Long
    For i = 0 To cnt - 1
        ' Tolerate a checkbox that was already deleted or renamed out from
        ' under us between the collect pass above and this delete call.
        On Error Resume Next
        ws.CheckBoxes(names(i)).Delete
        On Error GoTo 0
    Next i

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START Then Exit Sub

    Dim r As Long
    For r = DATA_START To lastRow
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then GoTo NextCB
        Dim cell As Range: Set cell = ws.Cells(r, 11)
        Dim newCB As CheckBox
        Set newCB = ws.CheckBoxes.Add(cell.Left + 1, cell.Top + 1, _
                                      cell.Width - 2, cell.Height - 2)
        newCB.Caption = ""
        newCB.Name = CB_PREFIX & r
        newCB.LinkedCell = cell.Address(External:=False)
NextCB:
    Next r
End Sub

' ----------------------------------------------------------------
'  PlaceUnitInSection  -  finds the right floor plan section on the
'                         month sheet and writes the apt#, Name, Floor
'                         Plan, and Current Rent to a blank col-B row
'                         (inserts one if needed).
' ----------------------------------------------------------------
Private Sub PlaceUnitInSection(mws As Worksheet, unitNum As String, grp As String, _
                                residentName As String, fpCode As String, currentRent As Variant)
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
            If inTarget Then secLast = r - 1: Exit For
            If LCase(aVal) = LCase(grp) Then
                inTarget = True
                secFirst = r + 1
            End If
        End If
    Next r
    If inTarget And secLast = 0 Then secLast = lastUsed
    If secFirst = 0 Then Exit Sub

    ' Skip if unit already present
    For r = secFirst To secLast
        If Trim(CStr(mws.Cells(r, 2).Value)) = unitNum Then Exit Sub
    Next r

    ' Find blank col-B row or insert one
    Dim blankRow As Long: blankRow = 0
    For r = secFirst To secLast
        If Trim(CStr(mws.Cells(r, 2).Value)) = "" Then blankRow = r: Exit For
    Next r
    If blankRow = 0 Then
        blankRow = secLast + 1
        mws.Rows(blankRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    mws.Cells(blankRow, 2).Value = unitNum
    If Trim(CStr(mws.Cells(blankRow, 3).Value)) = "" Then _
        mws.Cells(blankRow, 3).Value = residentName
    If Trim(CStr(mws.Cells(blankRow, 4).Value)) = "" Then _
        mws.Cells(blankRow, 4).Value = fpCode
    If Trim(CStr(mws.Cells(blankRow, 5).Value)) = "" Then
        If Not IsEmpty(currentRent) And IsNumeric(currentRent) Then mws.Cells(blankRow, 5).Value = CDbl(currentRent)
    End If
    Dim mtmCell As Range: Set mtmCell = mws.Cells(blankRow, 20)
    If Trim(CStr(mtmCell.Value)) = "" Then mtmCell.Value = "MTM"
    mtmCell.Interior.Color = RGB(255, 255, 0)
    mtmCell.Font.Bold = True
End Sub

' ----------------------------------------------------------------
'  NextIncreaseDate  -  Last Increase + 1 year + 1 day.
' ----------------------------------------------------------------
Private Function NextIncreaseDate(lastIncrease As Date) As Date
    NextIncreaseDate = DateAdd("yyyy", 1, lastIncrease) + 1
End Function
