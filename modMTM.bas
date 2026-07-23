Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler and import tools.
'
'  Reads: modConfig (PropConfig, LoadConfig, BAR_GREY, GetGroupForCode)
'         modReaders (MTMUnitRec, PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists, NameExists, IsSectionBar,
'                        IsSectionHeader, MonthSheetName,
'                        PendingWindowSheets)
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
'  availability or import behavior. Next Increase = Last Increase + 12
'  months exactly (see NextIncreaseDate), computed only when a unit has
'  no prior Next Increase on record - existing values are always carried
'  forward untouched on refresh/sort. If Last Increase is blank and the
'  unit's lease expiry
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
'
'  Pending (Manual) section - a second block, below the Confirmed
'  (Rent Roll) section handled above, sourced from column A (Renewal
'  Status) on the rolling 3-month window of month sheets anchored to
'  whichever month/year was most recently imported (see
'  GetMTMAnchorDate / SetMTMAnchor, and modImport.ImportMonthlyData
'  which calls SetMTMAnchor after every import). AddPendingUnit is
'  called live by modDynamic the instant col A is set to "MTM" on a
'  windowed sheet; BuildPendingSection rebuilds the whole section from
'  scratch on every RefreshMTMSheet, as the authoritative source of
'  truth. FindConfirmedLastRow replaces the old blind
'  Cells(Rows.Count,1).End(xlUp) pattern everywhere in this module so
'  Confirmed-section logic never reaches down into Pending's rows.
'  Pending also carries 4 extra cols (F-I: Expected Increase Date, New
'  MTM Rate, Market Rent, Notes) so a unit's info is ready to move over
'  once it becomes a real Confirmed row - Market Rent is auto-imported
'  from the Rent Roll each refresh (see BuildPendingSection), the other
'  3 are manually maintained by the CM and carried forward across
'  rebuilds (see SnapshotPendingSection). Col J is a Keep/Lock checkbox -
'  if checked, BuildPendingSection's resurrection pass re-adds that unit
'  even if it falls outside the 3-month scan window and even after it
'  graduates to Confirmed; only explicit resolution on a later sheet
'  removes it.
' ================================================================

Private Const MTM_SHEET     As String = "MTM"
Private Const MTM_SHEET_OLD As String = "MTM & STL"
Private Const DATA_START    As Long = 3
Private Const COL_COUNT     As Long = 11
Private Const CB_PREFIX     As String = "mtmChk_"

Private Const PENDING_SECTION_LABEL As String = "Pending (Manual)"
Private Const MTM_ANCHOR_NAME       As String = "MTM.AnchorDate"
Private Const MTM_ANCHOR_CELL       As String = "M1"   ' free cell - buttons live at M2+, title merge ends at K1
Private Const PENDING_COL_COUNT     As Long   = 10
Private Const CB_PEND_PREFIX        As String = "mtmPendChk_"

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
        If CreateFreshMTMSheet() Is Nothing Then Exit Sub
    End If

    Dim yardiPath As String
    yardiPath = PickFile("Select Yardi Rent Roll for MTM Refresh", "xlsx")
    If yardiPath = "" Then Exit Sub

    Dim yardiWB As Workbook
    On Error GoTo ErrHandler
    Set yardiWB = Workbooks.Open(yardiPath, ReadOnly:=True, UpdateLinks:=False)

    Dim mtmRecs() As MTMUnitRec
    Dim mtmDict As Object
    ' allMarketRent covers every occupied unit's Market Rent regardless of
    ' MTM status (mtmDict/mtmRecs only cover units already MTM in Yardi) -
    ' used to auto-populate Pending (Manual)'s Market Rent column, since
    ' the Rent Roll workbook is closed right below and unavailable by the
    ' time BuildPendingSection runs.
    Dim allMarketRent As Object: Set allMarketRent = CreateObject("Scripting.Dictionary")
    allMarketRent.CompareMode = 1
    ' allActualRent - same treatment as allMarketRent above, but for
    ' Current Rent (see BuildPendingSection).
    Dim allActualRent As Object: Set allActualRent = CreateObject("Scripting.Dictionary")
    allActualRent.CompareMode = 1
    Set mtmDict = ReadYardiMTM(cfg, yardiWB, mtmRecs, allMarketRent, allActualRent)
    yardiWB.Close False
    Set yardiWB = Nothing

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)

    ' Idempotent - creates the named range if missing, never overwrites
    ' an existing anchor value. Falls back to Now() if still unset.
    Dim anchorDate As Date: anchorDate = GetMTMAnchorDate()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    FormatMTMSheet ws, cfg
    WriteHeaders ws, cfg
    DoRefreshMTM ws, mtmDict, mtmRecs, anchorDate, allMarketRent, allActualRent
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
    Dim lastRow As Long: lastRow = FindConfirmedLastRow(ws)
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
    Dim lastRow As Long: lastRow = FindConfirmedLastRow(ws)
    If lastRow < DATA_START Then Exit Sub

    ws.Range(ws.Cells(DATA_START, 11), ws.Cells(lastRow, 11)).Value = False
End Sub

' ----------------------------------------------------------------
'  SortMTMSheet  -  manually re-sorts the MTM tab by col H (Next
'  Increase), on demand, without running a full RefreshMTMSheet.
'  Next Increase (H) is left exactly as-is - Sort only reorders rows
'  by the existing H values and refreshes Status (I) to match, it never
'  overwrites H itself. Also dedups the Pending (Manual) section by
'  col A (Unit), deleting any duplicate rows found there.
' ----------------------------------------------------------------
Public Sub SortMTMSheet()
    MigrateMTMSheetName

    If Not SheetExists(MTM_SHEET) Then
        MsgBox "MTM tracker sheet not found.", vbExclamation, "Sort MTM"
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim lastRow As Long: lastRow = FindConfirmedLastRow(ws)
    If lastRow < DATA_START Then
        MsgBox "No MTM units to sort.", vbInformation, "Sort MTM"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' Refresh Status (I) from whatever is already sitting in Next Increase
    ' (H) - H itself is never written here, so existing Next Increase data
    ' survives a Sort untouched.
    ' targetEnd = last calendar date of the month containing today+90 days.
    Dim tDate As Date: tDate = Now + 90
    Dim targetEnd As Date: targetEnd = DateSerial(Year(tDate), Month(tDate) + 1, 1) - 1
    Dim r As Long
    For r = DATA_START To lastRow
        ws.Cells(r, 9).Value = MTMStatusLabel(ws.Cells(r, 8).Value, targetEnd)
    Next r

    ' Row 2 is a live "MTMTable" ListObject header. Sorting via a raw
    ' Range.Sort on the data body is unreliable for a Table (Excel does not
    ' reliably keep the header pinned at row 2 and resets it to the Table
    ' Style's default look) - go through the ListObject's own Sort object
    ' instead, which is the supported way to sort part of a Table.
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("MTMTable")
    On Error GoTo ErrHandler
    If Not lo Is Nothing Then
        With lo.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(DATA_START, 8), SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    Else
        ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastRow, COL_COUNT)).Sort _
            Key1:=ws.Cells(DATA_START, 8), Order1:=xlAscending, Header:=xlNo
    End If

    ' Sorting the Table resets its header row to the Table Style's own
    ' fill/font (mirrors the reapply block in DoRefreshMTM after rebuilding
    ' MTMTable) - put the black-on-grey header look back.
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

    ' Re-sorting reshuffles which rows are odd/even - refresh the banding.
    ApplyRowBanding ws, lastRow

    ' Rebuild checkboxes at their new row positions
    SyncCheckboxes ws

    ' Remove any duplicate Units that have landed in the Pending (Manual)
    ' section (col A), keeping the first occurrence of each.
    DedupPendingSection ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Sort MTM error: " & Err.Description, vbCritical, "Sort MTM"
End Sub

' ----------------------------------------------------------------
'  DedupPendingSection  -  scans the Pending (Manual) block's data rows
'  for duplicate Units (col A) and deletes the duplicates, keeping the
'  first occurrence. A safety-net cleanup run from SortMTMSheet -
'  BuildPendingSection/AddPendingUnit already dedup on write, but this
'  catches any duplicates that end up in the section regardless.
' ----------------------------------------------------------------
Private Sub DedupPendingSection(ws As Worksheet)
    Dim divRow As Long, hdrRow As Long, dFirst As Long, dLast As Long
    If Not FindPendingRange(ws, divRow, hdrRow, dFirst, dLast) Then Exit Sub
    If dLast < dFirst Then Exit Sub

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Dim dupRows As Collection: Set dupRows = New Collection
    Dim r As Long
    For r = dFirst To dLast
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u <> "" Then
            If seen.Exists(u) Then
                dupRows.Add r
            Else
                seen(u) = True
            End If
        End If
    Next r

    ' Delete bottom-up (one row at a time, highest row first) rather than
    ' a single multi-area Union delete, so this doesn't depend on Excel's
    ' handling of a non-contiguous Range.Delete.
    Dim i As Long
    For i = dupRows.Count To 1 Step -1
        ws.Range(ws.Cells(dupRows(i), 1), ws.Cells(dupRows(i), COL_COUNT)).Delete Shift:=xlUp
    Next i

    SyncPendingCheckboxes ws
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

    ' Pending (Manual) usage reminder, fixed at cols M-Q starting a couple
    ' rows below the button grid (which lives within row 2's own height -
    ' see EnsureMTMButtons) - a static note, not tied to the Pending
    ' block's row position, so it never needs to move as Confirmed/Pending
    ' grow or shrink on refresh.
    With ws.Range(ws.Cells(4, "M"), ws.Cells(9, "Q"))
        .Merge
        .Value = "Reminder: units in the ""Pending (Manual)"" section below are not yet MTM on the Yardi Rent Roll. Once a unit shows up as its own row in the tracker above - usually within the next couple of months, after a Refresh MTM Tracker picks it up from the Rent Roll - it is automatically removed from Pending (Manual) on that same refresh, carrying over its Expected Increase Date/New MTM Rate/Notes. No manual cleanup needed. Check a row's Keep/Lock box to keep it in Pending even if it falls outside the 3-month scan window or graduates to Confirmed - only explicit resolution (e.g. marked Renewed) on a later sheet removes it."
        .Font.Name = "Calibri"
        .Font.Italic = True
        .Font.Size = 10
        .Font.Color = RGB(89, 89, 89)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

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
'  CreateFreshMTMSheet  -  the single "create a fresh, fully-formatted,
'  empty MTM sheet" path: adds the sheet object, names it, then applies
'  the same Confirmed-section scaffolding (FormatMTMSheet/WriteHeaders)
'  RefreshMTMSheet builds from scratch - just with Confirmed left at
'  0 rows (no Table yet) until a real Refresh populates it from the
'  rent roll. Called by RefreshMTMSheet, AddPendingUnit, and
'  SetMTMAnchor whenever any of them needs to create the MTM sheet for
'  the first time, so the three creation paths can never diverge.
'  Loads its own PropConfig (idempotent - LoadConfig just re-reads the
'  Property Setup sheet) since AddPendingUnit/SetMTMAnchor don't carry
'  one of their own. Returns Nothing if config can't be loaded (no
'  sheet is created in that case).
'
'  suppressActivate - FormatMTMSheet (and Sheets.Add itself) activates
'  the new MTM sheet, which steals focus from whatever sheet the user
'  was on. AddPendingUnit (live edit on a month sheet) and SetMTMAnchor
'  (end of a monthly import) pass True so the user's original sheet is
'  restored afterward; RefreshMTMSheet passes False (default) since
'  activating the MTM sheet for the user to look at is the point of
'  clicking that button.
' ----------------------------------------------------------------
Private Function CreateFreshMTMSheet(Optional ByVal suppressActivate As Boolean = False) As Worksheet
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Function

    Dim prevActive As Worksheet
    If suppressActivate Then
        On Error Resume Next
        Set prevActive = ThisWorkbook.ActiveSheet
        On Error GoTo 0
    End If

    Dim newWs As Worksheet
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = MTM_SHEET

    FormatMTMSheet newWs, cfg
    WriteHeaders newWs, cfg

    If suppressActivate And Not prevActive Is Nothing Then prevActive.Activate

    Set CreateFreshMTMSheet = newWs
End Function

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
'  Cols F (Last Increase), H (Next Increase), J (Notes), and K (checkbox)
'  are manual/carried and are pulled forward verbatim from the
'  pre-refresh snapshot if the unit existed in it (H is only computed
'  fresh from F when the snapshot has no existing H value).
'
'  anchorDate drives the Pending (Manual) section rebuilt at the end
'  of this sub (see BuildPendingSection) - any existing Pending block
'  is removed up front so Confirmed's row-count changes below don't
'  leave it stranded at the wrong offset.
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object, mtmRecs() As MTMUnitRec, anchorDate As Date, _
                          allMarketRent As Object, allActualRent As Object)
    ' Unwrap any live Table before the full-row delete below - deleting
    ' rows out from under a live ListObject down to zero data rows is
    ' an unsupported operation. .Unlist preserves all cell values/formatting.
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    ' Snapshot Pending's manually-maintained cols F/G/I (Expected Increase
    ' Date/New MTM Rate/Notes, plus H/Market Rent as a fallback) by Unit#
    ' before the block below is deleted - BuildPendingSection has no block
    ' left to read from by the time it runs at the end of this sub, so
    ' this has to happen first.
    Dim pendSnapshot As Object: Set pendSnapshot = SnapshotPendingSection(ws)

    ' Removed up front - not a workaround for FindConfirmedLastRow (which
    ' correctly excludes Pending's rows either way), but because the
    ' Confirmed delete-and-rewrite below changes row counts, and leaving
    ' the old Pending block in place would let it get partially
    ' overwritten/stranded before BuildPendingSection rebuilds it at the
    ' end of this sub.
    DeletePendingBlock ws

    Dim lastRow As Long
    lastRow = FindConfirmedLastRow(ws)
    If lastRow < DATA_START Then lastRow = DATA_START - 1

    ' Resolved before the destructive delete below - if this were ever to
    ' fail, we haven't wiped any data yet.
    ' targetEnd = last calendar date of the month containing today+90 days.
    Dim tDate As Date: tDate = Now + 90
    Dim targetEnd As Date: targetEnd = DateSerial(Year(tDate), Month(tDate) + 1, 1) - 1

    ' Snapshot pass: capture F/H/J/K for existing tracked units before
    ' clearing anything. Skip blank rows. H (Next Increase) is carried
    ' forward as-is rather than recomputed, so a manually-corrected value
    ' survives a refresh untouched - as a static value, though: .Value is
    ' what's captured and written back, so if H held a formula (e.g.
    ' =EDATE(...)) it becomes a plain date and will no longer recalculate
    ' if F is edited afterward.
    Dim snapshot As Object: Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = 1
    Dim r As Long
    For r = DATA_START To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u = "" Then GoTo NextSnapRow
        snapshot(u) = Array(ws.Cells(r, 6).Value, ws.Cells(r, 8).Value, ws.Cells(r, 10).Value, ws.Cells(r, 11).Value)
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
        WriteMTMDataRow ws, writeRow, CStr(key), mtmRecs(CLng(mtmDict(key))), snapshot, targetEnd, pendSnapshot
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

    ' Pending is rebuilt from scratch every refresh - the authoritative
    ' source of truth, so units no longer marked "MTM" on a windowed
    ' month sheet simply disappear from it.
    BuildPendingSection ws, anchorDate, pendSnapshot, allMarketRent, allActualRent, mtmDict
End Sub

' ----------------------------------------------------------------
'  ApplyRowBanding  -  applies straightforward alternating light
'  row fills (white / light grey) across A:COL_COUNT for the data
'  rows DATA_START..lastRow, as a guaranteed fallback for readability
'  (the MTMTable's own TableStyleLight9 row stripes are not always
'  visually apparent). Called after DoRefreshMTM writes/sorts the
'  rows, and again from SortMTMSheet after it re-sorts (since sorting
'  reshuffles which rows land on odd/even positions).
'
'  Also re-stamps col H (Next Increase) with a soft highlight fill +
'  bold font every time, since the row-banding fill above would
'  otherwise overwrite it - these dates matter enough operationally to
'  stand out from the rest of the row without being a jarring color.
' ----------------------------------------------------------------
Private Sub ApplyRowBanding(ws As Worksheet, lastRow As Long)
    Dim r As Long
    For r = DATA_START To lastRow
        If (r - DATA_START) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Interior.Color = RGB(255, 255, 255)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Interior.Color = RGB(242, 242, 242)
        End If
        With ws.Cells(r, 8)
            .Interior.Color = RGB(255, 242, 204)
            .Font.Bold = True
        End With
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
'  forward F/H/J/K from the pre-refresh snapshot (see DoRefreshMTM)
'  unconditionally if the unit existed in it. H (Next Increase) is only
'  computed from F (Last Increase) when the snapshot has no existing H
'  value - otherwise the prior value is kept exactly as-is.
'
'  If the unit did NOT already exist as a Confirmed row (brand new this
'  refresh - i.e. it's moving over from Pending), F/H instead fall back
'  to pendSnapshot's Expected Increase Date for that unit, if any, so
'  the CM's estimate carries into Last Increase rather than coming over
'  blank (this unit is also being dropped from Pending this same
'  refresh - see BuildPendingSection's mtmDict check).
' ----------------------------------------------------------------
Private Sub WriteMTMDataRow(ws As Worksheet, r As Long, unit As String, rec As MTMUnitRec, _
                             snapshot As Object, targetEnd As Date, pendSnapshot As Object)
    ws.Cells(r, 1).Value = unit
    ws.Cells(r, 2).Value = rec.Name
    ws.Cells(r, 3).Value = rec.FloorPlanCode
    If IsDate(rec.ExpiryVal) Then ws.Cells(r, 4).Value = CDate(rec.ExpiryVal)
    ws.Cells(r, 5).Value = rec.ActualRent
    If IsNumeric(rec.MarketRent) Then ws.Cells(r, 7).Value = CDbl(rec.MarketRent)

    Dim hasSnap As Boolean: hasSnap = snapshot.Exists(unit)
    Dim snapArr As Variant
    If hasSnap Then snapArr = snapshot(unit)

    Dim liVal As Variant, niVal As Variant
    If hasSnap Then
        liVal = snapArr(0)
        If Not IsEmpty(liVal) And CStr(liVal) <> "" Then ws.Cells(r, 6).Value = liVal
        niVal = snapArr(1)
        If IsDate(niVal) Then
            ' Carry the existing Next Increase forward untouched.
            ws.Cells(r, 8).Value = niVal
        ElseIf IsDate(liVal) Then
            ' No prior Next Increase on record for this unit - compute it.
            ws.Cells(r, 8).Value = NextIncreaseDate(CDate(liVal))
        End If
        Dim noteVal As Variant: noteVal = snapArr(2)
        If Not IsEmpty(noteVal) And CStr(noteVal) <> "" Then ws.Cells(r, 10).Value = noteVal
    ElseIf pendSnapshot.Exists(unit) Then
        Dim pendArr As Variant: pendArr = pendSnapshot(unit)

        ' Expected Increase Date IS the next-increase estimate itself, not
        ' a Last Increase to project 12 months forward from - write it
        ' straight to Next Increase (H). Last Increase (F) is left blank;
        ' Pending never tracked it, so there's nothing real to seed it with.
        Dim pendExpDate As Variant: pendExpDate = pendArr(0)
        If IsDate(pendExpDate) Then ws.Cells(r, 8).Value = CDate(pendExpDate)

        ' New MTM Rate (pendArr(1)) and Notes (pendArr(3)) have no
        ' equivalent Confirmed column - fold them into Notes (J) so
        ' they're not silently lost now that this unit's Pending row is
        ' being dropped this same refresh (see BuildPendingSection's
        ' mtmDict check). Never overwrites an existing manual note.
        Dim pendRate As Variant: pendRate = pendArr(1)
        Dim pendNotes As Variant: pendNotes = pendArr(3)
        Dim carriedNote As String: carriedNote = ""
        If Not IsEmpty(pendRate) And IsNumeric(pendRate) Then
            carriedNote = "New MTM Rate (from Pending): " & Format(CDbl(pendRate), "$#,##0")
        End If
        If Not IsEmpty(pendNotes) And Trim(CStr(pendNotes)) <> "" Then
            If carriedNote <> "" Then carriedNote = carriedNote & " - "
            carriedNote = carriedNote & Trim(CStr(pendNotes))
        End If
        If carriedNote <> "" And Trim(CStr(ws.Cells(r, 10).Value)) = "" Then
            ws.Cells(r, 10).Value = carriedNote
        End If
    End If

    ' Auto-note for stale MTM units with no Last Increase on record yet -
    ' never overwrites an existing manual note.
    If Trim(CStr(ws.Cells(r, 6).Value)) = "" And rec.StaleOut And Trim(CStr(ws.Cells(r, 10).Value)) = "" Then
        ws.Cells(r, 10).Value = "Lease expired 15+ months ago - verify resident ledger for actual increase history"
    End If

    ws.Cells(r, 9).Value = MTMStatusLabel(ws.Cells(r, 8).Value, targetEnd)

    If hasSnap Then
        ws.Cells(r, 11).Value = snapArr(3)
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

    Dim lastRow As Long: lastRow = FindConfirmedLastRow(ws)
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
'  NextIncreaseDate  -  Last Increase + exactly 12 months (matches the
'  =EDATE(LastIncrease, 12) convention used for manual entries - no
'  extra +1 day).
' ----------------------------------------------------------------
Private Function NextIncreaseDate(lastIncrease As Date) As Date
    NextIncreaseDate = DateAdd("m", 12, lastIncrease)
End Function

' ================================================================
'  PENDING (MANUAL) SECTION
'
'  A second block below Confirmed (Rent Roll), sourced from column A
'  (Renewal Status) on the rolling 3-month window of month sheets
'  anchored to GetMTMAnchorDate(). BuildPendingSection is the
'  authoritative rebuild (called at the end of DoRefreshMTM);
'  AddPendingUnit is the live single-row picker called by
'  modDynamic.HandlePendingStatusChange the instant col A is set to
'  "MTM" on a windowed sheet. Both dedup by trimmed Unit# (col A of
'  the Pending block).
' ================================================================

' ----------------------------------------------------------------
'  FindConfirmedLastRow  -  scans forward from DATA_START for the
'  Pending divider row (the first IsSectionHeader-true row) and
'  returns row - 2, since BuildPendingSection/AddPendingUnit always
'  place the divider 2 rows below Confirmed's true last data row (one
'  blank spacer row in between) - row - 1 would return that blank
'  spacer, not Confirmed's true last row. Clamped so it never returns
'  below DATA_START - 1 (empty-Confirmed-section case). Falls back to
'  the old blind Cells(Rows.Count,1).End(xlUp) behavior if no divider
'  is found (legacy sheets predating this feature, or a brand new
'  sheet with no Pending section built yet). Replaces every blind
'  lastRow lookup in this module so Confirmed-section logic (import/
'  clear/sort/refresh/checkboxes) never reaches down into Pending's
'  rows.
' ----------------------------------------------------------------
Private Function FindConfirmedLastRow(ws As Worksheet) As Long
    Dim trueEnd As Long: trueEnd = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = DATA_START To trueEnd
        If IsSectionHeader(ws, r) Then
            Dim result As Long: result = r - 2
            If result < DATA_START - 1 Then result = DATA_START - 1
            FindConfirmedLastRow = result
            Exit Function
        End If
    Next r
    FindConfirmedLastRow = trueEnd
End Function

' ----------------------------------------------------------------
'  FindPendingRange  -  locates an existing Pending divider/header/
'  data block below Confirmed's current end. Returns False if none
'  exists yet (dFirst/dLast are set so dLast < dFirst signals zero
'  data rows in an otherwise-existing empty block).
' ----------------------------------------------------------------
Private Function FindPendingRange(ws As Worksheet, ByRef divRow As Long, ByRef hdrRow As Long, _
                                    ByRef dFirst As Long, ByRef dLast As Long) As Boolean
    FindPendingRange = False
    divRow = 0: hdrRow = 0: dFirst = 0: dLast = 0

    Dim confirmedEnd As Long: confirmedEnd = FindConfirmedLastRow(ws)
    Dim trueEnd As Long: trueEnd = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If trueEnd <= confirmedEnd Then Exit Function

    Dim r As Long
    For r = confirmedEnd + 1 To trueEnd
        If IsSectionHeader(ws, r) Then
            If Trim(CStr(ws.Cells(r, 1).Value)) = PENDING_SECTION_LABEL Then divRow = r
            Exit For
        End If
    Next r
    If divRow = 0 Then Exit Function

    hdrRow = divRow + 1
    dFirst = divRow + 2
    dLast = dFirst - 1   ' default: zero data rows
    For r = dFirst To trueEnd
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then Exit For
        dLast = r
    Next r

    FindPendingRange = True
End Function

' ----------------------------------------------------------------
'  DeletePendingBlock  -  removes an existing Pending block (divider
'  row through last data row) bottom-up in one Delete call, mirroring
'  DoRefreshMTM's own Confirmed-section delete step. No-op if no
'  Pending block currently exists.
' ----------------------------------------------------------------
Private Sub DeletePendingBlock(ws As Worksheet)
    Dim divRow As Long, hdrRow As Long, dFirst As Long, dLast As Long
    If Not FindPendingRange(ws, divRow, hdrRow, dFirst, dLast) Then Exit Sub

    Dim blockEnd As Long: blockEnd = hdrRow
    If dLast >= dFirst Then blockEnd = dLast

    ws.Range(ws.Cells(divRow, 1), ws.Cells(blockEnd, COL_COUNT)).Delete Shift:=xlUp
End Sub

' ----------------------------------------------------------------
'  WritePendingSectionShell  -  writes the Pending divider row (grey/
'  bold, matching the month-sheet section-bar convention detected by
'  IsSectionHeader) and its column header row (Unit/Name/Floor Plan/
'  Current Rent/Source Month/Expected Increase Date/New MTM Rate/Market
'  Rent/Notes), starting at divRow. Shared by BuildPendingSection (full
'  rebuild) and AddPendingUnit (live, first-unit case) so the two
'  creation paths never drift apart.
' ----------------------------------------------------------------
Private Sub WritePendingSectionShell(ws As Worksheet, divRow As Long)
    With ws.Range(ws.Cells(divRow, 1), ws.Cells(divRow, COL_COUNT))
        .Merge
        .Value = PENDING_SECTION_LABEL
        .Font.Name = "Garamond"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = BAR_GREY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(divRow).RowHeight = 18.75

    Dim hdrRow As Long: hdrRow = divRow + 1
    Dim pHeaders As Variant
    pHeaders = Array("Unit", "Name", "Floor Plan", "Current Rent", "Source Month", _
                      "Expected Increase Date", "New MTM Rate", "Market Rent", "Notes", "Keep/Lock")
    Dim i As Long
    For i = 0 To UBound(pHeaders)
        ws.Cells(hdrRow, i + 1).Value = pHeaders(i)
    Next i
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, PENDING_COL_COUNT))
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
    ws.Rows(hdrRow).RowHeight = 32.25
End Sub

' ----------------------------------------------------------------
'  WritePendingDataRow  -  writes one Pending row (cols A-E: Unit,
'  Name, Floor Plan, Current Rent, Source Month, plus cols F-I: Expected
'  Increase Date, New MTM Rate, Market Rent, Notes). Expected Increase
'  Date/New MTM Rate/Notes are manually maintained by the CM while a
'  unit is still Pending, so the values are sitting right there to copy
'  across once the unit shows up as a real row in the Confirmed section.
'  Market Rent is auto-imported from the Rent Roll (see BuildPendingSection
'  / allMarketRent) rather than manually maintained. Sets font/alignment/
'  number format explicitly rather than relying on FormatMTMSheet's
'  blanket per-column formats, since Pending's columns reuse the same
'  physical columns as Confirmed's but with different meanings per row
'  (e.g. Pending's D is Current Rent, not Confirmed's Lease Expiry).
'
'  expectedIncreaseDate/newMtmRate/marketRent/notes are optional -
'  omitted (or blank) for a brand-new Pending row. BuildPendingSection
'  passes expectedIncreaseDate/newMtmRate/notes back in from its
'  pre-delete snapshot so those manually-entered values survive every
'  refresh instead of being wiped on rebuild; marketRent is instead
'  looked up fresh from the Rent Roll on every rebuild.
'
'  Source Month is written as a real Date (the 1st of the sourced
'  month/year, parsed off sourceSheetName via modSheetUtils
'  .ParseMonthSheet - the same parser modDynamic/modAdmin/modOverview
'  already use to go from a month-sheet name to (month, year)) rather
'  than the literal sheet-name string, so NumberFormat actually
'  applies. Falls back to writing the raw string if the name can't be
'  parsed (e.g. a non-standard sheet name), matching prior behavior.
' ----------------------------------------------------------------
Private Sub WritePendingDataRow(ws As Worksheet, r As Long, unitNum As String, residentName As String, _
                                  fpCode As String, currentRent As Variant, sourceSheetName As String, _
                                  Optional expectedIncreaseDate As Variant, Optional newMtmRate As Variant, _
                                  Optional marketRent As Variant, Optional notes As Variant, _
                                  Optional keepLock As Boolean = False)
    ws.Cells(r, 1).Value = unitNum
    ws.Cells(r, 2).Value = residentName
    ws.Cells(r, 3).Value = fpCode
    If Not IsEmpty(currentRent) And IsNumeric(currentRent) Then ws.Cells(r, 4).Value = CDbl(currentRent)

    Dim srcMonth As Long, srcYear As Long
    If modSheetUtils.ParseMonthSheet(sourceSheetName, srcMonth, srcYear) Then
        ws.Cells(r, 5).Value = DateSerial(srcYear, srcMonth, 1)
    Else
        ws.Cells(r, 5).Value = sourceSheetName
    End If

    ' An omitted Optional Variant is the Missing placeholder (subtype
    ' vbError), not Empty - IsEmpty(Missing) is False and VBA's "And" does
    ' not short-circuit, so CStr/CDate/CDbl on it would raise error 13.
    ' IsMissing must be checked first, in its own If, for every one of these.
    If Not IsMissing(expectedIncreaseDate) Then
        If Not IsEmpty(expectedIncreaseDate) And IsDate(expectedIncreaseDate) Then ws.Cells(r, 6).Value = CDate(expectedIncreaseDate)
    End If
    If Not IsMissing(newMtmRate) Then
        If Not IsEmpty(newMtmRate) And IsNumeric(newMtmRate) Then ws.Cells(r, 7).Value = CDbl(newMtmRate)
    End If
    If Not IsMissing(marketRent) Then
        If Not IsEmpty(marketRent) And IsNumeric(marketRent) Then ws.Cells(r, 8).Value = CDbl(marketRent)
    End If
    If Not IsMissing(notes) Then
        If Not IsEmpty(notes) And Trim(CStr(notes)) <> "" Then ws.Cells(r, 9).Value = CStr(notes)
    End If

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, PENDING_COL_COUNT))
        .Font.Name = "Calibri"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(r, 4).NumberFormat = "$#,##0"
    ws.Cells(r, 5).NumberFormat = "mmm-yy;@"
    ws.Cells(r, 6).NumberFormat = "mm/dd/yy;@"
    ws.Cells(r, 7).NumberFormat = "$#,##0"
    ws.Cells(r, 8).NumberFormat = "$#,##0"
    With ws.Cells(r, 9)
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With

    ' Expected Increase Date - same soft highlight as Confirmed's Next
    ' Increase (col H there), since it's the Pending-side counterpart and
    ' just as important to notice at a glance.
    With ws.Cells(r, 6)
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
    End With

    ' Keep/Lock checkbox cell - hide raw TRUE/FALSE, matching how
    ' Confirmed's col K Import checkbox cell is hidden.
    ws.Cells(r, 10).Value = keepLock
    ws.Cells(r, 10).NumberFormat = ";;;"
End Sub

' ----------------------------------------------------------------
'  SnapshotPendingSection  -  captures the Pending block's manually-
'  maintained cols F,G,I (Expected Increase Date/New MTM Rate/Notes) by
'  trimmed Unit# (col A) into a Dictionary, for BuildPendingSection to
'  carry forward across its own rebuild. Col H (Market Rent) is also
'  captured, purely as a fallback for BuildPendingSection to fall back
'  on if a unit isn't found in the freshly-imported Rent Roll data (see
'  BuildPendingSection/allMarketRent) - it's otherwise superseded by that
'  fresh lookup. Also captures col J (Keep/Lock, index 4) plus Name/Floor
'  Plan/Current Rent/Source Month (indices 5-8), so BuildPendingSection's
'  resurrection pass can re-add a locked unit that fell out of the
'  window scan using the same values it last had. Must be called BEFORE
'  the Pending block is deleted (see DoRefreshMTM) - returns an empty
'  Dictionary if no Pending block exists yet.
' ----------------------------------------------------------------
Private Function SnapshotPendingSection(ws As Worksheet) As Object
    Dim pendSnapshot As Object: Set pendSnapshot = CreateObject("Scripting.Dictionary")
    pendSnapshot.CompareMode = 1
    Dim sDivRow As Long, sHdrRow As Long, sFirst As Long, sLast As Long
    If FindPendingRange(ws, sDivRow, sHdrRow, sFirst, sLast) Then
        If sLast >= sFirst Then
            Dim sr As Long
            For sr = sFirst To sLast
                Dim sUnit As String: sUnit = Trim(CStr(ws.Cells(sr, 1).Value))
                If sUnit <> "" Then
                    pendSnapshot(sUnit) = Array(ws.Cells(sr, 6).Value, ws.Cells(sr, 7).Value, _
                                                 ws.Cells(sr, 8).Value, ws.Cells(sr, 9).Value, _
                                                 ws.Cells(sr, 10).Value, ws.Cells(sr, 2).Value, _
                                                 ws.Cells(sr, 3).Value, ws.Cells(sr, 4).Value, _
                                                 ws.Cells(sr, 5).Value)
                End If
            Next sr
        End If
    End If
    Set SnapshotPendingSection = pendSnapshot
End Function

' ----------------------------------------------------------------
'  SyncPendingCheckboxes  -  deletes all mtmPendChk_* Form Control
'  checkboxes and recreates one for every Pending data row with a
'  non-blank unit number in col A, linked to col J (Keep/Lock). Mirrors
'  SyncCheckboxes above but scoped to CB_PEND_PREFIX/the Pending block,
'  so the two sync passes never collide. Called after BuildPendingSection,
'  AddPendingUnit, and DedupPendingSection so checkboxes stay aligned
'  with their rows.
' ----------------------------------------------------------------
Private Sub SyncPendingCheckboxes(ws As Worksheet)
    ' Collect names first (can't delete while iterating the collection)
    Dim names() As String
    Dim cnt As Long: cnt = 0
    Dim cb As CheckBox
    For Each cb In ws.CheckBoxes
        If Left(cb.Name, Len(CB_PEND_PREFIX)) = CB_PEND_PREFIX Then
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

    Dim divRow As Long, hdrRow As Long, dFirst As Long, dLast As Long
    If Not FindPendingRange(ws, divRow, hdrRow, dFirst, dLast) Then Exit Sub
    If dLast < dFirst Then Exit Sub

    Dim r As Long
    For r = dFirst To dLast
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then GoTo NextPendCB
        Dim cell As Range: Set cell = ws.Cells(r, PENDING_COL_COUNT)
        Dim newCB As CheckBox
        Set newCB = ws.CheckBoxes.Add(cell.Left + 1, cell.Top + 1, _
                                      cell.Width - 2, cell.Height - 2)
        newCB.Caption = ""
        newCB.Name = CB_PEND_PREFIX & r
        newCB.LinkedCell = cell.Address(External:=False)
NextPendCB:
    Next r
End Sub

' ----------------------------------------------------------------
'  BuildPendingSection  -  authoritative rebuild, called at the end of
'  DoRefreshMTM. Deletes any existing Pending block, then rescans
'  modSheetUtils.PendingWindowSheets(anchorDate) oldest to newest, so a
'  unit's LATEST status among the sheets it appears on in the window
'  always wins - if it's "MTM" there it's (re)added/refreshed, if it's
'  anything else (renewed, blank, etc.) any earlier "MTM" entry for that
'  same unit from an older sheet in the window is dropped. This is what
'  keeps Pending from re-importing a unit indefinitely just because it
'  was MTM 1-2 months ago and happens to still fall inside the 3-month
'  window - only its most current status controls. A unit already present
'  in mtmDict (a real row in the Confirmed section this refresh, from the
'  Yardi Rent Roll import) is likewise excluded/dropped from Pending even
'  if a window month sheet still shows "MTM" for it - nothing ever clears
'  col A on the month sheet once a unit goes Confirmed, so without this a
'  unit the CM already deleted from Pending (per the reminder note once
'  it shows up above) would just reappear on the next refresh. Writes one
'  row per distinct unit directly under Confirmed's new end. Pending is
'  fully rebuilt from scratch every time.
'
'  Cols F/G/I (Expected Increase Date/New MTM Rate/Notes) are manually
'  maintained by the CM and have no source on the month sheets, so the
'  caller (DoRefreshMTM) must snapshot them via SnapshotPendingSection
'  BEFORE its own early DeletePendingBlock call runs - by the time this
'  sub's own DeletePendingBlock call below would run, the block is
'  already gone, so it cannot be the one to capture them. pendSnapshot
'  is looked up by Unit and passed back into WritePendingDataRow for any
'  unit that's still present after the rescan.
'
'  Col H (Market Rent) is instead looked up fresh from allMarketRent - a
'  Dictionary of every unit's Market Rent from the same Rent Roll import
'  that fed the Confirmed section this refresh (see RefreshMTMSheet /
'  modReaders.ReadYardiMTM) - keyed by Unit#, regardless of MTM status,
'  so a Pending unit (not yet MTM in Yardi) still resolves. Falls back to
'  the pre-rebuild snapshot's old Market Rent if the unit isn't found in
'  the Rent Roll this time (e.g. a name/unit mismatch), rather than going
'  blank.
'
'  Col D (Current Rent) gets the same treatment via allActualRent - a
'  fresh Rent Roll value wins; falls back to whichever value the fresh
'  scan / resurrection pass above already produced if the unit isn't
'  found in the Rent Roll this time, rather than going blank.
'
'  A unit whose only windowed sheet fell out of the 3-month scan (the
'  anchor only advances on a formal monthly import, so a unit freshly
'  marked "MTM" this month before that month's import runs can otherwise
'  be silently dropped here) is re-added by the resurrection pass below
'  if its Keep/Lock box (pendSnapshot index 4) was checked, even if it has
'  since graduated to Confirmed (mtmDict) - only explicit resolution (see
'  latestStatus) removes it.
' ----------------------------------------------------------------
Private Sub BuildPendingSection(ws As Worksheet, anchorDate As Date, pendSnapshot As Object, _
                                  allMarketRent As Object, allActualRent As Object, mtmDict As Object)
    DeletePendingBlock ws

    Dim divRow As Long: divRow = FindConfirmedLastRow(ws) + 2
    WritePendingSectionShell ws, divRow

    Dim pendDict As Object: Set pendDict = CreateObject("Scripting.Dictionary")
    pendDict.CompareMode = 1

    ' Tracks each unit's latest non-blank status across the window scan
    ' below (oldest to newest sheet order, so the last write for a given
    ' unit is always its most current status) - used by the resurrection
    ' pass after the scan loop to tell a still-pending unit apart from one
    ' that was explicitly resolved on a later sheet still inside the window.
    Dim latestStatus As Object: Set latestStatus = CreateObject("Scripting.Dictionary")
    latestStatus.CompareMode = 1

    Dim windowNames As Variant: windowNames = modSheetUtils.PendingWindowSheets(anchorDate)
    Dim wi As Long
    For wi = 0 To UBound(windowNames)
        Dim shNm As String: shNm = CStr(windowNames(wi))
        If SheetExists(shNm) Then
            Dim mws As Worksheet: Set mws = ThisWorkbook.Sheets(shNm)
            Dim lastUsed As Long: lastUsed = mws.UsedRange.Row + mws.UsedRange.Rows.Count - 1
            Dim r As Long
            For r = 3 To lastUsed
                If IsSectionHeader(mws, r) Then GoTo NextScanRow
                Dim dValChk As String: dValChk = Trim(CStr(mws.Cells(r, 4).Value))
                If InStr(1, dValChk, "Total", vbTextCompare) > 0 Then Exit For

                Dim uNum As String: uNum = Trim(CStr(mws.Cells(r, 2).Value))
                If uNum = "" Then GoTo NextScanRow

                ' Window sheets are scanned oldest to newest, so whichever
                ' status this unit shows on the most recent sheet it
                ' appears on always wins here. If that status is an
                ' EXPLICIT non-"MTM" value (Renewed/NTV/etc. - the
                ' resident renewed / situation resolved since an older
                ' sheet in the window flagged it), drop it - don't keep
                ' re-importing a unit just because it was MTM 1-2 months
                ' ago. A blank status is NOT treated as a resolution -
                ' PlaceUnitInSection can copy a unit onto a later window
                ' sheet without ever setting col A, and a blank there just
                ' means that sheet's review hasn't happened yet, not that
                ' the unit stopped being MTM.
                Dim statusVal As String: statusVal = Trim(CStr(mws.Cells(r, 1).Value))
                If statusVal <> "" Then latestStatus(uNum) = statusVal

                ' Keep/Lock override: a unit the CM has locked stays visible
                ' in Pending even after it graduates to a real Confirmed row,
                ' so the mtmDict exclusion just below only applies to units
                ' that aren't locked. Units with no prior snapshot entry are
                ' unlocked/new, so the exclusion still applies to them as
                ' before.
                Dim isUnitLocked As Boolean: isUnitLocked = False
                If pendSnapshot.Exists(uNum) Then
                    Dim snapLockArr As Variant: snapLockArr = pendSnapshot(uNum)
                    If Not IsEmpty(snapLockArr(4)) Then isUnitLocked = (snapLockArr(4) = True)
                End If

                If LCase(statusVal) = "mtm" Then
                    If mtmDict.Exists(uNum) And Not isUnitLocked Then
                        ' Already a real row in the Confirmed section this
                        ' refresh (picked up by the Yardi Rent Roll import
                        ' above) - don't also re-add/keep it in Pending.
                        ' Nothing clears col A on the month sheet once a
                        ' unit goes Confirmed, so without this check a unit
                        ' the CM already deleted from Pending (per the
                        ' reminder note) would just come right back on the
                        ' next refresh.
                        If pendDict.Exists(uNum) Then pendDict.Remove uNum
                    Else
                        pendDict(uNum) = Array(uNum, Trim(CStr(mws.Cells(r, 3).Value)), _
                                                Trim(CStr(mws.Cells(r, 4).Value)), mws.Cells(r, 5).Value, shNm)
                    End If
                ElseIf statusVal <> "" And pendDict.Exists(uNum) Then
                    pendDict.Remove uNum
                End If
NextScanRow:
            Next r
        End If
    Next wi

    ' Resurrection pass - a unit whose windowed sheet fell out of the
    ' 3-month scan above (the anchor only advances on a formal monthly
    ' import, so a unit freshly marked "MTM" this month before that
    ' month's import runs can otherwise get silently dropped here even
    ' though it's still legitimately pending) is re-added if the CM had
    ' checked its Keep/Lock box, unless it was explicitly resolved (a
    ' later, non-"MTM" status - see latestStatus above). Keep/Lock
    ' overrides Confirmed-graduation here too, so a locked unit stays in
    ' Pending even if it has since graduated to a real Confirmed row
    ' (mtmDict) - explicit resolution is the only thing that removes it.
    Dim snapKey As Variant
    For Each snapKey In pendSnapshot.Keys
        Dim snapUnit As String: snapUnit = CStr(snapKey)
        If pendDict.Exists(snapUnit) Then GoTo NextSnapKey

        Dim snapRowArr As Variant: snapRowArr = pendSnapshot(snapUnit)
        If IsEmpty(snapRowArr(4)) Then GoTo NextSnapKey
        If snapRowArr(4) <> True Then GoTo NextSnapKey

        If latestStatus.Exists(snapUnit) Then
            If LCase(latestStatus(snapUnit)) <> "mtm" Then GoTo NextSnapKey
        End If

        Dim srcMonthStr As String
        Dim snapSrcDate As Variant: snapSrcDate = snapRowArr(8)
        If IsDate(snapSrcDate) Then
            srcMonthStr = MonthSheetName(Month(CDate(snapSrcDate)), Year(CDate(snapSrcDate)))
        Else
            srcMonthStr = CStr(snapSrcDate)
        End If

        pendDict(snapUnit) = Array(snapUnit, CStr(snapRowArr(5)), CStr(snapRowArr(6)), snapRowArr(7), srcMonthStr)
NextSnapKey:
    Next snapKey

    Dim writeR As Long: writeR = divRow + 2
    Dim key As Variant
    For Each key In pendDict.Keys
        Dim rowArr As Variant: rowArr = pendDict(key)
        Dim uKey As String: uKey = CStr(rowArr(0))

        Dim hasSnapPend As Boolean: hasSnapPend = pendSnapshot.Exists(uKey)
        Dim psArr As Variant
        If hasSnapPend Then psArr = pendSnapshot(uKey)

        ' Market Rent: fresh Rent Roll value wins; fall back to the
        ' pre-rebuild snapshot's old value if this unit isn't in the Rent
        ' Roll this time around. mrToUse is Dim'd once for the whole loop
        ' (VBA hoists Dim to the top of the procedure) so it must be reset
        ' every iteration - otherwise a unit with neither source would
        ' silently inherit the previous unit's Market Rent.
        Dim mrToUse As Variant: mrToUse = Empty
        If allMarketRent.Exists(uKey) Then
            mrToUse = allMarketRent(uKey)
        ElseIf hasSnapPend Then
            mrToUse = psArr(2)
        End If

        ' Current Rent: fresh Rent Roll value wins; otherwise falls back to
        ' whatever already supplied Current Rent before this lookup existed
        ' (rowArr(3) - the fresh scan's value off the month sheet, or the
        ' pre-rebuild snapshot's carried-forward value for a resurrected
        ' unit). Reset every iteration for the same reason as mrToUse above.
        Dim crToUse As Variant: crToUse = rowArr(3)
        If allActualRent.Exists(uKey) Then
            crToUse = allActualRent(uKey)
        End If

        ' Keep/Lock: carried forward from the pre-rebuild snapshot if this
        ' unit had one, otherwise defaults to unchecked (also reset every
        ' iteration for the same reason as mrToUse above).
        Dim keepLockToUse As Boolean: keepLockToUse = False
        If hasSnapPend Then
            If Not IsEmpty(psArr(4)) Then keepLockToUse = (psArr(4) = True)
        End If

        If hasSnapPend Then
            WritePendingDataRow ws, writeR, uKey, CStr(rowArr(1)), CStr(rowArr(2)), crToUse, CStr(rowArr(4)), _
                                 psArr(0), psArr(1), mrToUse, psArr(3), keepLockToUse
        Else
            WritePendingDataRow ws, writeR, uKey, CStr(rowArr(1)), CStr(rowArr(2)), crToUse, CStr(rowArr(4)), _
                                 , , mrToUse, , keepLockToUse
        End If
        writeR = writeR + 1
    Next key

    SyncPendingCheckboxes ws
End Sub

' ----------------------------------------------------------------
'  EnsureMTMAnchorCell  -  ensures the MTM.AnchorDate named range
'  exists, pointing at an out-of-the-way cell (M1 - row 1 beyond the
'  A1:K1 title merge, and above the button grid which starts at M2).
'  Never overwrites an existing name's target - but if the MTM sheet
'  was deleted (and possibly recreated) since the name was created,
'  NameExists still returns True even though RefersTo now points at
'  #REF! (a stale/broken name), so validate the existing name still
'  resolves to a real range before trusting it, deleting and
'  recreating it if not.
' ----------------------------------------------------------------
Private Function EnsureMTMAnchorCell(ws As Worksheet) As Range
    If NameExists(MTM_ANCHOR_NAME) Then
        Dim testRng As Range
        On Error Resume Next
        Set testRng = ThisWorkbook.Names(MTM_ANCHOR_NAME).RefersToRange
        On Error GoTo 0
        If testRng Is Nothing Then ThisWorkbook.Names(MTM_ANCHOR_NAME).Delete
    End If
    If Not NameExists(MTM_ANCHOR_NAME) Then
        ThisWorkbook.Names.Add Name:=MTM_ANCHOR_NAME, _
            RefersTo:="='" & ws.Name & "'!" & ws.Range(MTM_ANCHOR_CELL).Address
    End If
    Set EnsureMTMAnchorCell = ThisWorkbook.Names(MTM_ANCHOR_NAME).RefersToRange
End Function

' ----------------------------------------------------------------
'  GetMTMAnchorDate  -  reads the MTM.AnchorDate named cell, falling
'  back to Now() if unset or if the MTM sheet doesn't exist yet.
' ----------------------------------------------------------------
Public Function GetMTMAnchorDate() As Date
    MigrateMTMSheetName
    If Not SheetExists(MTM_SHEET) Then
        GetMTMAnchorDate = Now
        Exit Function
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim cel As Range: Set cel = EnsureMTMAnchorCell(ws)
    If IsDate(cel.Value) Then
        GetMTMAnchorDate = CDate(cel.Value)
    Else
        GetMTMAnchorDate = Now
    End If
End Function

' ----------------------------------------------------------------
'  SetMTMAnchor  -  called by modImport.ImportMonthlyData right after
'  every successful import, so the Pending window always tracks
'  whichever month/year was most recently imported. Creates the MTM
'  sheet (full Confirmed-section scaffolding via CreateFreshMTMSheet -
'  same as RefreshMTMSheet) if it doesn't exist yet.
' ----------------------------------------------------------------
Public Sub SetMTMAnchor(ByVal mNum As Long, ByVal yr As Long)
    MigrateMTMSheetName
    If Not SheetExists(MTM_SHEET) Then
        If CreateFreshMTMSheet(True) Is Nothing Then Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)
    Dim cel As Range: Set cel = EnsureMTMAnchorCell(ws)
    cel.Value = DateSerial(yr, mNum, 1)
End Sub

' ----------------------------------------------------------------
'  AddPendingUnit  -  live single-row picker, called by
'  modDynamic.HandlePendingStatusChange the instant col A is set to
'  "MTM" on a windowed month sheet. Ensures the MTM sheet exists (full
'  Confirmed-section scaffolding via CreateFreshMTMSheet - same as
'  RefreshMTMSheet - if it doesn't exist yet), dedups by trimmed Unit#
'  against any existing Pending block, and either appends one row or
'  (if no Pending block exists yet at all) creates the divider+header
'  +this one row.
'
'  mtmRate is optional - the month sheet's own col S ("MTM Rate"), passed
'  straight through to WritePendingDataRow's New MTM Rate column (G) if
'  the caller has it. Omitted (still Missing all the way through) is
'  fine - see the IsMissing guards in WritePendingDataRow.
' ----------------------------------------------------------------
Public Sub AddPendingUnit(unitNum As String, residentName As String, fpCode As String, _
                           currentRent As Variant, sourceSheetName As String, _
                           Optional mtmRate As Variant)
    MigrateMTMSheetName
    If Not SheetExists(MTM_SHEET) Then
        If CreateFreshMTMSheet(True) Is Nothing Then Exit Sub
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(MTM_SHEET)

    Dim trimmedUnit As String: trimmedUnit = Trim(unitNum)
    If trimmedUnit = "" Then Exit Sub

    Dim divRow As Long, hdrRow As Long, dFirst As Long, dLast As Long
    If FindPendingRange(ws, divRow, hdrRow, dFirst, dLast) Then
        ' Dedup: skip if this unit is already present in Pending.
        Dim r As Long
        If dLast >= dFirst Then
            For r = dFirst To dLast
                If Trim(CStr(ws.Cells(r, 1).Value)) = trimmedUnit Then Exit Sub
            Next r
        End If

        Dim newRow As Long: newRow = dFirst
        If dLast >= dFirst Then newRow = dLast + 1
        WritePendingDataRow ws, newRow, trimmedUnit, residentName, fpCode, currentRent, sourceSheetName, _
                             , mtmRate
        SyncPendingCheckboxes ws
    Else
        ' No Pending block exists yet - create divider + header + this
        ' one row, directly under Confirmed's current end.
        Dim newDivRow As Long: newDivRow = FindConfirmedLastRow(ws) + 2
        WritePendingSectionShell ws, newDivRow
        WritePendingDataRow ws, newDivRow + 2, trimmedUnit, residentName, fpCode, currentRent, sourceSheetName, _
                             , mtmRate
        SyncPendingCheckboxes ws
    End If
End Sub
