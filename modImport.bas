Attribute VB_Name = "modImport"
Option Explicit

Private Const DEBUG_IMPORT As Boolean = False

' ================================================================
'  modImport  -  button handler, orchestration, and sheet writing.
'
'  Reads: modConfig (PropConfig, LoadConfig, GetGroupForCode, GroupIndex)
'         modReaders (MTMUnitRec, PickFile, ReadYardi, ReadYardiMTM,
'                     ReadUnitStats, ReadRP, LookupRP,
'                     ReadUnitRentsGrid, LookupGrid, ReadMovein,
'                     LookupFP, AddUnmapped, UnmappedList,
'                     HasUnmapped, ResetUnmapped)
'         modSheetUtils (IsSectionBar, InsertRowCopyFromSource,
'                        ClearDataCells, SheetExists, MonthSheetName,
'                        MonthYearPrefix)
'
'  Version 2.6.0
' ================================================================

' ----------------------------------------------------------------
'  SHEET RESOLUTION
' ----------------------------------------------------------------
Public Function ResolveMonthSheet(m As Integer, yr As Integer) As String
    Dim abbrevNm As String: abbrevNm = MonthSheetName(m, yr)
    If SheetExists(abbrevNm) Then ResolveMonthSheet = abbrevNm: Exit Function

    Dim yearNm As String: yearNm = MonthName(m) & " " & yr
    If SheetExists(yearNm) Then ResolveMonthSheet = yearNm: Exit Function

    Dim fullNm As String: fullNm = MonthName(m)
    If SheetExists(fullNm) Then ResolveMonthSheet = fullNm: Exit Function

    Dim legacy As Variant
    legacy = Array("Jan", "Feb", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    If SheetExists(CStr(legacy(m - 1))) Then
        ResolveMonthSheet = CStr(legacy(m - 1))
    Else
        ResolveMonthSheet = ""
    End If
End Function

' ================================================================
'  BUTTON CLICK HANDLER
' ================================================================
Public Sub ImportMonthlyData()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub
    ResetUnmapped

    Dim mStr As String
    mStr = InputBox("Month number (1-12):" & vbCrLf & _
                    "1=Jan 2=Feb 3=Mar 4=Apr 5=May 6=Jun" & vbCrLf & _
                    "7=Jul 8=Aug 9=Sep 10=Oct 11=Nov 12=Dec", _
                    "Import Monthly Data", Month(Now))
    If mStr = "" Then Exit Sub
    If Not IsNumeric(mStr) Then MsgBox "Enter a number 1-12.", vbExclamation: Exit Sub
    Dim mNum As Integer: mNum = CInt(mStr)
    If mNum < 1 Or mNum > 12 Then MsgBox "Enter a number 1-12.", vbExclamation: Exit Sub

    Dim yStr As String
    yStr = InputBox("Year:", "Import Monthly Data", cfg.yr)
    If yStr = "" Then Exit Sub
    If Not IsNumeric(yStr) Then MsgBox "Enter a 4-digit year.", vbExclamation: Exit Sub
    Dim yr As Integer: yr = CInt(yStr)
    If yr < 2000 Or yr > 2100 Then MsgBox "Enter a year between 2000 and 2100.", vbExclamation, "Import Monthly Data": Exit Sub

    MsgBox "You will be asked to select up to 5 report files." & vbCrLf & _
           "Click Cancel on any file to skip that report." & vbCrLf & vbCrLf & _
           "  1. Yardi Rent Roll (.xlsx)           REQUIRED" & vbCrLf & _
           "  2. Yardi Unit Statistics (.xlsx)     fills K, W" & vbCrLf & _
           "  3. RP Renewal Offer Analysis (.csv)  fills F, T (fallback)" & vbCrLf & _
           "  4. Unit Rents Grid (.xlsx)            fills F, N, T, Best Term" & vbCrLf & _
           "  5. Move-in Box Score (.xls)           fills M", _
           vbInformation, "Import Monthly Data"

    MsgBox "Step 1 of 5 - Select your YARDI RENT ROLL (.xlsx)", vbInformation, "Import"
    Dim yardiPath As String
    yardiPath = PickFile("Select Yardi Rent Roll", "xlsx")
    If yardiPath = "" Then Exit Sub

    MsgBox "Step 2 of 5 - Select your YARDI UNIT STATISTICS (.xlsx)" & vbCrLf & _
           "Fills: Col K (Occupied Avg) and Col W (Inplace Lease Avg)" & vbCrLf & _
           "Click Cancel to skip.", vbInformation, "Import"
    Dim statsPath As String: statsPath = PickFile("Select Yardi Unit Statistics", "xlsx")

    MsgBox "Step 3 of 5 - Select your REALPAGE RENEWAL OFFER ANALYSIS (.csv)" & vbCrLf & _
           "Fills: Col F (Yieldstar Inc), T (Current Term) when grid is missing" & vbCrLf & _
           "Click Cancel to skip.", vbInformation, "Import"
    Dim rpPath As String: rpPath = PickFile("Select RealPage Renewal Offer Analysis", "csv")

    MsgBox "Step 4 of 5 - Select your UNIT RENTS GRID (.xlsx)" & vbCrLf & _
           "Fills: Col F, N (New Lease Rent), T (Current Term), Best Term" & vbCrLf & _
           "Click Cancel to skip.", vbInformation, "Import"
    Dim gridPath As String: gridPath = PickFile("Select Unit Rents Grid", "xlsx")

    MsgBox "Step 5 of 5 - Select your MOVE-IN BOX SCORE (.xls)" & vbCrLf & _
           "Fills: Col M (Recent Avg Eff. Rent - 3-month avg by floor plan)" & vbCrLf & _
           "Click Cancel to skip.", vbInformation, "Import"
    Dim moveinPath As String: moveinPath = PickFile("Select Move-in Box Score", "xls")

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim msg As String
    msg = DoImport(cfg, mNum, yr, yardiPath, statsPath, rpPath, gridPath, moveinPath)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox msg, vbInformation, "Done"
    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Import Error"
End Sub

' ================================================================
'  CORE ORCHESTRATION
' ================================================================
Private Function DoImport(cfg As PropConfig, mNum As Integer, yr As Integer, _
                           yardiPath As String, statsPath As String, _
                           rpPath As String, gridPath As String, _
                           moveinPath As String) As String
    Dim shName As String: shName = ResolveMonthSheet(mNum, yr)
    If shName = "" Then
        DoImport = "No sheet found for " & MonthName(mNum) & " " & yr & "." & vbCrLf & _
                   "Run Generate Month Sheets for that year first."
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(shName)

    Dim yardiWB As Workbook
    Set yardiWB = Workbooks.Open(yardiPath, ReadOnly:=True, UpdateLinks:=False)
    Dim mthUnits() As Variant, mthCnt As Long
    ReadYardi cfg, yardiWB, mNum, yr, mthUnits, mthCnt

    Dim rpUnits() As Variant, rpCnt As Long: rpCnt = 0
    If rpPath <> "" Then
        If Dir(rpPath) <> "" Then
            Dim rpWB As Workbook
            Set rpWB = Workbooks.Open(rpPath, ReadOnly:=True, UpdateLinks:=False)
            ReadRP rpWB, rpUnits, rpCnt
            rpWB.Close False
        End If
    End If

    Dim mtmDict As Object
    Dim mtmRecs() As MTMUnitRec
    Set mtmDict = ReadYardiMTM(cfg, yardiWB, mtmRecs)
    yardiWB.Close False

    Dim fpAvgs() As Long
    ReDim fpAvgs(cfg.GroupCount - 1)
    If statsPath <> "" Then
        If Dir(statsPath) <> "" Then
            Dim statsWB As Workbook
            Set statsWB = Workbooks.Open(statsPath, ReadOnly:=True, UpdateLinks:=False)
            ReadUnitStats cfg, statsWB, fpAvgs
            statsWB.Close False
        End If
    End If

    Dim gridUnits() As Variant, gridCnt As Long: gridCnt = 0
    If gridPath <> "" Then
        If Dir(gridPath) <> "" Then
            Dim gridWB As Workbook
            Set gridWB = Workbooks.Open(gridPath, ReadOnly:=True, UpdateLinks:=False)
            ReadUnitRentsGrid cfg, gridWB, gridUnits, gridCnt
            gridWB.Close False
        End If
    End If

    Dim fpL() As Long
    ReDim fpL(cfg.GroupCount - 1)
    If moveinPath <> "" Then
        If Dir(moveinPath) <> "" Then
            Dim miWB As Workbook
            Set miWB = Workbooks.Open(moveinPath, ReadOnly:=True, UpdateLinks:=False)
            ReadMovein cfg, miWB, fpL
            miWB.Close False
        End If
    End If

    FillSheet cfg, ws, mthUnits, mthCnt, fpAvgs, fpL, rpUnits, rpCnt, gridUnits, gridCnt
    FillMTMRows cfg, ws, mtmDict, mtmRecs, fpAvgs, fpL, rpUnits, rpCnt, gridUnits, gridCnt

    Dim skipped As String: skipped = ""
    If statsPath = "" Then skipped = skipped & "  Col K, W - Yardi Unit Statistics not provided" & vbCrLf
    If rpPath = "" Then skipped = skipped & "  Col F, T fallback - RP Renewal Offer Analysis not provided" & vbCrLf
    If gridPath = "" Then skipped = skipped & "  Col F, N, T - Unit Rents Grid not provided" & vbCrLf
    If moveinPath = "" Then skipped = skipped & "  Col M - Move-in Box Score not provided" & vbCrLf

    DoImport = MonthName(mNum) & " " & yr & " - " & mthCnt & " units imported to '" & shName & "'." & _
               vbCrLf & vbCrLf & _
               "Always fill in:" & vbCrLf & _
               "  Col G  " & cfg.ShortName & " Rec. Increase (your input)" & vbCrLf & _
               "  Col I  Pet Fees (verify each unit)"
    If skipped <> "" Then
        DoImport = DoImport & vbCrLf & vbCrLf & "Skipped (files not selected):" & vbCrLf & skipped
    End If
    If HasUnmapped() Then
        DoImport = DoImport & vbCrLf & vbCrLf & _
                   "UNMAPPED YARDI CODES (units skipped - add these to the" & vbCrLf & _
                   "Property Setup code map and re-import):" & vbCrLf & "  " & UnmappedList()
    End If
End Function

' ================================================================
'  FILL SHEET
'  Writes collected data into the target month sheet.
'
'  Row insertion delegates to InsertRowCopyFromSource (modSheetUtils),
'  the shared 4-step unmerge/insert/format/formula helper also used
'  by modDynamic.InsertBufferRow. It unmerges the destination after
'  PasteFormats and before PasteFormulas, preventing the "merged
'  cell" crash (v2.0.0 fix).
' ================================================================
Private Sub FillSheet(cfg As PropConfig, ws As Worksheet, _
                       mthU() As Variant, mthCnt As Long, _
                       fpAvgs() As Long, fpL() As Long, _
                       rpU() As Variant, rpCnt As Long, _
                       gridU() As Variant, gridCnt As Long)
    If mthCnt = 0 Then Exit Sub

    ' --- Build section map ---
    Dim maxSec As Long: maxSec = cfg.GroupCount + 10
    Dim secLabel() As String, secFirst() As Long, secLast() As Long
    ReDim secLabel(maxSec): ReDim secFirst(maxSec): ReDim secLast(maxSec)
    Dim secCount As Long: secCount = 0

    Dim lastUsed As Long: lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Dim r As Long
    Dim boundaryRow As Long: boundaryRow = 0

    For r = 3 To lastUsed
        Dim aVal As String: aVal = Trim(CStr(ws.Cells(r, 1).Value))
        Dim dVal As String: dVal = Trim(CStr(ws.Cells(r, 4).Value))

        If InStr(1, dVal, "Total", vbTextCompare) > 0 Then
            boundaryRow = r: Exit For
        End If
        If InStr(1, aVal, "Total", vbTextCompare) > 0 Then
            boundaryRow = r: Exit For
        End If

        If IsSectionBar(ws, r) And aVal <> "" Then
            If secCount > 0 Then secLast(secCount - 1) = r - 1
            secLabel(secCount) = aVal
            secFirst(secCount) = r + 1
            secCount = secCount + 1
            If secCount > maxSec Then Exit For
        End If
    Next r
    If secCount = 0 Then Exit Sub

    If DEBUG_IMPORT Then
        MsgBox "FillSheet entry" & vbCrLf & _
               "  mthCnt   = " & mthCnt & vbCrLf & _
               "  secCount = " & secCount, _
               vbInformation, "FillSheet Debug"
    End If

    If boundaryRow > 0 Then
        secLast(secCount - 1) = boundaryRow - 1
    Else
        secLast(secCount - 1) = lastUsed
    End If

    Dim writeTermToT As Boolean
    writeTermToT = (InStr(1, CStr(ws.Cells(2, 21).Value), "Term", vbTextCompare) > 0)

    ' --- Assign each unit to its section ---
    Dim secUnitIdx() As Long, secUnitCnt() As Long
    ReDim secUnitIdx(maxSec, mthCnt): ReDim secUnitCnt(maxSec)

    Dim mi As Long, si As Long
    For mi = 0 To mthCnt - 1
        Dim grp As String: grp = GetGroupForCode(cfg, CStr(mthU(mi, 1)))
        If grp = "" Then AddUnmapped CStr(mthU(mi, 1)): GoTo NextUnit
        For si = 0 To secCount - 1
            If LCase(Trim(secLabel(si))) = LCase(Trim(grp)) Then
                secUnitIdx(si, secUnitCnt(si)) = mi
                secUnitCnt(si) = secUnitCnt(si) + 1
                Exit For
            End If
        Next si
NextUnit:
    Next mi

    ' --- Write sections (bottom-to-top to keep row numbers stable during inserts) ---
    For si = secCount - 1 To 0 Step -1
        If secUnitCnt(si) = 0 Then GoTo NextSec

        Dim avail As Long
        Dim availRows() As Long
        ReDim availRows(mthCnt + cfg.BufferRows + 5)
        avail = CountAvail(ws, secFirst(si), secLast(si), availRows)

        If DEBUG_IMPORT Then
            Dim need_ As Long: need_ = secUnitCnt(si) - avail
            MsgBox "Section [" & secLabel(si) & "]  (si=" & si & ")" & vbCrLf & _
                   "  secFirst   = " & secFirst(si) & vbCrLf & _
                   "  secLast    = " & secLast(si) & vbCrLf & _
                   "  secUnitCnt = " & secUnitCnt(si) & vbCrLf & _
                   "  avail      = " & avail & vbCrLf & _
                   "  need       = " & need_, _
                   vbInformation, "FillSheet Debug"
        End If

        ' Insert rows if more units than available slots
        Dim need As Long: need = secUnitCnt(si) - avail
        If need > 0 Then
            Dim ni As Long
            For ni = 1 To need
                Dim insertAt As Long: insertAt = secLast(si)
                ' Shared 4-step insert: unmerge source, insert, copy formats,
                ' unmerge dest, copy formulas (source auto-adjusts to the row
                ' that shifted down to insertAt + 1).
                InsertRowCopyFromSource ws, insertAt, insertAt
                ClearDataCells ws, insertAt
                secLast(si) = secLast(si) + 1
            Next ni
            avail = CountAvail(ws, secFirst(si), secLast(si), availRows)
        End If

        ' --- Write unit data ---
        Dim sectionGrp As String
        sectionGrp = GetGroupForCode(cfg, CStr(mthU(secUnitIdx(si, 0), 1)))
        Dim occAvg As Long: occAvg = LookupFP(cfg, fpAvgs, sectionGrp)
        Dim lAvg As Long:   lAvg = LookupFP(cfg, fpL, sectionGrp)

        Dim uI As Long
        For uI = 0 To secUnitCnt(si) - 1
            If uI >= avail Then Exit For
            Dim fillRow As Long: fillRow = availRows(uI)
            Dim idx As Long:     idx = secUnitIdx(si, uI)
            Dim unitNum As String: unitNum = CStr(mthU(idx, 0))

            Dim ysInc As Double, rpCurTerm As Long
            Dim hasRP As Boolean: hasRP = False
            If rpCnt > 0 Then hasRP = LookupRP(rpU, rpCnt, unitNum, ysInc, rpCurTerm)

            Dim newLease As Double, bestOff As Double, curEff As Double
            Dim bestTerm As Long, gridCurTerm As Long
            Dim hasGrid As Boolean: hasGrid = False
            If gridCnt > 0 Then _
                hasGrid = LookupGrid(gridU, gridCnt, unitNum, newLease, bestOff, curEff, bestTerm, gridCurTerm)

            ws.Cells(fillRow, 2).Value = unitNum                    ' B: Apt#
            ws.Cells(fillRow, 3).Value = mthU(idx, 2)              ' C: Resident
            ws.Cells(fillRow, 4).Value = mthU(idx, 1)              ' D: FloorPlan

            Dim ar As Double: ar = 0
            If IsNumeric(mthU(idx, 4)) Then ar = CDbl(mthU(idx, 4))
            ws.Cells(fillRow, 5).Value = ar                        ' E: CurrentRent

            If hasGrid And bestOff > 0 And curEff > 0 Then        ' F: YieldstarInc
                ws.Cells(fillRow, 6).Value = CLng(bestOff - curEff)
            ElseIf hasRP And ysInc <> 0 Then
                ws.Cells(fillRow, 6).Value = CLng(ysInc)
            End If

            ws.Cells(fillRow, 9).Value = 0                         ' I: Pet Fees default

            Dim mk As Double: mk = 0
            If IsNumeric(mthU(idx, 3)) Then mk = CDbl(mthU(idx, 3))
            ws.Cells(fillRow, 11).Value = mk                       ' K: MarketRent

            If occAvg > 0 Then ws.Cells(fillRow, 12).Value = occAvg  ' L: OccAvg
            If lAvg > 0 Then ws.Cells(fillRow, 13).Value = lAvg      ' M: RecentAvg
            If hasGrid And newLease > 0 Then ws.Cells(fillRow, 14).Value = CLng(newLease)  ' N: NewLeaseRent

            If IsDate(mthU(idx, 5)) Then _
                ws.Cells(fillRow, 16).Value = CDate(mthU(idx, 5))  ' P: LeaseEnd

            If hasGrid And gridCurTerm > 0 Then                    ' T: CurrentTerm
                ws.Cells(fillRow, 20).Value = gridCurTerm
            ElseIf hasRP And rpCurTerm > 0 Then
                ws.Cells(fillRow, 20).Value = rpCurTerm
            End If

            If hasGrid And bestTerm > 0 Then                       ' U: Notes / BestTerm
                If writeTermToT Then
                    ws.Cells(fillRow, 21).Value = bestTerm
                Else
                    ws.Cells(fillRow, 21).Value = "RP recommends " & bestTerm & "-month lease term"
                End If
            End If

            If occAvg > 0 Then ws.Cells(fillRow, 24).Value = occAvg  ' X: InplaceAvg
        Next uI
NextSec:
    Next si
End Sub

' Count empty (blank Apt# col B) rows in a section; fills availRows array
Private Function CountAvail(ws As Worksheet, rFirst As Long, rLast As Long, _
                             availRows() As Long) As Long
    Dim n As Long: n = 0
    Dim r As Long
    For r = rFirst To rLast
        If Trim(CStr(ws.Cells(r, 2).Value)) = "" Then
            If n <= UBound(availRows) Then
                availRows(n) = r
                n = n + 1
            End If
        End If
    Next r
    CountAvail = n
End Function

' ----------------------------------------------------------------
'  FillMTMRows  -  fills MTM rows on the month sheet that were
'                  skipped by FillSheet (col B has apt#). No longer
'                  gates on col E being blank, since ImportSelectedMTM
'                  may have already pre-filled Name/Floor Plan/Current
'                  Rent from the MTM tracker. This month's fresh report
'                  data for Name (col C), Floor Plan (col D), and
'                  Current Rent (col E) takes precedence and overwrites
'                  whatever was previously written there.
'
'                  mtmDict maps unit number -> index into mtmRecs()
'                  (see modReaders.ReadYardiMTM).
' ----------------------------------------------------------------
Private Sub FillMTMRows(cfg As PropConfig, ws As Worksheet, _
                         mtmDict As Object, mtmRecs() As MTMUnitRec, _
                         fpAvgs() As Long, fpL() As Long, _
                         rpU() As Variant, rpCnt As Long, _
                         gridU() As Variant, gridCnt As Long)
    If mtmDict Is Nothing Then Exit Sub
    If mtmDict.Count = 0 Then Exit Sub

    Dim lastUsed As Long
    lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    Dim rec As MTMUnitRec
    Dim r As Long
    For r = 3 To lastUsed
        If IsSectionBar(ws, r) Then GoTo NextMTMRow

        Dim dVal As String: dVal = Trim(CStr(ws.Cells(r, 4).Value))
        If InStr(1, dVal, "Total", vbTextCompare) > 0 Then Exit For

        Dim bVal As String: bVal = Trim(CStr(ws.Cells(r, 2).Value))
        If bVal = "" Then GoTo NextMTMRow
        If Not mtmDict.Exists(bVal) Then GoTo NextMTMRow

        rec = mtmRecs(CLng(mtmDict(bVal)))

        Dim grp As String: grp = GetGroupForCode(cfg, rec.FloorPlanCode)
        Dim occAvg As Long: occAvg = LookupFP(cfg, fpAvgs, grp)
        Dim lAvg As Long:   lAvg = LookupFP(cfg, fpL, grp)

        Dim ysInc As Double, rpCurTerm As Long
        Dim hasRP As Boolean: hasRP = False
        If rpCnt > 0 Then hasRP = LookupRP(rpU, rpCnt, bVal, ysInc, rpCurTerm)

        Dim newLease As Double, bestOff As Double, curEff As Double
        Dim bestTerm As Long, gridCurTerm As Long
        Dim hasGrid As Boolean: hasGrid = False
        If gridCnt > 0 Then
            hasGrid = LookupGrid(gridU, gridCnt, bVal, newLease, bestOff, curEff, bestTerm, gridCurTerm)
        End If

        ' Write columns
        ws.Cells(r, 3).Value = rec.Name
        ws.Cells(r, 4).Value = rec.FloorPlanCode
        ws.Cells(r, 5).Value = rec.ActualRent
        If hasGrid And bestOff > 0 And curEff > 0 Then
            ws.Cells(r, 6).Value = CLng(bestOff - curEff)
        ElseIf hasRP And ysInc <> 0 Then
            ws.Cells(r, 6).Value = CLng(ysInc)
        End If
        If IsNumeric(rec.MarketRent) Then ws.Cells(r, 11).Value = CDbl(rec.MarketRent)
        If occAvg > 0 Then ws.Cells(r, 12).Value = occAvg
        If lAvg > 0 Then ws.Cells(r, 13).Value = lAvg
        If hasGrid And newLease > 0 Then ws.Cells(r, 14).Value = CLng(newLease)
        If IsDate(rec.ExpiryVal) Then ws.Cells(r, 16).Value = CDate(rec.ExpiryVal)
        Dim mtmCell As Range: Set mtmCell = ws.Cells(r, 20)
        If Trim(CStr(mtmCell.Value)) = "" Then mtmCell.Value = "MTM"
        mtmCell.Interior.Color = RGB(255, 255, 0)
        mtmCell.Font.Bold = True
        If occAvg > 0 Then ws.Cells(r, 24).Value = occAvg
NextMTMRow:
    Next r
End Sub
