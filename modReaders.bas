Attribute VB_Name = "modReaders"
Option Explicit

' ================================================================
'  modReaders  -  all external file reading (Yardi, RP, Grid,
'                 Move-in Box Score) and related helpers.
'                 Pure data extraction - no sheet writes here.
'
'  Version 2.6.0
' ================================================================

' ----------------------------------------------------------------
'  MTM UNIT RECORD  -  one MTM-occupied unit from the Yardi Rent
'  Roll. Filled by ReadYardiMTM; consumed by modImport.FillMTMRows
'  and modMTM.WriteMTMDataRow via named fields (replaces the old
'  positional Array(...) values). Public because a Private Type
'  cannot appear in cross-module procedure signatures.
' ----------------------------------------------------------------
Public Type MTMUnitRec
    Name          As String
    FloorPlanCode As String
    MarketRent    As Variant   ' raw cell value - consumers check IsNumeric
    ActualRent    As Double
    ExpiryVal     As Variant   ' Date (past-due expiry) or String (non-date)
    StaleOut      As Boolean   ' True if past-date expiry is 15+ months ago
End Type

Private mUnmapped As String

' ----------------------------------------------------------------
'  FILE PICKER
' ----------------------------------------------------------------
Public Function PickFile(ttl As String, ext As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = ttl
        .Filters.Clear
        .Filters.Add ext & " Files", "*." & ext
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

' ----------------------------------------------------------------
'  HEADER SEARCH HELPERS
' ----------------------------------------------------------------
Public Function FindHeaderCol(ws As Worksheet, hRow As Long, txt As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
        If InStr(1, CStr(ws.Cells(hRow, c).Value), txt, vbTextCompare) > 0 Then
            FindHeaderCol = c: Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Public Function FindHeaderColExact(ws As Worksheet, hRow As Long, txt As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
        If LCase(Trim(CStr(ws.Cells(hRow, c).Value))) = LCase(Trim(txt)) Then
            FindHeaderColExact = c: Exit Function
        End If
    Next c
    FindHeaderColExact = 0
End Function

Public Function CleanNum(s As String) As Double
    Dim cleaned As String
    cleaned = Replace(Replace(Replace(s, "$", ""), ",", ""), " ", "")
    If IsNumeric(cleaned) Then CleanNum = CDbl(cleaned)
End Function

' ----------------------------------------------------------------
'  UNMAPPED CODE TRACKING
' ----------------------------------------------------------------
Public Sub ResetUnmapped()
    mUnmapped = ""
End Sub

Public Sub AddUnmapped(code As String)
    Dim cd As String: cd = LCase(Trim(code))
    If cd = "" Then Exit Sub
    If InStr(1, mUnmapped, "[" & cd & "]", vbTextCompare) = 0 Then
        mUnmapped = mUnmapped & "[" & cd & "]"
    End If
End Sub

Public Function UnmappedList() As String
    UnmappedList = Replace(Replace(mUnmapped, "][", ", "), "[", "")
    UnmappedList = Replace(UnmappedList, "]", "")
End Function

Public Function HasUnmapped() As Boolean
    HasUnmapped = (mUnmapped <> "")
End Function

' ----------------------------------------------------------------
'  FLOOR PLAN AVERAGE LOOKUP
' ----------------------------------------------------------------
Public Function LookupFP(cfg As PropConfig, fpVals() As Long, grp As String) As Long
    Dim i As Long: i = GroupIndex(cfg, grp)
    If i >= 0 Then LookupFP = fpVals(i) Else LookupFP = 0
End Function

' ================================================================
'  READ YARDI RENT ROLL (.xlsx)
'  Fills mthU(n, 0..5): Unit, Type, Name, MarketRent, ActualRent, LeaseEnd
' ================================================================
Public Sub ReadYardi(cfg As PropConfig, wb As Workbook, mth As Integer, yr As Integer, _
                     mthU() As Variant, mthCnt As Long)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    mthCnt = 0

    Dim cU As Long: cU = cfg.RRUnit
    Dim cT As Long: cT = cfg.RRType
    Dim cN As Long: cN = cfg.RRName
    Dim cM As Long: cM = cfg.RRMarket
    Dim cA As Long: cA = cfg.RRActual
    Dim cE As Long: cE = cfg.RRExpiry

    Dim dataRow As Long: dataRow = 0
    Dim r As Long
    For r = 1 To 50
        If MatchesAnyPattern(cfg, CStr(ws.Cells(r, cU).Value)) Then
            dataRow = r: Exit For
        End If
    Next r
    If dataRow = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cU).End(xlUp).Row
    ReDim mthU(lastRow, 5)

    For r = dataRow To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, cU).Value))
        If Not MatchesAnyPattern(cfg, u) Then GoTo NextUnit
        Dim nm As String: nm = Trim(CStr(ws.Cells(r, cN).Value))
        If nm = "" Or LCase(nm) = "vacant" Then GoTo NextUnit
        If Not IsNumeric(ws.Cells(r, cA).Value) Then GoTo NextUnit
        If CDbl(ws.Cells(r, cA).Value) <= 0 Then GoTo NextUnit

        If IsDate(ws.Cells(r, cE).Value) Then
            Dim ed As Date: ed = CDate(ws.Cells(r, cE).Value)
            If Month(ed) = mth And Year(ed) = yr Then
                mthU(mthCnt, 0) = u
                mthU(mthCnt, 1) = Trim(CStr(ws.Cells(r, cT).Value))
                mthU(mthCnt, 2) = nm
                mthU(mthCnt, 3) = ws.Cells(r, cM).Value
                mthU(mthCnt, 4) = ws.Cells(r, cA).Value
                mthU(mthCnt, 5) = ws.Cells(r, cE).Value
                mthCnt = mthCnt + 1
            End If
        End If
NextUnit:
    Next r
End Sub

' ================================================================
'  READ YARDI UNIT STATISTICS (.xlsx)  ->  col K and col W
'  Fills fpAvgs(groupIndex) with weighted occupied avg rent
' ================================================================
Public Sub ReadUnitStats(cfg As PropConfig, wb As Workbook, fpAvgs() As Long)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)

    Dim fpTotals() As Double, fpCounts() As Long
    ReDim fpTotals(cfg.GroupCount - 1)
    ReDim fpCounts(cfg.GroupCount - 1)

    Dim hRow As Long: hRow = 0
    Dim r As Long
    For r = 1 To 50
        If InStr(1, CStr(ws.Cells(r, 1).Value), "Unit Type", vbTextCompare) > 0 Then
            hRow = r: Exit For
        End If
    Next r
    If hRow = 0 Then Exit Sub

    Dim dataRow As Long: dataRow = 0
    For r = hRow + 1 To hRow + 6
        If InStr(CStr(ws.Cells(r, 1).Value), "(") > 0 Then dataRow = r: Exit For
    Next r
    If dataRow = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = dataRow To lastRow
        Dim aVal As String: aVal = Trim(CStr(ws.Cells(r, 1).Value))
        If aVal = "" Then GoTo NextStat
        If InStr(1, aVal, "Less ", vbTextCompare) > 0 Then Exit For
        If InStr(1, aVal, "Plus ", vbTextCompare) > 0 Then Exit For
        If InStr(1, aVal, "Net ", vbTextCompare) > 0 Then Exit For
        If InStr(1, aVal, "Eviction", vbTextCompare) > 0 Then Exit For
        If InStr(1, aVal, "Total", vbTextCompare) > 0 And InStr(aVal, "(") = 0 Then Exit For

        Dim p1 As Long: p1 = InStr(aVal, "(")
        Dim p2 As Long: p2 = InStr(aVal, ")")
        If p1 = 0 Or p2 = 0 Or p2 <= p1 Then GoTo NextStat
        Dim code As String: code = Trim(Mid(aVal, p1 + 1, p2 - p1 - 1))

        Dim grp As String: grp = GetGroupForCode(cfg, code)
        If grp = "" Then AddUnmapped code: GoTo NextStat

        Dim occCnt As Long
        If Not IsNumeric(ws.Cells(r, 3).Value) Then GoTo NextStat
        occCnt = CLng(ws.Cells(r, 3).Value)
        If occCnt <= 0 Then GoTo NextStat

        Dim avgRent As Double
        avgRent = CleanNum(CStr(ws.Cells(r, 8).Value))
        If avgRent <= 0 Then GoTo NextStat

        Dim gi As Long: gi = GroupIndex(cfg, grp)
        If gi >= 0 Then
            fpTotals(gi) = fpTotals(gi) + (avgRent * occCnt)
            fpCounts(gi) = fpCounts(gi) + occCnt
        End If
NextStat:
    Next r

    Dim i As Long
    For i = 0 To cfg.GroupCount - 1
        If fpCounts(i) > 0 Then fpAvgs(i) = CLng(fpTotals(i) / fpCounts(i))
    Next i
End Sub

' ================================================================
'  READ REALPAGE RENEWAL OFFER ANALYSIS (.csv)  ->  col F, T
'  Fills rpU(n, 0..3): Unit, IncDollar, CurTerm, OfferTerm
' ================================================================
Public Sub ReadRP(wb As Workbook, rpU() As Variant, rpCnt As Long)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    rpCnt = 0
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim rpU(lastRow, 3)

    Dim cUnit As Long:   cUnit = FindHeaderColExact(ws, 1, "Unit")
    Dim cTerm As Long:   cTerm = FindHeaderColExact(ws, 1, "Term")
    Dim cCurTerm As Long: cCurTerm = FindHeaderColExact(ws, 1, "Current Lease | Term")
    Dim cOff As Long:    cOff = FindHeaderColExact(ws, 1, "Renewal Offers | Rent")
    Dim cEff As Long:    cEff = FindHeaderColExact(ws, 1, "Current Lease | Effective Rent")
    If cUnit = 0 Then cUnit = FindHeaderCol(ws, 1, "Unit")
    If cCurTerm = 0 Then cCurTerm = FindHeaderCol(ws, 1, "Current Lease | Term")
    If cOff = 0 Then cOff = FindHeaderCol(ws, 1, "Renewal Offers | Rent")
    If cEff = 0 Then cEff = FindHeaderCol(ws, 1, "Current Lease | Effective Rent")
    If cUnit = 0 Or cTerm = 0 Then Exit Sub

    Dim bUnit() As String, bDiff() As Double
    ReDim bUnit(lastRow): ReDim bDiff(lastRow)
    Dim bPos As Long: bPos = 0

    Dim r As Long
    For r = 2 To lastRow
        Dim uv As String: uv = Trim(CStr(ws.Cells(r, cUnit).Value))
        If uv = "" Or uv = "Unit" Then GoTo NextRP

        Dim tv As String: tv = Trim(CStr(ws.Cells(r, cTerm).Value))
        If Not IsNumeric(tv) Then GoTo NextRP
        Dim tNum As Double: tNum = CDbl(tv)
        Dim diff As Double: diff = Abs(tNum - 12)

        Dim ofr As Double: ofr = 0
        Dim eff As Double: eff = 0
        Dim curTerm As Long: curTerm = 0
        If cOff > 0 Then If IsNumeric(ws.Cells(r, cOff).Value) Then ofr = CDbl(ws.Cells(r, cOff).Value)
        If cEff > 0 Then If IsNumeric(ws.Cells(r, cEff).Value) Then eff = CDbl(ws.Cells(r, cEff).Value)
        If cCurTerm > 0 Then If IsNumeric(ws.Cells(r, cCurTerm).Value) Then curTerm = CLng(ws.Cells(r, cCurTerm).Value)

        Dim inc As Double: inc = 0
        If eff > 0 And ofr > 0 Then inc = ofr - eff

        Dim found As Boolean: found = False
        Dim bi As Long
        For bi = 0 To bPos - 1
            If bUnit(bi) = uv Then
                If diff < bDiff(bi) Then
                    bDiff(bi) = diff
                    rpU(bi, 0) = uv: rpU(bi, 1) = inc
                    rpU(bi, 2) = curTerm: rpU(bi, 3) = CLng(tNum)
                End If
                found = True: Exit For
            End If
        Next bi

        If Not found Then
            rpU(bPos, 0) = uv: rpU(bPos, 1) = inc
            rpU(bPos, 2) = curTerm: rpU(bPos, 3) = CLng(tNum)
            bUnit(bPos) = uv: bDiff(bPos) = diff
            bPos = bPos + 1: rpCnt = rpCnt + 1
        End If
NextRP:
    Next r
End Sub

Public Function LookupRP(rpU() As Variant, rpCnt As Long, unitNum As String, _
                          outInc As Double, outCurTerm As Long) As Boolean
    outInc = 0: outCurTerm = 0
    Dim i As Long
    For i = 0 To rpCnt - 1
        If CStr(rpU(i, 0)) = unitNum Then
            outInc = CDbl(rpU(i, 1))
            outCurTerm = CLng(rpU(i, 2))
            LookupRP = True: Exit Function
        End If
    Next i
    LookupRP = False
End Function

' ================================================================
'  READ UNIT RENTS GRID (.xlsx)  ->  col F, N, T, BestTerm
'  Fills gridU(n, 0..5): Unit, NLRent, BestOff, CurEff, BestTm, CurTm
' ================================================================
Public Sub ReadUnitRentsGrid(cfg As PropConfig, wb As Workbook, _
                              gridU() As Variant, gridCnt As Long)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    gridCnt = 0

    Dim hRow As Long: hRow = 0
    Dim r As Long
    For r = 1 To 5
        If FindHeaderColExact(ws, r, "UNIT") > 0 Then hRow = r: Exit For
    Next r
    If hRow = 0 Then Exit Sub

    Dim cUnit As Long:   cUnit = FindHeaderColExact(ws, hRow, "UNIT")
    Dim cCurEff As Long: cCurEff = FindHeaderCol(ws, hRow, "CURRENT EFFECTIVE RENT")
    Dim cBest As Long:   cBest = FindHeaderColExact(ws, hRow, "BEST OFFER")
    Dim cBestTm As Long: cBestTm = FindHeaderCol(ws, hRow, "BEST OFFER TERM")
    Dim cNL As Long:     cNL = FindHeaderColExact(ws, hRow, "NEW LEASE RENT")
    Dim cCurTm As Long:  cCurTm = FindHeaderColExact(ws, hRow, "CURRENT TERM")
    If cUnit = 0 Then cUnit = cfg.GridUnit
    If cCurEff = 0 Then cCurEff = cfg.GridCurEff
    If cBest = 0 Then cBest = cfg.GridBest
    If cBestTm = 0 Then cBestTm = cfg.GridBestTerm
    If cNL = 0 Then cNL = cfg.GridNewLease

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cUnit).End(xlUp).Row
    ReDim gridU(lastRow, 5)

    For r = hRow + 1 To lastRow
        Dim uv As String: uv = Trim(CStr(ws.Cells(r, cUnit).Value))
        If uv = "" Or InStr(1, uv, "UNIT", vbTextCompare) > 0 Then GoTo NextGrid

        Dim nlRent As Double:  nlRent = CleanNum(CStr(ws.Cells(r, cNL).Value))
        Dim bestOff As Double: bestOff = CleanNum(CStr(ws.Cells(r, cBest).Value))
        Dim curEff As Double:  curEff = CleanNum(CStr(ws.Cells(r, cCurEff).Value))
        Dim bestTm As Long:  bestTm = 0
        Dim curTm As Long:   curTm = 0
        If IsNumeric(ws.Cells(r, cBestTm).Value) Then bestTm = CLng(ws.Cells(r, cBestTm).Value)
        If cCurTm > 0 Then
            If IsNumeric(ws.Cells(r, cCurTm).Value) Then curTm = CLng(ws.Cells(r, cCurTm).Value)
        End If

        gridU(gridCnt, 0) = uv:     gridU(gridCnt, 1) = nlRent
        gridU(gridCnt, 2) = bestOff: gridU(gridCnt, 3) = curEff
        gridU(gridCnt, 4) = bestTm:  gridU(gridCnt, 5) = curTm
        gridCnt = gridCnt + 1
NextGrid:
    Next r
End Sub

Public Function LookupGrid(gridU() As Variant, gridCnt As Long, unitNum As String, _
                            outNL As Double, outBestOff As Double, outCurEff As Double, _
                            outBestTerm As Long, outCurTerm As Long) As Boolean
    outNL = 0: outBestOff = 0: outCurEff = 0: outBestTerm = 0: outCurTerm = 0
    Dim i As Long
    For i = 0 To gridCnt - 1
        If CStr(gridU(i, 0)) = unitNum Then
            outNL = CDbl(gridU(i, 1))
            outBestOff = CDbl(gridU(i, 2))
            outCurEff = CDbl(gridU(i, 3))
            outBestTerm = CLng(gridU(i, 4))
            outCurTerm = CLng(gridU(i, 5))
            LookupGrid = True: Exit Function
        End If
    Next i
    LookupGrid = False
End Function

' ================================================================
'  READ MOVE-IN BOX SCORE (.xls)  ->  col M
'  Fills fpL(groupIndex) with 3-month avg effective rent by floor plan
' ================================================================
Public Sub ReadMovein(cfg As PropConfig, wb As Workbook, fpL() As Long)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)

    Dim fpTotals() As Double, fpCounts() As Long
    ReDim fpTotals(cfg.GroupCount - 1)
    ReDim fpCounts(cfg.GroupCount - 1)

    Dim cUnit As Long: cUnit = cfg.MIUnit

    Dim hRow As Long: hRow = 0
    Dim r As Long
    For r = 1 To 10
        If FindHeaderCol(ws, r, "Unit Type") > 0 Then hRow = r: Exit For
    Next r
    If hRow = 0 Then Exit Sub

    Dim cUT As Long:  cUT = FindHeaderCol(ws, hRow, "Unit Type")
    Dim cRent As Long: cRent = FindHeaderColExact(ws, hRow, "Rent")
    Dim cMI As Long:  cMI = FindHeaderCol(ws, hRow, "Move In")
    If cUT = 0 Then cUT = cfg.MIUnitType
    If cRent = 0 Then cRent = cfg.MIRent
    If cMI = 0 Then cMI = cfg.MIMoveIn

    Dim cutoff As Date: cutoff = DateAdd("m", -3, Date)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cUnit).End(xlUp).Row

    For r = hRow + 1 To lastRow
        Dim unitVal As String: unitVal = Trim(CStr(ws.Cells(r, cUnit).Value))
        If unitVal = "" Then GoTo NextMI
        If Not MatchesAnyPattern(cfg, unitVal) Then GoTo NextMI
        If Not IsDate(ws.Cells(r, cMI).Value) Then GoTo NextMI
        Dim miDate As Date: miDate = CDate(ws.Cells(r, cMI).Value)
        If miDate < cutoff Then GoTo NextMI

        Dim utCode As String: utCode = Trim(CStr(ws.Cells(r, cUT).Value))
        Dim grp As String: grp = GetGroupForCode(cfg, utCode)
        If grp = "" Then AddUnmapped utCode: GoTo NextMI

        Dim rent As Double
        If Not IsNumeric(ws.Cells(r, cRent).Value) Then GoTo NextMI
        rent = CDbl(ws.Cells(r, cRent).Value)
        If rent <= 0 Then GoTo NextMI

        Dim gi As Long: gi = GroupIndex(cfg, grp)
        If gi >= 0 Then
            fpTotals(gi) = fpTotals(gi) + rent
            fpCounts(gi) = fpCounts(gi) + 1
        End If
NextMI:
    Next r

    Dim i As Long
    For i = 0 To cfg.GroupCount - 1
        If fpCounts(i) > 0 Then fpL(i) = CLng(fpTotals(i) / fpCounts(i))
    Next i
End Sub

' ----------------------------------------------------------------
'  ReadYardiMTM  -  returns all MTM-occupied units (past-date or
'                   non-date lease expiry) from the Yardi Rent Roll.
'                   A future-dated expiry simply means the unit is
'                   not on the tracker at all.
'
'                   Returns a Dictionary keyed by unit number whose
'                   item is a Long index into the recs() array of
'                   MTMUnitRec (VBA cannot store user-defined types
'                   directly as Dictionary items). Consumers keep
'                   using dict.Exists / dict.Count / dict.Keys and
'                   read named fields off recs(dict(unit)).
'
'                   StaleOut = True if the past-date expiry is more
'                   than 15 months ago (per DateDiff("m", ...)),
'                   else False. Non-date expiries are always
'                   StaleOut = False.
' ----------------------------------------------------------------
Public Function ReadYardiMTM(cfg As PropConfig, wb As Workbook, _
                             recs() As MTMUnitRec) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' vbTextCompare
    ReDim recs(0)

    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cfg.RRUnit).End(xlUp).Row

    Dim cU As Long: cU = cfg.RRUnit
    Dim cT As Long: cT = cfg.RRType
    Dim cN As Long: cN = cfg.RRName
    Dim cM As Long: cM = cfg.RRMarket
    Dim cA As Long: cA = cfg.RRActual
    Dim cE As Long: cE = cfg.RRExpiry

    ' Find first data row (skip headers)
    Dim dataStart As Long: dataStart = 0
    Dim r As Long
    For r = 1 To 50
        Dim uVal As String: uVal = Trim(CStr(ws.Cells(r, cU).Value))
        If uVal <> "" And MatchesAnyPattern(cfg, uVal) Then
            dataStart = r: Exit For
        End If
    Next r
    If dataStart = 0 Then Set ReadYardiMTM = dict: Exit Function

    ReDim recs(lastRow)
    Dim cnt As Long: cnt = 0

    For r = dataStart To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, cU).Value))
        If u = "" Then GoTo NextRow
        If Not MatchesAnyPattern(cfg, u) Then GoTo NextRow

        Dim nm As String: nm = Trim(CStr(ws.Cells(r, cN).Value))
        If nm = "" Or LCase(nm) = "vacant" Then GoTo NextRow

        Dim actRaw As Variant: actRaw = ws.Cells(r, cA).Value
        If Not IsNumeric(actRaw) Then GoTo NextRow
        If CDbl(actRaw) <= 0 Then GoTo NextRow

        ' Keep only past-date or non-date expiry (opposite of ReadYardi).
        Dim expCell As Variant: expCell = ws.Cells(r, cE).Value
        Dim isMTM As Boolean: isMTM = False
        Dim expiryVal As Variant
        Dim staleOut As Boolean: staleOut = False
        If IsDate(expCell) Then
            If CDate(expCell) < Date Then
                isMTM = True
                expiryVal = CDate(expCell)
                staleOut = (DateDiff("m", CDate(expCell), Date) > 15)
            End If
        ElseIf Trim(CStr(expCell)) <> "" Then
            isMTM = True
            expiryVal = Trim(CStr(expCell))
            staleOut = False
        End If
        If Not isMTM Then GoTo NextRow

        If Not dict.Exists(u) Then
            recs(cnt).Name = nm
            recs(cnt).FloorPlanCode = Trim(CStr(ws.Cells(r, cT).Value))
            recs(cnt).MarketRent = ws.Cells(r, cM).Value
            recs(cnt).ActualRent = CDbl(actRaw)
            recs(cnt).ExpiryVal = expiryVal
            recs(cnt).StaleOut = staleOut
            dict.Add u, cnt
            cnt = cnt + 1
        End If
NextRow:
    Next r

    Set ReadYardiMTM = dict
End Function
