Attribute VB_Name = "modMTM"
Option Explicit

' ================================================================
'  modMTM  -  MTM Tracker Sheet refresh handler.
'
'  Reads: modConfig (PropConfig, LoadConfig)
'         modReaders (PickFile, ReadYardiMTM)
'         modSheetUtils (SheetExists)
'
'  Version 2.1.0
' ================================================================

Private Const MTM_SHEET As String = "MTM"

' ----------------------------------------------------------------
'  RefreshMTMSheet  -  button handler. Prompts for Yardi Rent Roll,
'                      refreshes the MTM tracker sheet.
' ----------------------------------------------------------------
Public Sub RefreshMTMSheet()
    Dim cfg As PropConfig
    If Not LoadConfig(cfg, True) Then Exit Sub

    If Not SheetExists(MTM_SHEET) Then
        MsgBox "Sheet """ & MTM_SHEET & """ not found in this workbook.", vbExclamation, "MTM Refresh"
        Exit Sub
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

    WriteHeaders ws
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
'  WriteHeaders  -  ensures row 1 has the correct column headers.
' ----------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Unit", "Name", "Floor Plan", "Lease Expiry", _
                    "Current Rent", "Market Rent", "Next Increase", _
                    "Status", "Notes")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

' ----------------------------------------------------------------
'  DoRefreshMTM  -  core merge logic: update existing rows,
'                   flag gone units, add new units, sort.
' ----------------------------------------------------------------
Private Sub DoRefreshMTM(ws As Worksheet, mtmDict As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Build unit -> row index map from existing sheet data
    Dim sheetRows As Object: Set sheetRows = CreateObject("Scripting.Dictionary")
    sheetRows.CompareMode = 1
    Dim r As Long
    For r = 2 To lastRow
        Dim u As String: u = Trim(CStr(ws.Cells(r, 1).Value))
        If u <> "" Then sheetRows(u) = r
    Next r

    ' Pass 1 — update rows for units still in mtmDict; collect new units
    Dim newUnits As Object: Set newUnits = CreateObject("Scripting.Dictionary")
    newUnits.CompareMode = 1
    Dim key As Variant
    For Each key In mtmDict.Keys
        Dim arr As Variant: arr = mtmDict(key)
        If sheetRows.Exists(key) Then
            r = sheetRows(key)
            ws.Cells(r, 2).Value = CStr(arr(0))
            ws.Cells(r, 3).Value = CStr(arr(1))
            If IsDate(arr(4)) Then ws.Cells(r, 4).Value = CDate(arr(4))
            ws.Cells(r, 5).Value = CDbl(arr(3))
            If IsNumeric(arr(2)) Then ws.Cells(r, 6).Value = CDbl(arr(2))
            If IsDate(arr(4)) Then
                Dim nd As Variant: nd = NextIncreaseDate(CDate(arr(4)))
                If Not IsNull(nd) Then ws.Cells(r, 7).Value = CDate(nd)
            End If
            ' Preserve col H (Status) and col I (Notes)
        Else
            newUnits(key) = arr
        End If
    Next key

    ' Pass 2 — flag units on sheet that are no longer in mtmDict
    For Each key In sheetRows.Keys
        If Not mtmDict.Exists(key) Then
            If Trim(CStr(ws.Cells(sheetRows(key), 8).Value)) = "Active MTM" Then
                ws.Cells(sheetRows(key), 8).Value = ChrW(9888) & " Review - may have renewed"
            End If
        End If
    Next key

    ' Pass 3 — append new units
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For Each key In newUnits.Keys
        arr = newUnits(key)
        ws.Cells(nextRow, 1).Value = CStr(key)
        ws.Cells(nextRow, 2).Value = CStr(arr(0))
        ws.Cells(nextRow, 3).Value = CStr(arr(1))
        If IsDate(arr(4)) Then ws.Cells(nextRow, 4).Value = CDate(arr(4))
        ws.Cells(nextRow, 5).Value = CDbl(arr(3))
        If IsNumeric(arr(2)) Then ws.Cells(nextRow, 6).Value = CDbl(arr(2))
        If IsDate(arr(4)) Then
            Dim nd2 As Variant: nd2 = NextIncreaseDate(CDate(arr(4)))
            If Not IsNull(nd2) Then ws.Cells(nextRow, 7).Value = CDate(nd2)
        End If
        ws.Cells(nextRow, 8).Value = "Active MTM"
        nextRow = nextRow + 1
    Next key

    ' Sort by col G (Next Increase) ascending; blanks go to bottom
    Dim lastRowAfter As Long
    lastRowAfter = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowAfter >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRowAfter, 9)).Sort _
            Key1:=ws.Cells(2, 7), Order1:=xlAscending, Header:=xlNo
    End If
End Sub

' ----------------------------------------------------------------
'  NextIncreaseDate  -  next 12-month anniversary of leaseExpiry
'                       that falls after today.
' ----------------------------------------------------------------
Private Function NextIncreaseDate(leaseExpiry As Date) As Variant
    Dim n As Long: n = 1
    Dim d As Date
    Do While n <= 100
        d = DateAdd("m", n * 12, leaseExpiry)
        If d > Date Then NextIncreaseDate = d: Exit Function
        n = n + 1
    Loop
    NextIncreaseDate = Null
End Function
