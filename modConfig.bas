Attribute VB_Name = "modConfig"
Option Explicit

' ================================================================
'  modConfig  -  shared constants, PropConfig type, and all config
'                loading / lookup helpers.
'
'  IMPORT ORDER: import this module FIRST. All other modules use
'  PropConfig and the public constants declared here.
'
'  Version 2.7.0
' ================================================================

Public Const VER As String = "2.7.0  (2026-07-07)"

Public Const SETUP_SHEET As String = "Property Setup"
Public Const BAR_GREY    As Long = 14277081     ' RGB(217,217,217)
Public Const INPUT_FILL  As Long = 13431551     ' RGB(255,242,204)

' ----------------------------------------------------------------
'  PROPERTY CONFIG TYPE  -  shared by all modules
'  (must be declared in exactly one standard module)
' ----------------------------------------------------------------
Public Type PropConfig
    FullName      As String
    ShortName     As String
    yr            As Long
    PatternCount  As Long
    Patterns()    As String
    MTMCap        As Double
    MTMThrough    As Date
    BufferRows    As Long
    GroupCount    As Long
    GroupNames()  As String
    CodeCount     As Long
    codes()       As String
    CodeGroupIdx() As Long
    RRUnit        As Long
    RRType        As Long
    RRName        As Long
    RRMarket      As Long
    RRActual      As Long
    RRExpiry      As Long
    GridUnit      As Long
    GridCurEff    As Long
    GridBest      As Long
    GridBestTerm  As Long
    GridNewLease  As Long
    MIUnit        As Long
    MIUnitType    As Long
    MIRent        As Long
    MIMoveIn      As Long
End Type

' ================================================================
'  LOAD CONFIG
' ================================================================
Public Function LoadConfig(cfg As PropConfig, Optional showErrors As Boolean = True) As Boolean
    LoadConfig = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SETUP_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        If showErrors Then MsgBox "No '" & SETUP_SHEET & "' sheet found." & vbCrLf & _
            "Run CreateSetupSheet first.", vbExclamation, "Property Setup"
        Exit Function
    End If

    On Error GoTo BadConfig

    cfg.FullName = Trim(CStr(NmRange("PS.PropertyName").Value))
    cfg.ShortName = Trim(CStr(NmRange("PS.ShortName").Value))
    cfg.yr = CLng(NmRange("PS.Year").Value)
    cfg.MTMCap = CDbl(NmRange("PS.MTMCap").Value)
    cfg.MTMThrough = CDate(NmRange("PS.MTMThrough").Value)
    cfg.BufferRows = CLng(NmRange("PS.BufferRows").Value)

    Dim patRaw As String
    patRaw = Trim(CStr(NmRange("PS.UnitPatterns").Value))
    Dim parts() As String: parts = Split(patRaw, ",")
    Dim i As Long, n As Long: n = 0
    ReDim cfg.Patterns(UBound(parts))
    For i = 0 To UBound(parts)
        If Trim(parts(i)) <> "" Then
            cfg.Patterns(n) = UCase(Trim(parts(i)))
            n = n + 1
        End If
    Next i
    cfg.PatternCount = n

    Dim top As Range: Set top = NmRange("PS.GroupsTop")
    Dim r As Long: r = top.Row
    Dim cnt As Long: cnt = 0
    ReDim cfg.GroupNames(49)
    Do While Trim(CStr(ws.Cells(r, top.Column).Value)) <> ""
        If cnt > 49 Then Exit Do
        cfg.GroupNames(cnt) = Trim(CStr(ws.Cells(r, top.Column).Value))
        cnt = cnt + 1: r = r + 1
    Loop
    cfg.GroupCount = cnt

    Set top = NmRange("PS.CodesTop")
    r = top.Row: cnt = 0
    ReDim cfg.codes(199): ReDim cfg.CodeGroupIdx(199)
    Do While Trim(CStr(ws.Cells(r, top.Column).Value)) <> ""
        If cnt > 199 Then Exit Do
        Dim cd As String, gp As String
        cd = LCase(Trim(CStr(ws.Cells(r, top.Column).Value)))
        gp = Trim(CStr(ws.Cells(r, top.Column + 1).Value))
        Dim gi As Long: gi = GroupIndex(cfg, gp)
        If gi < 0 Then
            If showErrors Then MsgBox "Setup error: Yardi code '" & cd & "' maps to group" & vbCrLf & _
                "'" & gp & "' which is not in the Floor Plan Groups list." & vbCrLf & _
                "(Group names must match exactly.)", vbExclamation, "Property Setup"
            Exit Function
        End If
        For i = 0 To cnt - 1
            If cfg.codes(i) = cd Then
                If showErrors Then MsgBox "Setup error: Yardi code '" & cd & "' is listed twice.", _
                    vbExclamation, "Property Setup"
                Exit Function
            End If
        Next i
        cfg.codes(cnt) = cd
        cfg.CodeGroupIdx(cnt) = gi
        cnt = cnt + 1: r = r + 1
    Loop
    cfg.CodeCount = cnt

    Set top = NmRange("PS.FallbacksTop")
    For r = top.Row To top.Row + 30
        Dim lbl As String: lbl = LCase(Trim(CStr(ws.Cells(r, top.Column).Value)))
        If lbl = "" Then Exit For
        Dim v As Long: v = 0
        If IsNumeric(ws.Cells(r, top.Column + 1).Value) Then v = CLng(ws.Cells(r, top.Column + 1).Value)
        Select Case lbl
            Case "rent roll: unit col":          cfg.RRUnit = v
            Case "rent roll: unit type col":     cfg.RRType = v
            Case "rent roll: resident col":      cfg.RRName = v
            Case "rent roll: market rent col":   cfg.RRMarket = v
            Case "rent roll: actual rent col":   cfg.RRActual = v
            Case "rent roll: lease expiry col":  cfg.RRExpiry = v
            Case "rents grid: unit col":         cfg.GridUnit = v
            Case "rents grid: cur eff rent col": cfg.GridCurEff = v
            Case "rents grid: best offer col":   cfg.GridBest = v
            Case "rents grid: best term col":    cfg.GridBestTerm = v
            Case "rents grid: new lease col":    cfg.GridNewLease = v
            Case "box score: unit col":          cfg.MIUnit = v
            Case "box score: unit type col":     cfg.MIUnitType = v
            Case "box score: rent col":          cfg.MIRent = v
            Case "box score: move-in col":       cfg.MIMoveIn = v
        End Select
    Next r

    Dim msg As String: msg = ""
    If cfg.FullName = "" Then msg = msg & "- Property Full Name is blank" & vbCrLf
    If cfg.GroupCount = 0 Then msg = msg & "- No floor plan groups listed" & vbCrLf
    If cfg.CodeCount = 0 Then msg = msg & "- No Yardi codes listed" & vbCrLf
    If cfg.PatternCount = 0 Then msg = msg & "- No unit number pattern" & vbCrLf
    If cfg.BufferRows < 1 Then msg = msg & "- Buffer Rows must be at least 1" & vbCrLf
    If cfg.MTMCap <= 0 Then msg = msg & "- MTM Cap % is blank or zero" & vbCrLf
    If msg <> "" Then
        If showErrors Then MsgBox "Fix the Property Setup sheet:" & vbCrLf & msg, _
            vbExclamation, "Property Setup"
        Exit Function
    End If

    Dim declared As Long
    On Error Resume Next
    declared = CLng(NmRange("PS.GroupCountCell").Value)
    On Error GoTo BadConfig
    If declared > 0 And declared <> cfg.GroupCount And showErrors Then
        MsgBox "Note: '# of Floor Plan Groups' says " & declared & " but the list has " & _
               cfg.GroupCount & ". Using the list (" & cfg.GroupCount & ").", _
               vbInformation, "Property Setup"
    End If

    LoadConfig = True
    Exit Function
BadConfig:
    If showErrors Then MsgBox "Could not read the Property Setup sheet (missing or broken" & vbCrLf & _
        "named ranges). Re-run CreateSetupSheet to repair it." & vbCrLf & vbCrLf & _
        "Detail: " & Err.Description, vbExclamation, "Property Setup"
End Function

Private Function NmRange(nm As String) As Range
    Set NmRange = ThisWorkbook.Names(nm).RefersToRange
End Function

' ================================================================
'  LOOKUP HELPERS  (used by modImport, modReaders, modSetup)
' ================================================================
Public Function GroupIndex(cfg As PropConfig, grpName As String) As Long
    Dim i As Long
    For i = 0 To cfg.GroupCount - 1
        If LCase(Trim(cfg.GroupNames(i))) = LCase(Trim(grpName)) Then
            GroupIndex = i: Exit Function
        End If
    Next i
    GroupIndex = -1
End Function

Public Function GetGroupForCode(cfg As PropConfig, code As String) As String
    Dim cd As String: cd = LCase(Trim(code))
    Dim i As Long
    For i = 0 To cfg.CodeCount - 1
        If cfg.codes(i) = cd Then
            GetGroupForCode = cfg.GroupNames(cfg.CodeGroupIdx(i))
            Exit Function
        End If
    Next i
    GetGroupForCode = ""
End Function

Public Function MatchesPattern(s As String, pat As String) As Boolean
    MatchesPattern = False
    If Len(s) <> Len(pat) Then Exit Function
    Dim i As Long, ch As String, pc As String
    For i = 1 To Len(pat)
        pc = Mid(pat, i, 1)
        ch = Mid(s, i, 1)
        Select Case pc
            Case "N": If ch < "0" Or ch > "9" Then Exit Function
            Case "A": If UCase(ch) < "A" Or UCase(ch) > "Z" Then Exit Function
            Case Else: If UCase(ch) <> pc Then Exit Function
        End Select
    Next i
    MatchesPattern = True
End Function

Public Function MatchesAnyPattern(cfg As PropConfig, s As String) As Boolean
    Dim i As Long
    For i = 0 To cfg.PatternCount - 1
        If MatchesPattern(Trim(s), cfg.Patterns(i)) Then
            MatchesAnyPattern = True: Exit Function
        End If
    Next i
    MatchesAnyPattern = False
End Function
