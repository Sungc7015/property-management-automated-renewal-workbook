Attribute VB_Name = "modAdmin"
Option Explicit

' ================================================================
'  modAdmin  -  one-time setup and ongoing health monitoring.
'
'  SetupWorkbook: adds 5 buttons to the Overview sheet.
'  HealthCheck:   verifies module version, config, sheets, names.
'
'  Version 2.1.0 - carved from modRenewalDynamic + modPropertySetup v1.2.1
' ================================================================

' ================================================================
'  SETUP WORKBOOK  -  run once after importing all 8 .bas files
' ================================================================
Public Sub SetupWorkbook()
    Dim hasConfig   As Boolean
    Dim hasSheetU   As Boolean
    Dim hasReaders  As Boolean
    Dim hasImport   As Boolean
    Dim hasDynamic  As Boolean
    Dim hasSetup    As Boolean
    Dim hasOverview As Boolean
    Dim hasAdmin    As Boolean
    Dim vbc As Object

    On Error GoTo NoTrust
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Select Case vbc.Name
            Case "modConfig":    hasConfig = True
            Case "modSheetUtils": hasSheetU = True
            Case "modReaders":   hasReaders = True
            Case "modImport":    hasImport = True
            Case "modDynamic":   hasDynamic = True
            Case "modSetup":     hasSetup = True
            Case "modOverview":  hasOverview = True
            Case "modAdmin":     hasAdmin = True
        End Select
    Next vbc
    On Error GoTo 0

    Dim missing As String: missing = ""
    If Not hasConfig Then missing = missing & "  modConfig.bas" & vbCrLf
    If Not hasSheetU Then missing = missing & "  modSheetUtils.bas" & vbCrLf
    If Not hasReaders Then missing = missing & "  modReaders.bas" & vbCrLf
    If Not hasImport Then missing = missing & "  modImport.bas" & vbCrLf
    If Not hasDynamic Then missing = missing & "  modDynamic.bas" & vbCrLf
    If Not hasSetup Then missing = missing & "  modSetup.bas" & vbCrLf
    If Not hasOverview Then missing = missing & "  modOverview.bas" & vbCrLf
    If Not hasAdmin Then missing = missing & "  modAdmin.bas" & vbCrLf

    If missing <> "" Then
        MsgBox "These modules must be imported before running Setup:" & vbCrLf & missing & vbCrLf & _
               "VBA Editor > File > Import File > select each .bas file", _
               vbExclamation, "Setup Error"
        Exit Sub
    End If

    ' Find or create the Overview sheet for buttons
    Dim ovName As String: ovName = ""
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Overview", vbTextCompare) > 0 Then
            ovName = ws.Name: Exit For
        End If
    Next ws
    If ovName = "" Then ovName = ThisWorkbook.Sheets(1).Name

    Dim ovWS As Worksheet: Set ovWS = ThisWorkbook.Sheets(ovName)

    AddButton ovWS, "btn_ImportMonthly", "Import Monthly Data", 10, 5, "modImport.ImportMonthlyData"
    AddButton ovWS, "btn_GenerateSheets", "Generate Month Sheets", 180, 5, "modSetup.GenerateMonthSheets"
    AddButton ovWS, "btn_SetupSheet", "Create Setup Sheet", 350, 5, "modSetup.CreateSetupSheet"
    AddButton ovWS, "btn_Overview", "Create Overview", 520, 5, "modOverview.CreateOverviewSheet"
    AddButton ovWS, "btn_HealthCheck", "Health Check", 690, 5, "modAdmin.HealthCheck"

    RefreshBufferCache

    On Error Resume Next
    ThisWorkbook.BuiltinDocumentProperties("Author").Value = "Christopher Sung"
    ThisWorkbook.BuiltinDocumentProperties("Comments").Value = _
        "Property renewal workbook system, version " & VER & ", created by Christopher Sung."
    On Error GoTo 0

    MsgBox "Setup complete!  (version " & VER & ")" & vbCrLf & vbCrLf & _
           "Buttons added to: " & ovName & vbCrLf & vbCrLf & _
           "Reminder: ensure the Workbook_SheetChange event in ThisWorkbook" & vbCrLf & _
           "calls modDynamic.HandleSheetChange, then run Create Setup Sheet.", _
           vbInformation, "Setup Complete"
    Exit Sub

NoTrust:
    MsgBox "Setup can't verify modules because VBA project access is blocked." & vbCrLf & vbCrLf & _
           "Enable: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "Macro Settings > 'Trust access to the VBA project object model'," & vbCrLf & _
           "then run Setup again.", vbExclamation, "Setup - VBA Access Needed"
End Sub

Private Sub AddButton(ws As Worksheet, btnName As String, caption As String, _
                       leftPos As Single, topPos As Single, macro As String)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = btnName Then shp.Delete: Exit For
    Next shp
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  Left:=leftPos, top:=topPos, Width:=160, Height:=24)
    With btn
        .Name = btnName
        .TextFrame.Characters.Text = caption
        With .TextFrame.Characters.Font
            .Name = "Arial": .Size = 10: .Bold = True
            .Color = RGB(255, 255, 255)
        End With
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Line.Visible = msoFalse
        .OnAction = macro
    End With
End Sub

' ================================================================
'  HEALTH CHECK
' ================================================================
Public Sub HealthCheck()
    Dim rpt As String, issues As Long: issues = 0

    ' Version - single constant from modConfig
    rpt = "MODULE VERSION: " & VER & vbCrLf & vbCrLf

    Dim cfg As PropConfig
    If Not LoadConfig(cfg, False) Then
        rpt = rpt & "PROPERTY SETUP" & vbCrLf & "  >> Setup sheet missing or invalid." & vbCrLf & vbCrLf
        issues = issues + 1
    End If

    If FindOverviewName() = "" Then
        rpt = rpt & "OVERVIEW" & vbCrLf & "  >> No overview sheet (run Create Overview)." & vbCrLf & vbCrLf
        issues = issues + 1
    End If

    Dim sheetCount As Long: sheetCount = 0
    Dim missingNames As String: missingNames = ""
    Dim shObj As Object, Sh As Worksheet, mm As Long, yy As Long
    For Each shObj In ThisWorkbook.Sheets
        If TypeOf shObj Is Worksheet Then
            Set Sh = shObj
            If ParseMonthSheet(Sh.Name, mm, yy) Then
                sheetCount = sheetCount + 1
                If yy > 0 Then
                    If Not NameExists(MonthYearPrefix(mm, yy) & ".Renewals") Then
                        missingNames = missingNames & "  " & Sh.Name & _
                                       " - no stats names (won't appear in overview)" & vbCrLf
                        issues = issues + 1
                    End If
                End If
            End If
        End If
    Next shObj
    rpt = rpt & "MONTH SHEETS: " & sheetCount & " found" & vbCrLf
    If missingNames <> "" Then rpt = rpt & missingNames
    rpt = rpt & vbCrLf

    Dim brokenN As String: brokenN = ""
    Dim nm As Object
    For Each nm In ThisWorkbook.Names
        If InStr(nm.RefersTo, "#REF") > 0 Then
            brokenN = brokenN & "  " & nm.Name & vbCrLf
            issues = issues + 1
        End If
    Next nm
    rpt = rpt & "DEFINED NAMES" & vbCrLf
    If brokenN <> "" Then
        rpt = rpt & "  >> Broken (#REF) - regenerate the affected month:" & vbCrLf & brokenN
    Else
        rpt = rpt & "  (all resolve)" & vbCrLf
    End If

    Dim head As String
    If issues = 0 Then
        head = "HEALTH CHECK - all clear" & vbCrLf & String(42, "-") & vbCrLf & vbCrLf
    Else
        head = "HEALTH CHECK - " & issues & " issue(s) found" & vbCrLf & String(42, "-") & vbCrLf & vbCrLf
    End If
    MsgBox head & rpt, IIf(issues = 0, vbInformation, vbExclamation), "Health Check"
End Sub
