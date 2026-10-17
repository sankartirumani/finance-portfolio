Attribute VB_Name = "SapExtraction"
Option Explicit

'=============================================================================
' SAP report extraction via SAP GUI Scripting (sanitized, generic version)
'
' Every organisation-specific value (transaction code, controlling area,
' report variant, drill-down labels, file paths) is read from the "Settings"
' sheet - nothing company-specific is hard-coded here.
'
' Requires: SAP GUI for Windows with scripting enabled (client + server side,
' by your SAP admin) and a saved SAP Logon shortcut (.sap file).
'=============================================================================

Public session As Object

' Layout of the Settings sheet (column B = value). Row numbers are the contract
' between the workbook and this code - change both together.
Private Const ROW_SHORTCUT As Long = 2      ' full path to the .sap shortcut
Private Const ROW_SELECTION_DIR As Long = 3 ' folder holding the saved selection file
Private Const ROW_SELECTION_FILE As Long = 4
Private Const ROW_YEAR As Long = 5
Private Const ROW_PERIOD As Long = 6
Private Const ROW_OUT_DIR As Long = 8
Private Const ROW_OUT_FILE As Long = 9
Private Const ROW_TCODE As Long = 10
Private Const ROW_CTRL_AREA As Long = 11
Private Const ROW_LAYOUT As Long = 12       ' exact ALV layout/variant name
Private Const ROW_VALUE_TYPE As Long = 13   ' e.g. cost element / account filter
Private Const ROW_DRILL_TEXT As Long = 14   ' text searched for when drilling down
Private Const ROW_KILL As Long = 15         ' "yes" = close SAP when finished

'--------------------------------------------------------------------------
' ENTRY POINT - one button on the sheet is assigned to this macro
'--------------------------------------------------------------------------
Public Sub RunExtraction(ByVal settingsSheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(settingsSheetName)

    StartSapShortcut CStr(ws.Cells(ROW_SHORTCUT, 2).Value)
    Application.Wait Now + TimeValue("0:00:10")      ' let SAP Logon + SSO finish

    AttachToSession
    RunReport ws
    ExportWithLayout CStr(ws.Cells(ROW_LAYOUT, 2).Value), _
                     CStr(ws.Cells(ROW_OUT_DIR, 2).Value), _
                     CStr(ws.Cells(ROW_OUT_FILE, 2).Value)

    If LCase$(CStr(ws.Cells(ROW_KILL, 2).Value)) = "yes" Then CloseSapSession
End Sub

'--------------------------------------------------------------------------
' Start SAP by "double-clicking" the saved shortcut. The shortcut carries the
' system connection and single sign-on (SNC) settings, so this code never
' sees or stores a username or password.
'--------------------------------------------------------------------------
Private Sub StartSapShortcut(ByVal shortcutPath As String)
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run Chr(34) & shortcutPath & Chr(34), 1, False
    Application.Wait Now + TimeValue("0:00:05")
End Sub

Private Sub AttachToSession()
    Dim gui As Object, conn As Object
    Set gui = GetObject("SAPGUI").GetScriptingEngine
    Set conn = gui.Children(0)
    Set session = conn.Children(0)
End Sub

'--------------------------------------------------------------------------
' Run the report: open transaction, set the controlling area, fill selection
' criteria. The multi-value selection is loaded from a text file that SAP
' itself saved earlier ("upload from text file" in the multiple-selection
' popup) so long lists never have to be typed by the macro.
'--------------------------------------------------------------------------
Private Sub RunReport(ws As Worksheet)
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = CStr(ws.Cells(ROW_TCODE, 2).Value)
    session.FindById("wnd[0]").SendVKey 0

    ' Controlling-area popup (appears the first time in a session)
    session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = CStr(ws.Cells(ROW_CTRL_AREA, 2).Value)
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press

    session.FindById("wnd[0]/usr/txt$1GJAHR").Text = CStr(ws.Cells(ROW_YEAR, 2).Value)
    session.FindById("wnd[0]/usr/ctxt$1PERIB").Text = CStr(ws.Cells(ROW_PERIOD, 2).Value)
    session.FindById("wnd[0]/usr/ctxt$1KSTAR").Text = CStr(ws.Cells(ROW_VALUE_TYPE, 2).Value)

    ' Load the multi-value selection from the saved text file
    session.FindById("wnd[0]/usr/btn%__1KOSET_%_APP_%-VALU_PUSH").Press
    session.FindById("wnd[1]/tbar[0]/btn[23]").Press                  ' import from file
    session.FindById("wnd[2]/usr/ctxtDY_PATH").Text = CStr(ws.Cells(ROW_SELECTION_DIR, 2).Value)
    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").Text = CStr(ws.Cells(ROW_SELECTION_FILE, 2).Value)
    session.FindById("wnd[2]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press                   ' copy + close

    session.FindById("wnd[0]/tbar[1]/btn[8]").Press                   ' execute (F8)
    Application.Wait Now + TimeValue("0:00:10")                       ' large result sets are slow

    DrillTo CStr(ws.Cells(ROW_DRILL_TEXT, 2).Value)
End Sub

' Use the report's own Find (Ctrl+F) to land on a labelled line, then drill in.
' NOTE: the label coordinates (lbl[x,y]) were captured with the SAP script
' recorder and are the most fragile part - see docs/lessons-learned.md.
Private Sub DrillTo(ByVal labelText As String)
    session.FindById("wnd[0]").SendVKey 71                            ' Ctrl+F
    session.FindById("wnd[1]/usr/txtRSYSF-STRING").Text = labelText
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    session.FindById("wnd[2]").SendVKey 2                             ' double-click hit
    session.FindById("wnd[0]/tbar[1]/btn[33]").Press                  ' open layout selector
    Application.Wait Now + TimeValue("0:00:05")
End Sub

'--------------------------------------------------------------------------
' Pick an ALV layout by EXACT name. SAP's Find matches substrings, so we keep
' pressing "find next" until the highlighted row equals the requested name.
' If it is never found we abort instead of exporting the wrong layout.
'--------------------------------------------------------------------------
Private Function SelectExactLayout(grid As Object, ByVal variantName As String) As Boolean
    grid.CurrentCellRow = -1
    grid.SelectColumn "VARIANT"
    grid.ContextMenu
    grid.SelectContextMenuItem "&FIND"
    session.FindById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = variantName

    Dim found As Boolean, i As Integer
    For i = 1 To 30
        session.FindById("wnd[2]").SendVKey 0
        If Trim$(grid.GetCellValue(grid.CurrentCellRow, "VARIANT")) = variantName Then
            found = True
            Exit For
        End If
    Next i
    session.FindById("wnd[2]").SendVKey 12                            ' close search popup

    If found Then
        grid.SelectedRows = CStr(grid.CurrentCellRow)
        grid.ClickCurrentCell
    Else
        MsgBox "Layout '" & variantName & "' not found - export aborted.", vbExclamation
    End If
    SelectExactLayout = found
End Function

Private Sub ExportWithLayout(ByVal variantName As String, ByVal outDir As String, ByVal outFile As String)
    Dim grid As Object
    Set grid = session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
    If Not SelectExactLayout(grid, variantName) Then Exit Sub

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select   ' List > Export > Spreadsheet

    ' Some systems show an extra "format" popup first, others go straight to the
    ' file dialog - probe for the path field and confirm the popup if it is absent.
    Dim pathField As Object
    On Error Resume Next
    Set pathField = session.FindById("wnd[1]/usr/ctxtDY_PATH")
    On Error GoTo 0
    If pathField Is Nothing Then session.FindById("wnd[1]/tbar[0]/btn[0]").Press

    session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = outDir
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = outFile
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press                  ' save

    On Error Resume Next                                              ' "overwrite?" popup
    session.FindById("wnd[1]/usr/btnBUTTON_1").Press
    On Error GoTo 0

    Application.Wait Now + TimeValue("0:00:03")
End Sub

'--------------------------------------------------------------------------
' Close SAP. WARNING: taskkill ends ALL SAP Logon windows on the machine.
'--------------------------------------------------------------------------
Public Sub CloseSapSession()
    On Error Resume Next
    If Not session Is Nothing Then session.FindById("wnd[0]").Close
    CreateObject("WScript.Shell").Run "taskkill /IM saplogon.exe /F", 0, True
    Application.Wait Now + TimeValue("0:00:03")
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
' Run two extractions back to back, restarting SAP between them (keeps the
' GUI responsive after a heavy report) and stamping completion times.
'--------------------------------------------------------------------------
Public Sub RunAll()
    Dim names As Variant, i As Long
    names = Array("ExtractA", "ExtractB")        ' sheet names of your settings sheets
    For i = LBound(names) To UBound(names)
        RunExtraction CStr(names(i))
        ThisWorkbook.Sheets("PARAM").Cells(4 + i, 2).Value = Now
        CloseSapSession
        Application.Wait Now + TimeValue("0:00:05")
    Next i
    MsgBox "Extraction completed.", vbInformation
End Sub
