Attribute VB_Name = "Stores"
Option Explicit

'==============================================================
' STORES module - multi-sheet view toggles
'
' Refactored 2026-06-08:
'   - Every button-macro now applies to ALL dashboard sheets
'   - Per-sheet logic is parameterised by Worksheet, no ActiveSheet
'   - Hidden Data sheets are skipped automatically
'   - ScreenUpdating + Calculation are saved and restored
'   - Missing shapes/ranges fail silently (sheet keeps going)
'
' Public macros (button-callable, names preserved so existing
' button assignments keep working):
'   GrowthPercent       Growthpercentmin
'   GrowthActual        GrowthMin
'   GrowthPerGuest
'   PnL                 PnLmin
'   GSTOP               GSTOPGraph
'==============================================================


'--- PUBLIC: button-callable macros (one-line wrappers) ---

Sub GrowthPercent()
    ApplyToDashboards "GrowthPercent"
End Sub

Sub Growthpercentmin()
    ApplyToDashboards "GrowthPercentMin"
End Sub

Sub GrowthActual()
    ApplyToDashboards "GrowthActual"
End Sub

Sub GrowthMin()
    ApplyToDashboards "GrowthMin"
End Sub

Sub GrowthPerGuest()
    ApplyToDashboards "GrowthPerGuest"
End Sub

Sub PnL()
    ApplyToDashboards "PnL"
End Sub

Sub PnLmin()
    ApplyToDashboards "PnLmin"
End Sub

Sub GSTOP()
    ApplyToDashboards "GSTOP"
End Sub

Sub GSTOPGraph()
    ApplyToDashboards "GSTOPGraph"
End Sub


'--- PRIVATE: loop driver ---

Private Sub ApplyToDashboards(mode As String)
    Dim ws As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevCalc As XlCalculation

    prevScreenUpdating = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanExit

    For Each ws In ThisWorkbook.Worksheets
        If IsDashboardSheet(ws) Then
            ApplyMode ws, mode
        End If
    Next ws

CleanExit:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
End Sub


'--- PRIVATE: dashboard detection ---
' A "dashboard" sheet is visible AND has the Growth$ or PnL shape.
' Hidden Data sheets, Side-By-Side (if it lacks those shapes), and
' any future helper sheet are skipped automatically.

Private Function IsDashboardSheet(ws As Worksheet) As Boolean
    If ws.Visible <> xlSheetVisible Then
        IsDashboardSheet = False
        Exit Function
    End If

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("Growth$")
    If shp Is Nothing Then Set shp = ws.Shapes("PnL")
    On Error GoTo 0

    IsDashboardSheet = Not (shp Is Nothing)
End Function


'--- PRIVATE: dispatch ---

Private Sub ApplyMode(ws As Worksheet, mode As String)
    Select Case mode
        Case "GrowthPercent":    DoGrowthPercent ws
        Case "GrowthPercentMin": DoGrowthPercentMin ws
        Case "GrowthActual":     DoGrowthActual ws
        Case "GrowthMin":        DoGrowthMin ws
        Case "GrowthPerGuest":   DoGrowthPerGuest ws
        Case "PnL":              DoPnL ws
        Case "PnLmin":           DoPnLmin ws
        Case "GSTOP":            DoGSTOP ws
        Case "GSTOPGraph":       DoGSTOPGraph ws
    End Select
End Sub


'--- PRIVATE: per-mode workers ---

Private Sub DoGrowthPercent(ws As Worksheet)
    SetWidth ws, "E7,U7,AK7", 8
    SetShape ws, "Growth%", False
    SetShape ws, "Growth%min", True
    SetShape ws, "Border Growth1", True
    SetShape ws, "Border Growth2", True
    SetShape ws, "Border Growth3", True
    SetShape ws, "Border GrowthT1", True
    SetShape ws, "Border GrowthT2", True
    SetShape ws, "Border GrowthT3", True
    SetShape ws, "Border %1", True
    SetShape ws, "Border %2", True
    SetShape ws, "Border %3", True
End Sub

Private Sub DoGrowthPercentMin(ws As Worksheet)
    SetWidth ws, "E7,U7,AK7", 0
    SetShape ws, "Growth%", True
    SetShape ws, "Growth%min", False
    SetShape ws, "Border Growth1", False
    SetShape ws, "Border Growth2", False
    SetShape ws, "Border Growth3", False
    SetShape ws, "Border GrowthT1", False
    SetShape ws, "Border GrowthT2", False
    SetShape ws, "Border GrowthT3", False
    SetShape ws, "Border %1", False
    SetShape ws, "Border %2", False
    SetShape ws, "Border %3", False
End Sub

Private Sub DoGrowthActual(ws As Worksheet)
    SetWidth ws, "F7,V7,AL7", 20
    SetWidth ws, "G7,W7,AM7", 0
    SetShape ws, "Growth$", False
    SetShape ws, "Growth$min", True
    SetShape ws, "GrowthPerGuest", True
    SetShape ws, "GrowthActual", False
    SetShape ws, "GrowthPerGuestq", True
    SetShape ws, "GrowthActualq", False
    SetShape ws, "GrowthPerGuesty", True
    SetShape ws, "GrowthActualy", False
    SetShape ws, "Border Growth1", True
    SetShape ws, "Border Growth2", True
    SetShape ws, "Border Growth3", True
    SetShape ws, "Border GrowthT1", True
    SetShape ws, "Border GrowthT2", True
    SetShape ws, "Border GrowthT3", True
End Sub

Private Sub DoGrowthMin(ws As Worksheet)
    SetWidth ws, "F8,G8,V8,W8,AL8,AM8", 0
    SetShape ws, "Growth$", True
    SetShape ws, "Growth$min", False
    SetShape ws, "GrowthPerGuest", False
    SetShape ws, "GrowthActual", False
    SetShape ws, "GrowthPerGuestq", False
    SetShape ws, "GrowthActualq", False
    SetShape ws, "GrowthPerGuesty", False
    SetShape ws, "GrowthActualy", False
    SetShape ws, "Border Growth1", False
    SetShape ws, "Border Growth2", False
    SetShape ws, "Border Growth3", False
    SetShape ws, "Border GrowthT1", False
    SetShape ws, "Border GrowthT2", False
    SetShape ws, "Border GrowthT3", False
End Sub

Private Sub DoGrowthPerGuest(ws As Worksheet)
    SetWidth ws, "G7,W7,AM7", 20
    SetWidth ws, "F7,V7,AL7", 0
    SetShape ws, "Growth$", False
    SetShape ws, "Growth$min", True
    SetShape ws, "GrowthPerGuest", False
    SetShape ws, "GrowthActual", True
    SetShape ws, "GrowthPerGuestq", False
    SetShape ws, "GrowthActualq", True
    SetShape ws, "GrowthPerGuesty", False
    SetShape ws, "GrowthActualy", True
End Sub

Private Sub DoPnL(ws As Worksheet)
    SetWidth ws, "J8,L8,O8,Q8,Z8,AB8,AE8,AG8,AP8,AR8,AU8,AW8", 8
    SetWidth ws, "M8,R8,AC8,AH8,AS8,AX8", 6
    SetShape ws, "PnL", False
    SetShape ws, "PnLMin", True
    SetShape ws, "GSTOPGraph", True
    SetShape ws, "GSTOP", False
    SetShape ws, "GSTOPGraphq", True
    SetShape ws, "GSTOPq", False
    SetShape ws, "GSTOPGraphy", True
    SetShape ws, "GSTOPy", False
    SetShape ws, "Border PnL1", True
    SetShape ws, "Border PnL2", True
    SetShape ws, "Border PnL3", True
End Sub

Private Sub DoPnLmin(ws As Worksheet)
    SetWidth ws, "I8:R8,Y8:AH8,AO8:AX8", 0
    SetShape ws, "PnL", True
    SetShape ws, "PnLmin", False
    SetShape ws, "GSTOPGraph", False
    SetShape ws, "GSTOP", False
    SetShape ws, "GSTOPGraphq", False
    SetShape ws, "GSTOPq", False
    SetShape ws, "GSTOPGraphy", False
    SetShape ws, "GSTOPy", False
    SetShape ws, "Border PnL1", False
    SetShape ws, "Border PnL2", False
    SetShape ws, "Border PnL3", False
End Sub

Private Sub DoGSTOP(ws As Worksheet)
    SetWidth ws, "I8,K8,N8,P8,Y8,AA8,AD8,AF8,AO8,AQ8,AT8,AV8", 8
    SetWidth ws, "J8,L8,O8,Q8,Z8,AB8,AE8,AG8,AP8,AR8,AU8,AW8", 0
    SetShape ws, "GSTOPGRAPH", False
    SetShape ws, "GSTOP", True
    SetShape ws, "GSTOPGRAPHq", False
    SetShape ws, "GSTOPq", True
    SetShape ws, "GSTOPGRAPHy", False
    SetShape ws, "GSTOPy", True
End Sub

Private Sub DoGSTOPGraph(ws As Worksheet)
    SetWidth ws, "I8,K8,N8,P8,Y8,AA8,AD8,AF8,AO8,AQ8,AT8,AV8", 0
    SetWidth ws, "J8,L8,O8,Q8,Z8,AB8,AE8,AG8,AP8,AR8,AU8,AW8", 8
    SetShape ws, "GSTOPGRAPH", True
    SetShape ws, "GSTOP", False
    SetShape ws, "GSTOPGRAPHq", True
    SetShape ws, "GSTOPq", False
    SetShape ws, "GSTOPGRAPHy", True
    SetShape ws, "GSTOPy", False
End Sub


'--- PRIVATE: safe setters (skip if shape/range missing on a sheet) ---

Private Sub SetShape(ws As Worksheet, shapeName As String, visible As Boolean)
    On Error Resume Next
    ws.Shapes(shapeName).Visible = visible
    On Error GoTo 0
End Sub

Private Sub SetWidth(ws As Worksheet, addr As String, w As Double)
    On Error Resume Next
    ws.Range(addr).ColumnWidth = w
    On Error GoTo 0
End Sub
