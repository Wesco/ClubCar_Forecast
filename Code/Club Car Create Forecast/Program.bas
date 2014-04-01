Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Application.ScreenUpdating = False
    ImportForecast
    DatesToCol
    SeparateAP
    CreatePTable "A Whse", "PTable1", "PivotTableA"
    CreatePTable "P Whse", "PTable2", "PivotTableP"
    FormatPivTable "PivotTableA"
    FormatPivTable "PivotTableP"
    SaveForecast
    Clean
    ActiveWorkbook.Saved = True
    MsgBox "Complete!"
    Application.ScreenUpdating = True
    Application.Quit
End Sub

Sub Clean()
    Dim s As Worksheet
    Dim PrevDispAlert As Boolean
    
    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next
    
    Application.DisplayAlerts = PrevDispAlert
    
    Sheets("Macro").Select
    Range("C6").Select
End Sub
