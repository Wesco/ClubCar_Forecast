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
    Worksheets("A Whse").Cells.Delete
    Worksheets("P Whse").Cells.Delete
    Worksheets("PivotTableA").Cells.Delete
    Worksheets("PivotTableP").Cells.Delete
    Worksheets("Temp").Cells.Delete
    Worksheets("Macro").Select
    ActiveWorkbook.Saved = True
    MsgBox "Complete!"
    Application.ScreenUpdating = True
    Application.Quit
End Sub

Sub Clean()
    Dim s As Worksheet
    
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.Cells.Delete
        End If
    Next
    
    Sheets("Macro").Select
    Range("C6").Select
End Sub
