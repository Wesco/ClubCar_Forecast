Attribute VB_Name = "RemNonStk"
Option Explicit

Sub RemoveNonStock()
    Dim iRows As Integer
    Dim iCounter As Integer

    Worksheets("Non-Stock Items").Select
    iRows = ActiveSheet.UsedRange.Rows.Count + 1

    Worksheets("Combined Forecast").Select
    ActiveSheet.Range("A:O").AutoFilter Field:=3, Criteria1:="#N/A"
    ActiveSheet.UsedRange.Copy Destination:=Worksheets("Non-Stock Items").Cells(iRows, 1)

    Worksheets("Non-Stock Items").Select

    Rows(iRows).Delete Shift:=xlUp
    Range(Cells(iRows, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3)).Delete Shift:=xlToLeft
    ActiveSheet.UsedRange.Select
    Selection.EntireColumn.AutoFit
    Range("A1").Select

    Worksheets("Combined Forecast").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("A1").Select

    Worksheets("Forecast").Select
    iRows = ActiveSheet.UsedRange.Rows.Count
    Range("C1").Select

    For iCounter = 1 To iRows
        If ActiveCell.Text = "#N/A" Then
            Rows(ActiveCell.Row).Delete Shift:=xlUp
        End If
        If ActiveCell.Text <> "#N/A" Then
            ActiveCell.Offset(1, 0).Select
        End If
    Next

    Worksheets("Bulk").Select
End Sub

