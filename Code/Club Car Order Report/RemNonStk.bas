Attribute VB_Name = "RemNonStk"
Option Explicit

Sub RemoveNonStock()
    Dim iRows As Integer
    Dim iCounter As Integer

    Worksheets("Non-Stock Items").Select
    iRows = ActiveSheet.UsedRange.Rows.Count + 1

    Worksheets("Combined Forecast").Select
    ActiveSheet.Range("A:O").AutoFilter Field:=3, Criteria1:=""
    ActiveSheet.UsedRange.Copy Destination:=Worksheets("Non-Stock Items").Cells(iRows, 1)

    Worksheets("Non-Stock Items").Select

    Rows(iRows).Delete Shift:=xlUp
    Range(Cells(iRows, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3)).Delete Shift:=xlToLeft
    ActiveSheet.UsedRange.Columns.AutoFit

    Worksheets("Combined Forecast").Select
    ActiveSheet.AutoFilterMode = False

    Worksheets("Forecast").Select
    iRows = ActiveSheet.UsedRange.Rows.Count

    For iCounter = iRows To 2 Step -1
        If Cells(iCounter, 3).Value = "" Then
            Rows(iCounter).Delete
        End If
    Next

    Worksheets("Bulk").Select
End Sub

