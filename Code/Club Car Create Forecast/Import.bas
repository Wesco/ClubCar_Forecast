Attribute VB_Name = "Import"
Option Explicit

Sub ImportForecast()
    Dim Fcst As String
    Dim TotalRows As Long
    Fcst = Application.GetOpenFilename("ExportReport (*.xls; *.aspx), *.xls;*.aspx")

    On Error GoTo CANCEL_OPEN
    Workbooks.Open (Fcst)
    On Error GoTo 0

    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Temp").Range("A1")
    ActiveWorkbook.Close
    Worksheets("Temp").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row + 1

    Workbooks.Open FileName:="\\br3615gaps\gaps\Club Car\Southern ASM Forecast\ASM Forecast.xlsx"
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Temp").Cells(TotalRows, 1)
    ActiveWorkbook.Close

    Cells.WrapText = False
    Rows.EntireRow.AutoFit
    Columns.EntireColumn.AutoFit
    Delete Fcst
    Exit Sub

CANCEL_OPEN:
    Err.Raise Number:=53, Description:="User aborted import operation."

End Sub
