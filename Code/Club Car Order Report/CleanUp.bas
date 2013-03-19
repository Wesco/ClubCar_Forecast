Attribute VB_Name = "CleanUp"
Option Explicit

Sub RemoveOldData()
    removefilter ("Bulk")
    removefilter ("Kit BOM")
    removefilter ("A Forecast")
    removefilter ("P Forecast")
    removefilter ("Gaps")
    removefilter ("Temp")
    removefilter ("Combined Forecast")
    removefilter ("Non-Stock Items")
    removefilter ("Forecast")
    removefilter ("PTableForecast")
    removefilter ("PTableKitParts")
    removefilter ("Hotsheet")

    Worksheets("Bulk").Range("F:T").Delete
    Worksheets("Kit BOM").Range("E:P").Delete
    Worksheets("Forecast").Cells.Delete
    Worksheets("Non-Stock Items").Cells.Delete
    Worksheets("A Forecast").Cells.Delete
    Worksheets("P Forecast").Cells.Delete
    Worksheets("Gaps").Cells.Delete
    Worksheets("Temp").Cells.Delete
    Worksheets("Combined Forecast").Cells.Delete
    Worksheets("PTableForecast").Cells.Delete
    Worksheets("PTableKitParts").Cells.Delete
    Worksheets("Hotsheet").Cells.Delete
    ActiveWorkbook.Save
End Sub

Sub removefilter(Optional wksht As String)
On Error Resume Next
    If wksht = "" Then    'If no arg is given clear the activesheet
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilterMode = False
        End If
    Else        'else clear the specified sheet
        If ActiveWorkbook.Worksheets(wksht).AutoFilterMode = True Then
            ActiveWorkbook.Worksheets(wksht).AutoFilterMode = False
        End If
    End If
    On Error GoTo 0
End Sub

