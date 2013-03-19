Attribute VB_Name = "Combine"
Option Explicit

Sub PTableAP()
    Dim iRowCount As Integer, i As Integer
    Dim rColHeaders As Range
    Dim aPTableFields() As String
    Dim vCell As Variant

    Worksheets("P Forecast").Select
    Range("A1").Select
    Columns(ActiveCell.CurrentRegion.Columns.Count).Delete
    Columns(2).Delete
    With Range("A:M")
        iRowCount = .CurrentRegion.Rows.Count
        Range(Cells(1, 1), Cells(iRowCount, .CurrentRegion.Columns.Count)).Copy Destination:=Worksheets("Temp").Range("A1")
    End With
    Application.CutCopyMode = False
    Worksheets("A Forecast").Select
    Range("A1").Select
    Columns(ActiveCell.CurrentRegion.Columns.Count).Delete
    Columns(2).Delete
    With Range("A:M")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Copy Destination:=Worksheets("Temp").Cells(iRowCount + 1, 1)
    End With
    Application.CutCopyMode = False
    Worksheets("Temp").Select
    Rows(iRowCount + 1).Delete
    Set rColHeaders = Range("A1:M1")
    ReDim aPTableFields(1 To rColHeaders.Columns.Count)
    For Each vCell In rColHeaders
        i = i + 1
        aPTableFields(i) = vCell.Text
    Next
    With Range("A:M")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Select
    End With
    With Selection
        .AutoFilter
        ActiveWorkbook.Worksheets("Temp").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Temp").AutoFilter.Sort.SortFields.Add _
                Key:=Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 1)), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
    End With
    With ActiveWorkbook.Worksheets("Temp").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    With Selection
        ActiveWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="Temp!" & Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Address, _
                Version:=xlPivotTableVersion14).CreatePivotTable _
                TableDestination:="PTableForecast!R3C1", TableName:="PTableCombined", _
                DefaultVersion:=xlPivotTableVersion14
    End With
    Worksheets("PTableForecast").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(1))
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PTableCombined")
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(2)), "Sum of Aug", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(3)), "Sum of Sep", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(4)), "Sum of Oct", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(5)), "Sum of Nov", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(6)), "Sum of Dec", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(7)), "Sum of Jan", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(8)), "Sum of Feb", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(9)), "Sum of Mar", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(10)), "Sum of Apr", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(11)), "Sum of May", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(12)), "Sum of Jun", xlSum
        .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(13)), "Sum of Jul", xlSum
    End With
    Rows("1:2").Delete
    With Range("A:A")
        Application.DisplayAlerts = False
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Copy
        .PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Application.DisplayAlerts = True
        Range("A1:M1").Value = aPTableFields
        Rows(.CurrentRegion.Rows.Count).Delete
    End With
    removefilter ("Temp")
    Worksheets("Temp").Cells.Delete
    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Copy _
                Destination:=Worksheets("Temp").Range("A1")
        Application.CutCopyMode = False
    End With
End Sub

Sub FilterNS()
    Dim i As Long, n As Long

    Worksheets("Temp").Select
    With Range("A:M")
        .CurrentRegion.Cut Destination:=Worksheets("Temp").Range("B1")
    End With
    Application.CutCopyMode = False
    Range("A1").Value = "Sim_num"
    Range("A2").Formula = "=VLOOKUP(B2, master!A:B,2,FALSE)"
    With Range("B:B")
        Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(.CurrentRegion.Rows.Count, 1))
    End With
    With Range("A:A")
        .CurrentRegion.Value = .CurrentRegion.Value
    End With
    With Range("A:N")
        .CurrentRegion.AutoFilter Field:=1, Criteria1:="=Non-Stock", Operator:=xlOr, Criteria2:="=#N/A"
        .CurrentRegion.Copy Destination:=Worksheets("Non-Stock Items").Range("A1")
        .CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Select
    End With
    With Range("A:A").SpecialCells(xlCellTypeVisible)
        i = .SpecialCells(xlCellTypeConstants).Count
    End With
    Selection.ClearContents
    Application.CutCopyMode = False
    removefilter
    Range("A1").Select

    Do While ActiveCell.Text <> ""
        ActiveCell.Offset(1, 0).Select
        Do While ActiveCell.Text = ""
            If i > n Then
                n = n + 1
                Rows(ActiveCell.Row).Delete
            Else
                Exit Do
            End If
        Loop
    Loop
    Range("A:N").SpecialCells(xlCellTypeConstants).Copy Destination:=Worksheets("Combined Forecast").Range("A1")
    removefilter ("Temp")
    Worksheets("Temp").Cells.Delete
End Sub


