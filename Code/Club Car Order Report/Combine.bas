Attribute VB_Name = "Combine"
Option Explicit

Sub PTableAP()
    Dim TotalRows As Long
    Dim TotalCols As Long
    Dim aPTableFields() As String
    Dim PrevDispAlert As Boolean
    Dim i As Integer

    'Get P forecast data
    Worksheets("P Forecast").Select
    PrevDispAlert = Application.DisplayAlerts
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Columns(TotalCols).Delete   'Remove Totals column
    Columns(2).Delete           'Remove descriptions column
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Temp").Range("A1")

    'Get A forecast data
    Worksheets("A Forecast").Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    Columns(TotalCols).Delete   'Remove Totals column
    Columns(2).Delete           'Remove descriptions column
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Temp").Range("A" & TotalRows + 1)

    Worksheets("Temp").Select
    Rows(TotalRows + 1).Delete  'Remove A forecast header
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Load column headers into an array
    ReDim aPTableFields(1 To TotalCols)
    For i = 1 To TotalCols
        aPTableFields(i) = Cells(1, i).Text
    Next

    'Sort item numbers smallest to largest
    With ActiveWorkbook.Worksheets("Temp").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortTextAsNumbers
        .SetRange Range("A1:M" & TotalRows)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Create a pivot table out of the combined forecasts
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                      SourceData:=Sheets("Temp").Range("A1:M" & TotalRows), _
                                      Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:=Sheets("PTableForecast").Range("A1"), _
                                      TableName:="PTableCombined", _
                                      DefaultVersion:=xlPivotTableVersion14

    Sheets("PTableForecast").Select
    
    'Add fields to the pivot table
    With ActiveSheet.PivotTables("PTableCombined")
        .PivotFields(aPTableFields(1)).Orientation = xlRowField
        .PivotFields(aPTableFields(1)).Position = 1
        For i = 2 To UBound(aPTableFields)
            .AddDataField ActiveSheet.PivotTables("PTableCombined").PivotFields(aPTableFields(i)), "Sum of " & Format(aPTableFields(i), "mmm"), xlSum
        Next
    End With
    
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    
    'Get data from the pivot table and store it as values
    Application.DisplayAlerts = False
    Range("A1:M" & TotalRows).Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Application.DisplayAlerts = PrevDispAlert

    'Fix the column headers
    Range("A1:M1").Value = aPTableFields

    'Remove the totals
    Rows(TotalRows).Delete
    ActiveSheet.AutoFilterMode = False

    'Copy the combined data to the temp worksheet
    Sheets("Temp").Cells.Delete
    Range("A1:M" & TotalRows).Copy Destination:=Sheets("Temp").Range("A1")
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
