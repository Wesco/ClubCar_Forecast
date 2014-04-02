Attribute VB_Name = "Combine"
Option Explicit

Sub PTableAP()
    Dim TotalRows As Long
    Dim TotalCols As Integer
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
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long
    Dim n As Long

    Worksheets("Temp").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1
    Columns(1).Insert

    'Add SIMs
    Range("A1").Value = "SIM"
    Range("A2:A" & TotalRows).Formula = "=VLOOKUP(B2, Master!A:B,2,FALSE)"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    'Copy and remove non-stock items
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="=Non-Stock", Operator:=xlOr, Criteria2:="=#N/A"
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Non-Stock Items").Range("A1")
    ColHeaders = Range("A1:N1").Value
    Cells.Delete
    Rows(1).Insert
    ActiveSheet.Range("A1:N1").Value = ColHeaders
    ActiveSheet.AutoFilterMode = False

    'Copy the remainnig data
    Range("C1:N1").NumberFormat = "mmm-yy"
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Combined Forecast").Range("A1")
    Cells.Delete
End Sub
