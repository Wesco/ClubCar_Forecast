Attribute VB_Name = "Restructure"
Option Explicit

Sub DatesToCol()
    Dim TotalRows As Long
    
    Sheets("Temp").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    
    Range("H1").Value = "Month"
    Range("H2:H" & TotalRows).Formula = "=TEXT(E2,""mmm"")"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value
    
    Range("I1").Value = "Year"
    Range("I2:I" & TotalRows).Formula = "=TEXT(E2,""yyyy"")"
    Range("I2:I" & TotalRows).Value = Range("I2:I" & TotalRows).Value

    With ActiveWorkbook.Worksheets("Temp").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("E1:E" & TotalRows), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        .SetRange Range("A1:I" & TotalRows)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub SeparateAP()
    Dim TotalRows As Long
    
    Sheets("Temp").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    Range("A1:I" & TotalRows).AutoFilter Field:=4, Criteria1:="A"
    ActiveSheet.UsedRange.Copy Destination:=Sheets("A Whse").Range("A1")
    ActiveSheet.AutoFilter.ShowAllData

    Range("A1:I" & TotalRows).AutoFilter Field:=4, Criteria1:="P"
    ActiveSheet.UsedRange.Copy Destination:=Sheets("P Whse").Range("A1")
    ActiveSheet.AutoFilter.ShowAllData
    
    Worksheets("Temp").AutoFilterMode = False
End Sub

'Source = Data Sheet Name, PTableName = Unique Name new PTable, Destination = Destination Sheet Name
Sub CreatePTable(Source As String, PTableName As String, Destination As String)
    Dim iRows As Long
    Dim iCols As Long

    Worksheets(Source).Select
    iRows = ActiveSheet.UsedRange.Rows.Count
    iCols = ActiveSheet.UsedRange.Columns.Count

    ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=Worksheets(Source).Range(Cells(1, 1), Cells(iRows, iCols)), _
            Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=Worksheets(Destination).Range("A1"), _
            TableName:=PTableName, _
            DefaultVersion:=xlPivotTableVersion14

    Sheets(Destination).Select

    With ActiveSheet.PivotTables(PTableName)
        .PivotFields("Part").Orientation = xlRowField
        .PivotFields("Part").Position = 1
        .PivotFields("Part").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("Part").LayoutForm = xlTabular
        .PivotFields("Part Description").Orientation = xlRowField
        .PivotFields("Part Description").Position = 2
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Year").Position = 1
        .PivotFields("Year").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Month").Position = 2
        .AddDataField ActiveSheet.PivotTables(PTableName).PivotFields("Forecast Qty"), "Sum of Forecast Qty", xlSum
        .ColumnGrand = False
    End With

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:1").ClearContents
    Range("A1").Select
End Sub

Sub FormatPivTable(wksheet As String)
    Dim i As Integer
    Dim iMonth As Integer
    Dim lCols As Long
    Dim lRows As Long
    Dim sYear As String
    Dim vCell As Variant
    Dim rngYear As Range
    Dim rngMonth As Range
    Dim rngHeaders As Range

    Worksheets(wksheet).Select

    iMonth = 12
    lCols = ActiveSheet.UsedRange.Columns.Count
    lRows = ActiveSheet.UsedRange.Rows.Count
    Set rngYear = Range(Cells(2, 3), Cells(2, lCols))
    Set rngMonth = Range(Cells(3, 3), Cells(3, lCols))
    Set rngHeaders = Range(Cells(1, 3), Cells(1, lCols))

    For i = 1 To rngYear.Columns.Count
        If rngYear(1, i).Value = "Grand Total" Then
            rngHeaders(1, i).Value = "Total"
            Exit For
        End If

        If rngYear(1, i).Value <> "" Then
            sYear = rngYear(1, i).Value
        Else
            rngYear(1, i).Value = sYear
        End If
        rngHeaders(1, i).Value = rngMonth(1, i).Value & "-" & rngYear(1, i).Value
    Next

    Range("A1").Value = "Item Number"
    Range("B1").Value = "Description"
    Range("2:3").Delete Shift:=xlUp

    ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Select
    Selection.Value = 0
    Range("C1").Select

    For i = 1 To rngHeaders.Columns.Count - 1
        If Format(ActiveCell.Value, "yyyymm") < Format(Date, "yyyymm") Then
            Columns(ActiveCell.Column).Delete Shift:=xlToLeft
        Else
            Exit For
        End If
    Next

    lCols = ActiveSheet.UsedRange.Columns.Count - 1
    lRows = ActiveSheet.UsedRange.Rows.Count
    Set rngHeaders = Range(Cells(1, 3), Cells(1, lCols))

    For Each vCell In rngHeaders
        vCell.Value = Format(vCell.Value, "mmm")
    Next

    'Search column headers for the total column
    If ActiveSheet.UsedRange.Columns.Count < 15 Then
        i = 1
        Do While Cells(1, i).Value <> "Total"
            i = i + 1
        Loop

        'After finding the total column go left one to find the
        'last column containing a month
        i = i - 1
        Do While Format(iMonth, "mmm") <> Cells(1, i).Text
            iMonth = iMonth - 1
        Loop

        'After finding the number of the current month
        'add a month and insert as a new column
        'until 15 months have been reached
        Do While Cells(1, 15).Text <> "Total"
            iMonth = iMonth + 1
            i = i + 1
            Columns(i).Insert
            Cells(1, i).Value = Format(iMonth, "mmm")
            Range(Cells(2, i), Cells(lRows, i)).Value = 0
        Loop
    End If

    Do While Cells(1, 15).Text <> "Total"
        Columns(15).Delete
    Loop

    Range(Cells(2, 15), Cells(ActiveSheet.UsedRange.Rows.Count, 15)).ClearContents
    Range("O2").Formula = "=SUM(C2:N2)"
    Range("O2").AutoFill Destination:=Range(Cells(2, 15), Cells(ActiveSheet.UsedRange.Rows.Count, 15))
    Range(Cells(2, 15), Cells(ActiveSheet.UsedRange.Rows.Count, 15)).Value = _
    Range(Cells(2, 15), Cells(ActiveSheet.UsedRange.Rows.Count, 15)).Value
End Sub














