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
    Dim TotalRows As Long
    Dim TotalCols As Long

    Sheets(Source).Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=Sheets(Source).Range(Cells(1, 1), Cells(TotalRows, TotalCols)), _
            Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=Sheets(Destination).Range("A1"), _
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

    Rows("1:1").Delete
    Range("A1").Select
End Sub

Sub FormatPivTable(wksheet As String)
    Dim TotalCols As Long
    Dim TotalRows As Long
    Dim sYear As String
    Dim i As Integer

    Sheets(wksheet).Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Add new column headers
    For i = 3 To TotalCols - 1
        If Cells(1, i).Value <> "" Then
            sYear = Cells(1, i).Value
        End If
        Cells(1, i).Value = Cells(2, i).Value & "-" & sYear
    Next
    
    'Remove the "Total" column
    Columns(TotalCols).Delete
    
    'Add column headers
    Range("A1").Value = "Item Number"
    Range("B1").Value = "Description"
    
    'Remove old column headers
    Rows(2).Delete Shift:=xlUp
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    On Error Resume Next
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Value = 0
    On Error GoTo 0

    'If the month has passed remove it from the list
    For i = TotalCols - 1 To 3 Step -1
        If Format(Cells(1, i).Value, "yyyymm") < Format(Date, "yyyymm") Then
            Columns(i).Delete Shift:=xlToLeft
        End If
    Next

    'Remove the "Total" column
    Columns(ActiveSheet.UsedRange.Columns.Count).Delete
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    
    'If there are less than 15 columns then add
    'months and fill the needed qty with 0
    If TotalCols < 15 Then
        For i = TotalCols + 1 To 14
            Cells(1, i).Value = DateAdd("m", 1, Cells(1, i - 1).Value)
            Cells(1, i).NumberFormat = "mmm-yy"
            Range(Cells(2, i), Cells(TotalRows, i)).Value = 0
        Next
    Else
        For i = TotalCols To 15 Step -1
            Columns(i).Delete
        Next
    End If
    
    'Add Totals
    Range("O1").Value = "Total"
    Range("O2:O" & TotalRows).Formula = "=SUM(C2:N2)"
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub
