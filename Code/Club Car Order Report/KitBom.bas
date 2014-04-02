Attribute VB_Name = "KitBom"
Option Explicit

Sub CreateKitBOM()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim sAddr As String
    Dim i As Integer
    Dim j As Integer

    Worksheets("Kit BOM").Select
    Range("E1:P1").Value = Sheets("Combined Forecast").Range("C1:N1").Value
    Range("E1:P1").NumberFormat = "mmm-yy"
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    For j = 5 To TotalCols
        For i = 2 To TotalRows
            If Cells(i, 4).Value = "KIT" Then
                sAddr = Cells(i, j).Address(False, False)    'Address of the current KIT total
                'vlookup KIT SIM on combined forecast to get total needed for the current month
                Cells(i, j).Formula = "=IFERROR(VLOOKUP(" & Cells(i, 3).Address(False, False) & ",'Combined Forecast'!A:N," & j - 2 & ",FALSE),0)"
            Else
                'Multiply the kit total by the number of components needed per kit
                Cells(i, j).Formula = "=" & sAddr & "*" & Cells(i, 4).Address(False, False)
            End If
        Next
    Next
End Sub

Sub AddKitMaterial()
    Dim i As Long, n As Long, l As Long, x As Long
    Dim rowheaders As Variant

    Worksheets("Kit BOM").Select
    With Range("A:P")
        .CurrentRegion.Value = .CurrentRegion.Value
        .CurrentRegion.AutoFilter Field:=2, Criteria1:="=I", Operator:=xlAnd
        Range(Cells(1, 3), Cells(.CurrentRegion.Rows.Count, 16)).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Temp").Range("A1")
    End With

    Worksheets("Combined Forecast").Select
    With Range("A:A")
        Range(.Cells(2, 1), .Cells(.CurrentRegion.Rows.Count, 14)).Copy
    End With
    Worksheets("Temp").Select
    Range("A2").End(xlDown).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Columns("B:B").Delete
    rowheaders = Range("A1:M1")
    With Range("A:M")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 13)).Select
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
                TableDestination:="PTableKitParts!R3C1", _
                TableName:="PTKitParts", _
                DefaultVersion:=xlPivotTableVersion14
    End With
    Worksheets("PTableKitParts").Select
    Cells(3, 1).Select

    With ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 1))
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PTKitParts")
        For i = 2 To UBound(rowheaders, 2)
            .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(Format(rowheaders(1, i), "mmm-yy")), "Sum of " & Format(rowheaders(1, i), "mmm"), xlSum
        Next
    End With
    Rows("1:2").Delete

    With Range("A:A")
        Application.DisplayAlerts = False
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Copy
        .PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Application.DisplayAlerts = True
        Range("A1:M1").Value = rowheaders
        Rows(.CurrentRegion.Rows.Count).Delete
    End With

    Worksheets("Combined Forecast").Cells.Delete

    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Copy Destination:=Worksheets("Combined Forecast").Range("A1")
    End With

    Application.CutCopyMode = False
End Sub

Sub KitDescItemLookup()
    Worksheets("Combined Forecast").Select

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("B1").Value = "Item Number"
    Range("C1").Value = "Description"

    Range("B2").Formula = "=VLOOKUP(A2,master!B:C,2,FALSE)"
    Range("C2").Formula = "=VLOOKUP(A2, Gaps!A:B, 2, FALSE)"

    With Range("B:C")
        Range("B2").AutoFill Destination:=Range(.Cells(2, 1), .Cells(.CurrentRegion.Rows.Count, 1))
        Range("C2").AutoFill Destination:=Range(.Cells(2, 2), .Cells(.CurrentRegion.Rows.Count, 2))

        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Select
        Selection.Value = Selection.Value

        Range(Cells(1, 2), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Select
        Selection.Value = Selection.Value
    End With

    Application.CutCopyMode = False
    Worksheets("Combined Forecast").Select
End Sub
