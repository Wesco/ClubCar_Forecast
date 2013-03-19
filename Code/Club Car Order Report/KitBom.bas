Attribute VB_Name = "KitBom"
Option Explicit

Sub CreateKitBOM()
    Dim rowheaders As Variant
    rowheaders = Array("='Combined Forecast'!C1", _
                       "='Combined Forecast'!D1", _
                       "='Combined Forecast'!E1", _
                       "='Combined Forecast'!F1", _
                       "='Combined Forecast'!G1", _
                       "='Combined Forecast'!H1", _
                       "='Combined Forecast'!I1", _
                       "='Combined Forecast'!J1", _
                       "='Combined Forecast'!K1", _
                       "='Combined Forecast'!L1", _
                       "='Combined Forecast'!M1", _
                       "='Combined Forecast'!N1")
    Dim sAddr As String
    Dim i As Integer: i = 1
    Dim n As Integer: n = 1
    Dim x As Integer: x = 4
    Worksheets("Kit BOM").Select
    Range("E1:P1").Value = rowheaders
    Range("D2").Select
    Do While i < 13
        If ActiveCell.Text = "KIT" Then
            sAddr = Replace(ActiveCell.Offset(0, n).Address, "$", "")
            ActiveCell.Offset(0, i).Formula = "=IFERROR(VLOOKUP(" & Replace(ActiveCell.Offset(0, -1).Address, "$", "") & ",'Combined Forecast'!A:O," & x & ",FALSE),0)"
        Else
            ActiveCell.Offset(0, i).Formula = "=" & sAddr & "*" & Replace(ActiveCell.Address, "$", "")
        End If
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Text = "" Then
            i = i + 1
            n = n + 1
            x = x + 1
            Range("D2").Select
        End If
    Loop
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
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 2)), "Sum of Aug", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 3)), "Sum of Sep", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 4)), "Sum of Oct", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 5)), "Sum of Nov", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 6)), "Sum of Dec", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 7)), "Sum of Jan", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 8)), "Sum of Feb", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 9)), "Sum of Mar", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 10)), "Sum of Apr", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 11)), "Sum of May", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 12)), "Sum of Jun", xlSum
        .AddDataField ActiveSheet.PivotTables("PTKitParts").PivotFields(rowheaders(1, 13)), "Sum of Jul", xlSum
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


