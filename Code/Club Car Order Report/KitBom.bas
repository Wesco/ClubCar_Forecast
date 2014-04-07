Attribute VB_Name = "KitBom"
Option Explicit

Sub CreateKitBOM()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim sAddr As String
    Dim i As Integer
    Dim j As Integer

    Sheets("Kit BOM").Select
    Range("E1:P1").Value = Sheets("Combined Forecast").Range("C1:N1").Value
    Range("E1:P1").NumberFormat = "d-mmm-yy"
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

    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
End Sub

Sub AddKitMaterial()
    Dim PrevDispAlert As Boolean
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long

    PrevDispAlert = Application.DisplayAlerts

    'Copy the kit components forecast
    Sheets("Kit BOM").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    ActiveSheet.UsedRange.AutoFilter Field:=2, Criteria1:="=I", Operator:=xlAnd
    Range("C1:P" & TotalRows).Copy Destination:=Sheets("Temp").Range("A1")

    'Copy the forecast data to combine it with the kit components
    Worksheets("Combined Forecast").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("A2:N" & TotalRows).Copy Destination:=Sheets("Temp").Range("A" & Sheets("Temp").UsedRange.Rows.Count + 1)

    Worksheets("Temp").Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove column with mixed data (contains qty per kit / part numbers)
    Columns("B:B").Delete
    ColHeaders = Range("A1:M1")

    'Sort by SIM number
    With ActiveWorkbook.Worksheets("Temp").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange Range("A1:M" & TotalRows)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Create a pivot table out of the combined forecasts
    ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=Sheets("Temp").UsedRange, _
            Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=Sheets("PTableKitParts").Range("A1"), _
            TableName:="PTKitParts", _
            DefaultVersion:=xlPivotTableVersion14

    'Add fields to the pivot table
    Worksheets("PTableKitParts").Select
    With ActiveSheet.PivotTables("PTKitParts")
        .PivotFields(ColHeaders(1, 1)).Orientation = xlRowField
        .PivotFields(ColHeaders(1, 1)).Position = 1

        For i = 2 To UBound(ColHeaders, 2)
            .AddDataField .PivotFields(Format(ColHeaders(1, i), "d-mmm-yy")), "Sum of " & Format(ColHeaders(1, i), "d-mmm-yy"), xlSum
        Next
    End With

    'Store the pivot table as values
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Application.DisplayAlerts = False
    Range("A1:M" & TotalRows).Copy
    Range("A1:M" & TotalRows).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Application.DisplayAlerts = PrevDispAlert

    'Set the column headers
    Range("A1:M1").Value = ColHeaders
    Range("A1:M1").NumberFormat = "d-mmm-yy"

    'Remove grand total
    Rows(TotalRows).Delete
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    Sheets("Combined Forecast").Cells.Delete
    Range("A1:M" & TotalRows).Copy Destination:=Sheets("Combined Forecast").Range("A1")
End Sub

Sub KitDescItemLookup()
    Dim TotalRows As Long

    Sheets("Combined Forecast").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    Columns("B:C").Insert

    'Lookukp item numbers
    Range("B1").Value = "Item Number"
    Range("B2:B" & TotalRows).Formula = "=VLOOKUP(A2,master!B:C,2,FALSE)"
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value

    'Lookup part descriptions
    Range("C1").Value = "Description"
    Range("C2:C" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:B,2,FALSE),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value
End Sub
