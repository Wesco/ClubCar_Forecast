Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatGaps()
    Dim iRows As Long

    Worksheets("Gaps").Select
    Columns(1).ClearContents
    Range("A1").Value = "Sim_no"
    Range("A2").Formula = "=B2 & C2"
    With Range("A:B")
        iRows = .CurrentRegion.Rows.Count
        Range("A2").AutoFill Destination:=Range(Cells(2, 1), .Cells(iRows, 1))
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 1)).Value = Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 1)).Value
    End With
    Range(Cells(1, 1), Cells(iRows, 1)).NumberFormat = "00000000000"
    Range("B:D").Delete
    With Range("A:F")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 6)).Value = Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 6)).Value
    End With
    With Range("AC:AG")
        Range(Cells(1, 29), Cells(.CurrentRegion.Rows.Count, 33)).Value = Range(Cells(1, 29), Cells(.CurrentRegion.Rows.Count, 33)).Value
    End With
    With Range("AI:AI")
        Range(Cells(1, 35), Cells(.CurrentRegion.Rows.Count, 35)).Value = Range(Cells(1, 35), Cells(.CurrentRegion.Rows.Count, 35)).Value
    End With
    With Range("CF:CF")
        Range(Cells(1, 84), Cells(.CurrentRegion.Rows.Count, 84)).Value = Range(Cells(1, 84), Cells(.CurrentRegion.Rows.Count, 84)).Value
    End With
End Sub

Sub RedBelowZero()
    Worksheets("Forecast").Select
    Columns("M:X").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub FormatHots()
    Worksheets("Hotsheet").Select
    With Range("A:A")
        Range(Cells(2, 25), Cells(.CurrentRegion.Rows.Count, 25)).NumberFormat = "[$-409]d-mmm;@"
    End With
    Columns("L:W").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    'Selection.FormatConditions(1).Font.Color = -16383844
    'Selection.FormatConditions(1).Font.TintAndShade = 0
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Select
    End With
    Selection.AutoFilter Field:=15, Criteria1:=RGB(255, 199, 206), Operator:=xlFilterCellColor
End Sub

Sub AddVisCol()

    Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("M1").Value = "Vis"

    Range("M2").Select
    Range("M2").SparklineGroups.Add Type:=xlSparkColumn, SourceData:="N2:Y2"

    Columns("M:M").ColumnWidth = 22.29

    Selection.SparklineGroups.Item(1).Points.Negative.Visible = True
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 3289650
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0

    With Range("M:M")
        Range("M2").AutoFill Destination:=Range(Cells(2, 13), Cells(.CurrentRegion.Rows.Count, 13))
    End With

    Range("M1").Select
    With Range("A:A")
        ActiveSheet.ListObjects.Add(xlSrcRange, _
                                    Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)), _
                                  , xlYes).Name = "Table1"
    End With
    Range("A2").Select

    Columns("G:G").Insert
    Range("G1").Value = "Net Stock"
    Range("G2").Formula = "=SUM(D2,F2)"
    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
    With Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
        .Value = .Value
    End With
    Columns("G:G").Delete
End Sub

Sub AddSparkLines()
    Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L1").Value = "Vis"
    Range("L2").Select
    Range("L2").SparklineGroups.Add Type:=xlSparkColumn, SourceData:="M2:X2"

    Columns("L:L").ColumnWidth = 22.29

    Selection.SparklineGroups.Item(1).Points.Negative.Visible = True
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 3289650
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0

    With Range("L:L")
        Range("L2").AutoFill Destination:=Range(Cells(2, 12), Cells(.CurrentRegion.Rows.Count, 12))
    End With

    Range("L1").Select
    With Range("A:A")
        ActiveSheet.ListObjects.Add(xlSrcRange, _
                                    Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)), _
                                  , xlYes).Name = "Table1"
    End With
    Range("A2").Select
End Sub



