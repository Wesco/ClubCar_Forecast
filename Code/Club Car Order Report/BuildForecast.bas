Attribute VB_Name = "BuildForecast"
Option Explicit

Sub CreateForecast()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long
    Dim j As Integer

    Sheets("Combined Forecast").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Get SIM Numbers
    Range("A2:A" & TotalRows).Copy Destination:=Sheets("Forecast").Range("A2")

    'Get Item Numbers
    Range("B2:B" & TotalRows).Copy Destination:=Sheets("Forecast").Range("B2")

    'Get Descriptions
    Range("C2:C" & TotalRows).Copy Destination:=Sheets("Forecast").Range("C2")

    Sheets("Forecast").Select

    'Add column headers
    Range("A1:L1").Value = Array("Sims", "Items", "Description", "On Hand", "Reserve", "OO", "BO", "WDC", "Last Cost", "UOM", "Supplier", "A/P")
    Range("M1:X1").Formula = "='Combined Forecast'!D1"
    Range("M1:X1").Value = Range("M1:X1").Value

    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'On Hand
    Range("D2:D" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:C, 3, False),0)"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value

    'On Reserve
    Range("E2:E" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:D, 4, False),0)"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    'On Order
    Range("F2:F" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:F,6,FALSE),0)"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    'On Back Order
    Range("G2:G" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:E, 5, False),0)"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value

    'WDC Qty On Hand
    Range("H2:H" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:AG, 33, False),0)"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value

    'Average Unit Cost
    Range("I2:I" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:AC, 29, False),0)"
    Range("I2:I" & TotalRows).Value = Range("I2:I" & TotalRows).Value

    'Unit of Measure
    Range("J2:J" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:AF, 32, False),0)"
    Range("J2:J" & TotalRows).Value = Range("J2:J" & TotalRows).Value

    'Supplier
    Range("K2:K" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Gaps!A:AI, 35, False),"""")"
    Range("K2:K" & TotalRows).Value = Range("K2:K" & TotalRows).Value

    'Month 1 Forecast
    Range("M2:M" & TotalRows).Formula = "=D2-VLOOKUP(A2,'Combined Forecast'!A:P,4,FALSE)"

    'Months 2 - 12 forecast
    For i = 14 To TotalCols
        Range(Cells(2, i), Cells(TotalRows, i)).Formula = "=" & Cells(2, i - 1).Address(False, False) & "-VLOOKUP(A2,'Combined Forecast'!A:P," & i - 9 & ",FALSE)"
    Next

    'UOM conversions
    For i = 2 To TotalRows
        If Cells(i, 1).Value = "5113106375" Then
            For j = 4 To 8
                Cells(i, j).Formula = "=CONVERT(" & Cells(i, j).Value & "*36,""yd"",""ft"")"
            Next
        ElseIf Cells(i, 1).Value = "99814198888" Then
            For j = 4 To 8
                Cells(i, j).Value = Cells(i, j).Value * 50
            Next
        End If
    Next


    With Range("A:X")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 24)).Select
        With Range(Cells(1, 1), Cells(TotalRows, TotalCols))
            .Value = .Value
            .HorizontalAlignment = xlCenter
        End With
    End With

    Range(Cells(2, 3), Cells(TotalRows, 3)).HorizontalAlignment = xlLeft
    Range(Cells(2, 2), Cells(TotalRows, 2)).HorizontalAlignment = xlRight
End Sub

Sub FillAP()
    Worksheets("Forecast").Select
    Dim ares As Variant, pres As Variant, bres As Variant, kres As Variant, cValue As Variant, simValue As Variant
    Dim sA As String, sP As String, sB As String, sK As String
    Dim i As Integer: i = 1

    Range("L2").Select
    With Range("A:A")
        Do While i < .CurrentRegion.Rows.Count
            i = i + 1
            cValue = ActiveCell.Offset(0, -10).Value
            simValue = ActiveCell.Offset(0, -11).Value

            ares = Application.VLookup(cValue, Worksheets("A Forecast").Range("A:A"), 1, False)
            If IsError(ares) = False Then
                sA = "A"
            Else
                Set ares = Nothing
                sA = vbNullString
            End If

            pres = Application.VLookup(cValue, Worksheets("P Forecast").Range("A:A"), 1, False)
            If IsError(pres) = False Then
                sP = "P"
            Else
                Set pres = Nothing
                sP = vbNullString
            End If

            bres = Application.VLookup(simValue, Worksheets("Bulk").Range("B:B"), 1, False)
            If IsError(bres) = False Then
                sB = "B"
            Else
                Set pres = Nothing
                sB = vbNullString
            End If

            kres = Application.VLookup(simValue, Worksheets("Kit BOM").Range("C:C"), 1, False)
            If IsError(kres) = False Then
                sK = "K"
            Else
                Set pres = Nothing
                sK = vbNullString
            End If

            ActiveCell.Value = sA & sP & sB & sK
            ActiveCell.Offset(1, 0).Select
        Loop
    End With
End Sub

Sub CreateBulk()
    Dim rowheaders() As Variant
    Dim rowdata() As Variant
    Dim TotalRows As Long
    Dim TotalCols As Long

    rowheaders = Array( _
                 "Type", "Sim", "Desc", "Supp", "Notes", "OH", "RES", "BO", "OO", "Last Cost", _
                 "='Combined Forecast'!D1", _
                 "='Combined Forecast'!E1", _
                 "='Combined Forecast'!F1", _
                 "='Combined Forecast'!G1", _
                 "='Combined Forecast'!H1", _
                 "=""End "" & 'Combined Forecast'!D1", _
                 "=""End "" & 'Combined Forecast'!E1", _
                 "=""End "" & 'Combined Forecast'!F1", _
                 "=""End "" & 'Combined Forecast'!G1", _
                 "=""End "" & 'Combined Forecast'!H1")


    rowdata = Array( _
              "=IFERROR(VLOOKUP(B2, Gaps!A:C, 3, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, Gaps!A:D, 4, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, Gaps!A:E, 5, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, Gaps!A:F, 6, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, Gaps!A:C, 29, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, 'Combined Forecast'!A:O, 4, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, 'Combined Forecast'!A:O, 5, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, 'Combined Forecast'!A:O, 6, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, 'Combined Forecast'!A:O, 7, FALSE), 0)", _
              "=IFERROR(VLOOKUP(B2, 'Combined Forecast'!A:O, 8, FALSE), 0)", _
              "=F2-K2", "=P2-L2", "=Q2-M2", "=R2-N2", "=S2-O2")



    Worksheets("Bulk").Select

    Range("A1:T1") = rowheaders
    Range("F2:T2") = rowdata

    With Range("F:L")
        Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(.CurrentRegion.Rows.Count, 6))
        Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(.CurrentRegion.Rows.Count, 7))
        Range("H2").AutoFill Destination:=Range(Cells(2, 8), Cells(.CurrentRegion.Rows.Count, 8))
        Range("I2").AutoFill Destination:=Range(Cells(2, 9), Cells(.CurrentRegion.Rows.Count, 9))
        Range("J2").AutoFill Destination:=Range(Cells(2, 10), Cells(.CurrentRegion.Rows.Count, 10))
        Range("K2").AutoFill Destination:=Range(Cells(2, 11), Cells(.CurrentRegion.Rows.Count, 11))
        Range("L2").AutoFill Destination:=Range(Cells(2, 12), Cells(.CurrentRegion.Rows.Count, 12))
        Range("M2").AutoFill Destination:=Range(Cells(2, 13), Cells(.CurrentRegion.Rows.Count, 13))
        Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(.CurrentRegion.Rows.Count, 14))
        Range("O2").AutoFill Destination:=Range(Cells(2, 15), Cells(.CurrentRegion.Rows.Count, 15))
        Range("P2").AutoFill Destination:=Range(Cells(2, 16), Cells(.CurrentRegion.Rows.Count, 16))
        Range("Q2").AutoFill Destination:=Range(Cells(2, 17), Cells(.CurrentRegion.Rows.Count, 17))
        Range("R2").AutoFill Destination:=Range(Cells(2, 18), Cells(.CurrentRegion.Rows.Count, 18))
        Range("S2").AutoFill Destination:=Range(Cells(2, 19), Cells(.CurrentRegion.Rows.Count, 19))
        Range("T2").AutoFill Destination:=Range(Cells(2, 20), Cells(.CurrentRegion.Rows.Count, 20))
    End With

    With Range("A:T")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Select
    End With

    With Selection
        .Value = .Value
    End With

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="J"

    With Range("A:T").SpecialCells(xlCellTypeVisible)
        .SpecialCells(xlCellTypeConstants).Font.Bold = True
    End With

    ActiveSheet.ShowAllData
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="I"

    With Range("A:T").SpecialCells(xlCellTypeVisible)
        .SpecialCells(xlCellTypeConstants).Font.Bold = False
    End With

    ActiveSheet.ShowAllData
    ActiveSheet.UsedRange.AutoFilter Field:=2, Criteria1:=RGB(204, 255, 204), Operator:=xlFilterCellColor

    With Range("A:T").SpecialCells(xlCellTypeVisible)
        .SpecialCells(xlCellTypeConstants).Interior.Color = 13434828
    End With

    ActiveSheet.ShowAllData
    ActiveSheet.Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter Field:=3, _
                                                                           Criteria1:=RGB(255, 255, 153), _
                                                                           Operator:=xlFilterCellColor

    With Range("A:T").SpecialCells(xlCellTypeVisible)
        .SpecialCells(xlCellTypeConstants).Interior.Color = 10092543
    End With

    ActiveSheet.ShowAllData
    Selection.AutoFilter

    Range("A1:T1").Font.Bold = True
    Range("A1:T1").HorizontalAlignment = xlCenter

    With Range("F:T")
        Range(.Cells(2, 1), .Cells(.CurrentRegion.Rows.Count, 11)).HorizontalAlignment = xlCenter
    End With

    With Range("A1:T1").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
