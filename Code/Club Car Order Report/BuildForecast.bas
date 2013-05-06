Attribute VB_Name = "BuildForecast"
Option Explicit

Sub CreateForecast()
    Dim iLength As Integer
    Dim n As Integer
    Dim i As Integer
    Dim x As Integer

    n = 4
    x = 12

    Worksheets("Combined Forecast").Select
    With Range("A:A")
        iLength = (.CurrentRegion.Rows.Count)
    End With

    Worksheets("Forecast").Select

    'Sim Number
    Range("A2").Formula = "='Combined Forecast'!A2"
    Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iLength, 1))

    'Item Number
    Range("B2").Formula = "='Combined Forecast'!B2"
    Range("B2").AutoFill Destination:=Range(Cells(2, 2), Cells(iLength, 2))

    'Description
    Range("C2").Formula = "='Combined Forecast'!C2"
    Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(iLength, 3))

    'On Hand
    Range("D2").Formula = "=VLOOKUP(A2, Gaps!A:C, 3, False)"
    Range("D2").AutoFill Destination:=Range(Cells(2, 4), Cells(iLength, 4))

    'On Reserve
    Range("E2").Formula = "=VLOOKUP(A2, Gaps!A:D, 4, False)"
    Range("E2").AutoFill Destination:=Range(Cells(2, 5), Cells(iLength, 5))

    'On Order
    Range("F2").Formula = "=VLOOKUP(A2,Gaps!A:F,6,FALSE)"
    Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(iLength, 6))

    'On Back Order
    Range("G2").Formula = "=VLOOKUP(A2, Gaps!A:E, 5, False)"
    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(iLength, 7))

    'WDC Qty On Hand
    Range("H2").Formula = "=VLOOKUP(A2, Gaps!A:AG, 33, False)"
    Range("H2").AutoFill Destination:=Range(Cells(2, 8), Cells(iLength, 8))

    'Average Unit Cost
    Range("I2").Formula = "=VLOOKUP(A2, Gaps!A:AC, 29, False)"
    Range("I2").AutoFill Destination:=Range(Cells(2, 9), Cells(iLength, 9))

    'Unit of Measure
    Range("J2").Formula = "=VLOOKUP(A2, Gaps!A:AF, 32, False)"
    Range("J2").AutoFill Destination:=Range(Cells(2, 10), Cells(iLength, 10))

    'Supplier
    Range("K2").Formula = "=VLOOKUP(A2, Gaps!A:AI, 35, False)"
    Range("K2").AutoFill Destination:=Range(Cells(2, 11), Cells(iLength, 11))

    'Month 1 Forecast
    Range("M2").Formula = "=D2-VLOOKUP(A2,'Combined Forecast'!A:P,4,FALSE)"
    Range("M2").AutoFill Destination:=Range(Cells(2, 13), Cells(iLength, 13))

    Range("N2").Select

    'Months 2 - 12
    Do While i < 11
        i = i + 1
        n = n + 1
        x = x + 1
        ActiveCell.Formula = "=" & Replace(Cells(2, x).Address, "$", "") & "-VLOOKUP(A2,'Combined Forecast'!A:P, " & n & ",FALSE)"
        ActiveCell.AutoFill Destination:=Range(Cells(2, ActiveCell.Column), Cells(iLength, ActiveCell.Column))
        ActiveCell.Offset(0, 1).Select
    Loop

    Range("A1").Select
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(i, 1).Value = "5113106375" Then
            'On Hand
            Cells(i, 4).Formula = "=CONVERT(" & Cells(i, 4).Value & "*36,""yd"",""ft"")"
            'On Reserve
            Cells(i, 5).Formula = "=CONVERT(" & Cells(i, 5).Value & "*36,""yd"",""ft"")"
            'On Order
            Cells(i, 6).Formula = "=CONVERT(" & Cells(i, 6).Value & "*36,""yd"",""ft"")"
            'Back Order
            Cells(i, 7).Formula = "=CONVERT(" & Cells(i, 7).Value & "*36,""yd"",""ft"")"
            'WDC
            Cells(i, 8).Formula = "=CONVERT(" & Cells(i, 8).Value & "*36,""yd"",""ft"")"
        End If
        If Cells(i, 1).Value = "99814198888" Then
            'On Hand
            Cells(i, 4).Value = Cells(i, 4).Value * 50
            'On Reserve
            Cells(i, 5).Value = Cells(i, 5).Value * 50
            'On Order
            Cells(i, 6).Value = Cells(i, 6).Value * 50
            'Back Order
            Cells(i, 7).Value = Cells(i, 7).Value * 50
            'WDC
            Cells(i, 8).Value = Cells(i, 8).Value * 50
        End If
    Next

    With Range("A:X")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 24)).Select
        With Selection
            .Value = .Value
        End With
    End With

    Cells.HorizontalAlignment = xlCenter
    Range(Cells(2, 3), Cells(iLength, 3)).HorizontalAlignment = xlLeft
    Range(Cells(2, 2), Cells(iLength, 2)).HorizontalAlignment = xlRight
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

Sub CreateHeaders()
    Worksheets("Forecast").Select
    Range("A1").Value = "Sims"
    Range("B1").Value = "Items"
    Range("C1").Value = "Description"
    Range("D1").Value = "On Hand"
    Range("E1").Value = "Reserve"
    Range("F1").Value = "OO"
    Range("G1").Value = "BO"
    Range("H1").Value = "WDC"
    Range("I1").Value = "Last Cost"
    Range("J1").Value = "UOM"
    Range("K1").Value = "Supplier"
    Range("L1").Value = "A/P"
    Range("M1").Formula = "='Combined Forecast'!D1"
    Range("N1").Formula = "='Combined Forecast'!E1"
    Range("O1").Formula = "='Combined Forecast'!F1"
    Range("P1").Formula = "='Combined Forecast'!G1"
    Range("Q1").Formula = "='Combined Forecast'!H1"
    Range("R1").Formula = "='Combined Forecast'!I1"
    Range("S1").Formula = "='Combined Forecast'!J1"
    Range("T1").Formula = "='Combined Forecast'!K1"
    Range("U1").Formula = "='Combined Forecast'!L1"
    Range("V1").Formula = "='Combined Forecast'!M1"
    Range("W1").Formula = "='Combined Forecast'!N1"
    Range("X1").Formula = "='Combined Forecast'!O1"
End Sub


