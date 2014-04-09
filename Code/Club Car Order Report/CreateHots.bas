Attribute VB_Name = "CreateHots"
Option Explicit

Sub Hotsheet()
    Dim sLoc As String: sLoc = "\\BR3615GAPS\gaps\Hotsheet\"
    Dim iCounter As Integer: iCounter = 0
    Dim TotalRows As Long
    
    For iCounter = 0 To 14
        If FileOrDirExists(sLoc & "Club Car Hot " & Format(Date - iCounter, "m-dd-yy") & ".xlsx") = True Then
            sLoc = sLoc & "Club Car Hot " & Format(Date - iCounter, "m-dd-yy") & ".xlsx"
            removefilter ("Temp")
            Worksheets("Temp").Cells.Delete
            Workbooks.Open (sLoc)
            
            On Error Resume Next
            ActiveSheet.ShowAllData
            removefilter
            On Error GoTo 0
            
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Temp").Range("A1")

            Application.CutCopyMode = False
            Application.DisplayAlerts = False
            Workbooks("Club Car Hot " & Format(Date - iCounter, "m-dd-yy") & ".xlsx").Close
            Application.DisplayAlerts = True
            Exit For
        End If
    Next

    Worksheets("Forecast").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("A1:H" & TotalRows & ",J1:X" & TotalRows).Copy Destination:=Worksheets("Hotsheet").Range("A1")
    Worksheets("Hotsheet").Select
    Range("X1").Value = "Notes"
    Range("Y1").Value = "Note Date"
    Range("X2").Formula = "=IFERROR(IF(IFERROR(VLOOKUP(A2, Temp!A:Y,25,FALSE), """") = 0, """", VLOOKUP(A2, Temp!A:Y,25,FALSE)), """")"
    Range("Y2").Formula = "=IFERROR(IF(IFERROR(VLOOKUP(A2, Temp!A:Z,26,FALSE), """") = 0, """", VLOOKUP(A2, Temp!A:Z,26,FALSE)), """")"

    With Range("A:A")
        Range("X2").AutoFill Destination:=Range(Cells(2, 24), Cells(.CurrentRegion.Rows.Count, 24))
        Range("Y2").AutoFill Destination:=Range(Cells(2, 25), Cells(.CurrentRegion.Rows.Count, 25))
        Range(Cells(1, 24), Cells(.CurrentRegion.Rows.Count, 24)).Select
        Selection.Value = Selection.Value
        Range(Cells(1, 25), Cells(.CurrentRegion.Rows.Count, 25)).Select
        Selection.Value = Selection.Value
    End With
    
'    Columns("G:G").Insert
'    Range("G1").Value = "Net Stock"
'    Range("G2").Formula = "=SUM(D2,F2)"
'    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
'    With Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
'        .Value = .Value
'    End With
'
'    Columns("G:G").Delete
    
    Application.CutCopyMode = False
End Sub

