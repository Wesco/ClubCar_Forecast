Attribute VB_Name = "CreateHots"
Option Explicit

Sub Hotsheet()
    Dim sLoc As String: sLoc = "\\BR3615GAPS\gaps\Hotsheet\"
    Dim iCounter As Integer
    Dim TotalRows As Long
    
    For iCounter = 0 To 15
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

            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
            Exit For
        End If
    Next

    Worksheets("Forecast").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("A1:H" & TotalRows & ",J1:X" & TotalRows).Copy Destination:=Worksheets("Hotsheet").Range("A1")
    
    Worksheets("Hotsheet").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    
    Range("X1").Value = "Notes"
    Range("X2:X" & TotalRows).Formula = "=IFERROR(IF(IFERROR(VLOOKUP(A2, Temp!A:Y,25,FALSE), """") = 0, """", VLOOKUP(A2, Temp!A:Y,25,FALSE)), """")"
    Range("X2:X" & TotalRows).Value = Range("X2:X" & TotalRows).Value
    
    Range("Y1").Value = "Note Date"
    Range("Y2:Y" & TotalRows).Formula = "=IFERROR(IF(IFERROR(VLOOKUP(A2, Temp!A:Z,26,FALSE), """") = 0, """", VLOOKUP(A2, Temp!A:Z,26,FALSE)), """")"
    Range("Y2:Y" & TotalRows).Value = Range("Y2:Y" & TotalRows).Value
End Sub

