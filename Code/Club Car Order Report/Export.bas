Attribute VB_Name = "Export"
Option Explicit

Sub ExportForecast()
    Dim MacroWkbk As String: MacroWkbk = ActiveWorkbook.Name
    Dim sLoc As String: sLoc = "\\BR3615GAPS\gaps\Hotsheet\Club Car\Order Report\"
    Dim NewWkbk As String

    Workbooks(MacroWkbk).Sheets("Forecast").Copy
    Columns.AutoFit
    NewWkbk = ActiveWorkbook.Name
    AddVisCol

    On Error Resume Next

    Workbooks(MacroWkbk).Sheets("Bulk").Copy After:=Workbooks(NewWkbk).Sheets(Sheets.Count)
    Columns.AutoFit
    Workbooks(MacroWkbk).Sheets("Non-Stock Items").Copy After:=Workbooks(NewWkbk).Sheets(Sheets.Count)
    Columns.AutoFit
    Workbooks(MacroWkbk).Sheets("Info").Copy After:=Workbooks(NewWkbk).Sheets(Sheets.Count)
    Columns.AutoFit
    Err.Clear
    
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs FileName:="\\BR3615GAPS\gaps\Club Car\Order Report\Order Report " & Format(Date, "m-dd-yy") & ".xlsx"
    Select Case Err.Number
        Case Is = 0
            ActiveWorkbook.Close
        Case Else
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
    End Select
    Err.Clear

    Workbooks(MacroWkbk).Sheets("Hotsheet").Copy
    AddSparkLines

    ActiveWorkbook.BreakLink Name:="C:\Users\treische\Desktop\New folder\Club Car\Club Car Report.xlsm", Type:=xlExcelLinks
    Err.Clear

    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs FileName:="\\BR3615GAPS\gaps\Hotsheet\Club Car Hot " & Format(Date, "m-dd-yy") & ".xlsx"
    Select Case Err.Number
        Case Is = 0
            ActiveWorkbook.Close
        Case Else
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
    End Select
    On Error GoTo 0
End Sub



