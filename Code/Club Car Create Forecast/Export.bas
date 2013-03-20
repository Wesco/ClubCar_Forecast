Attribute VB_Name = "Export"
Option Explicit

Sub SaveForecast()
    Dim sPath As String
    sPath = "\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"

    If FileExists(sPath) <> True Then
        MkDir sPath
    End If
    
    Worksheets("PivotTableA").Copy
    On Error Resume Next
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs sPath & "Warehouse A forecast " & Format(Date, "m-dd-yy") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    
    Worksheets("PivotTableP").Copy
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs sPath & "Warehouse P forecast " & Format(Date, "m-dd-yy") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    On Error GoTo 0
    
End Sub
