Attribute VB_Name = "Main"
Option Explicit

Sub OrderReport()
    Application.ScreenUpdating = False
    
    Dim starttime As Double: starttime = Timer
    Dim MacroWkbk As String: MacroWkbk = ActiveWorkbook.Name
    
    Worksheets("Info").Range("B4").Value = Format(Date, "m/d/yyyy")
    Worksheets("Info").Range("B5").Value = Environ("USERNAME")

    ImportData
    FormatGaps
    
    PTableAP
    FilterNS
    CreateKitBOM
    AddKitMaterial
    KitDescItemLookup
    
    CreateHeaders
    CreateForecast
    FillAP
    RedBelowZero
    CreateBulk
    RemoveNonStock
    Hotsheet
    FormatHots

    Workbooks(MacroWkbk).Worksheets("Info").Range("B1").Value = Timer - starttime
    ExportForecast

    Worksheets("Macro").Select
    MsgBox ("Complete!")
    'Email SendTo:="JBarnhill@wesco.com", Subject:="Club Car Forecast", Body:="""\\BR3615GAPS\gaps\Club Car\Order Report\Order Report " & Format(Date, "m-dd-yy") & ".xlsx"""
    'Email SendTo:="ACoffey@wesco.com", Subject:="Club Car Hotsheet", Body:="""\\BR3615GAPS\gaps\Hotsheet\Club Car Hot " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.ScreenUpdating = True
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim s As Worksheet
    
    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" And s.Name <> "Kit BOM" And s.Name <> "Bulk" And s.Name <> "Master" Then
            s.Select
            s.Cells.Delete
            s.Range("A1").Select
        ElseIf s.Name = "Kit BOM" Then
            s.Select
            s.Range("E:ZZ").Delete
            s.Range("A1").Select
        ElseIf s.Name = "Bulk" Then
            s.Select
            s.Range("F:ZZ").Delete
            s.Range("A1").Select
        End If
    Next
    
    Sheets("Macro").Select
    Range("C6").Select
    Application.DisplayAlerts = PrevDispAlert
End Sub
