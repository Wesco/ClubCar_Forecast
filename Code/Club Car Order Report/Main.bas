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
    RemoveOldData

    Worksheets("Macro").Select
    MsgBox ("Complete!")
    Email SendTo:="JBarnhill@wesco.com", Subject:="Club Car Forecast", Body:="""\\BR3615GAPS\gaps\Club Car\Order Report\Order Report " & Format(Date, "m-dd-yy") & ".xlsx"""
    'ActiveWorkbook.Save
    Application.ScreenUpdating = True
End Sub

