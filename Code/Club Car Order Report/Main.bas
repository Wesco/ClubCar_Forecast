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
    Email SendTo:="ACoffey@wesco.com", Subject:="Club Car Forecast", Body:="""\\BR3615GAPS\gaps\Club Car\Order Report\Order Report " & Format(Date, "m-dd-yy") & ".xlsx"""
    Email SendTo:="ACoffey@wesco.com", Subject:="Club Car Hotsheet", Body:="""\\BR3615GAPS\gaps\Hotsheet\Club Car Hot " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.ScreenUpdating = True
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For Each s In ThisWorkbook.Sheets
        s.Select
        s.AutoFilterMode = False
        If s.Name <> "Macro" And s.Name <> "Kit BOM" And _
           s.Name <> "Bulk" And s.Name <> "Master" And s.Name <> "Info" Then
            s.Cells.Delete
        ElseIf s.Name = "Kit BOM" Then
            s.Range("E:P").Delete
        ElseIf s.Name = "Bulk" Then
            s.Range("F:ZZ").Delete
        ElseIf s.Name = "Info" Then
            s.Range("B:B").Delete
        ElseIf s.Name = "Bulk" Then
            s.Range("F:T").Delete
        End If
        s.Range("A1").Select
    Next

    Sheets("Macro").Select
    Range("C6").Select
    Application.DisplayAlerts = PrevDispAlert
End Sub
