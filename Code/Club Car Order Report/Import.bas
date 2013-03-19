Attribute VB_Name = "Import"
Option Explicit

Sub ImportData()
    Dim sWkbk As String: sWkbk = ThisWorkbook.Name
    Dim sGapsLoc As String: sGapsLoc = "\\br3615gaps\GAPS\3615 Gaps Download\" & Format(Date, "yyyy") & "\"
    Dim sGaps As String: sGaps = "3615 "
    Dim sFcstLoc As String: sFcstLoc = "\\br3615gaps\GAPS\Club Car\Forecast\" & Format(Date, "yyyy") & "\"
    Dim sWhseP As String: sWhseP = "Warehouse P forecast "
    Dim sWhseA As String: sWhseA = "Warehouse A forecast "
    Dim sDate As String: sDate = Format(Date, "m-dd-yy")
    Dim rngAFcst As Range: Set rngAFcst = Worksheets("A Forecast").Range("A1")
    Dim rngPFcst As Range: Set rngPFcst = Worksheets("P Forecast").Range("A1")
    Dim rngGaps As Range: Set rngGaps = Worksheets("Gaps").Range("A1")
    Dim i As Integer: i = 0

    'Find the most up to date forecast
    Do While FileOrDirExists(sFcstLoc & sWhseP & sDate & ".xlsx") = False
        i = i + 1
        sDate = Format(Date - i, "m-dd-yy")
    Loop
    
    Worksheets("Info").Range("B3").Value = sDate
    'Store most up to date forecast filenames
    sWhseP = sWhseP & sDate & ".xlsx"
    sWhseA = sWhseA & sDate & ".xlsx"
    'Reset loop variables
    sDate = Format(Date, "m-dd-yy")
    i = 0

    'Find the most up to date gaps file
    Do While FileOrDirExists(sGapsLoc & sGaps & sDate & ".xlsx") = False
        i = i + 1
        sDate = Format(Date - i, "m-dd-yy")
    Loop
    
    Worksheets("Info").Range("B2").Value = sDate
    
    'Store most up to date gaps filename
    sGaps = sGaps & sDate & ".xlsx"

    'Import Warehouse A Forecast
    Workbooks.Open (sFcstLoc & sWhseA)
    'Range("A:O").Copy Destination:=rngAFcst ' code not needed - remove this
    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 15)).Copy Destination:=rngAFcst
    End With
    Workbooks(sWhseA).Close

    'Import Warehosue P Forecast
    Workbooks.Open (sFcstLoc & sWhseP)
    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 15)).Copy Destination:=rngPFcst
    End With
    Workbooks(sWhseP).Close

    'Import Gaps
    Workbooks.Open (sGapsLoc & sGaps)
    With Range("A:A")
        Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, 98)).Copy Destination:=rngGaps
    End With
    Workbooks(sGaps).Close

End Sub

Function FileOrDirExists(File As String) As Boolean
    Dim iTemp As Integer

    'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(File)

    'Check if error exists and set response appropriately
    Select Case Err.Number
        Case Is = 0
            FileOrDirExists = True
        Case Else
            FileOrDirExists = False
    End Select

    'Resume error checking
    On Error GoTo 0
End Function

