Attribute VB_Name = "CleanUp"
Option Explicit

Sub removefilter(Optional wksht As String)
On Error Resume Next
    If wksht = "" Then    'If no arg is given clear the activesheet
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilterMode = False
        End If
    Else        'else clear the specified sheet
        If ActiveWorkbook.Worksheets(wksht).AutoFilterMode = True Then
            ActiveWorkbook.Worksheets(wksht).AutoFilterMode = False
        End If
    End If
    On Error GoTo 0
End Sub

