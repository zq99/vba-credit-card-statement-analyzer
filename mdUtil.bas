Attribute VB_Name = "mdUtil"
Option Explicit


Public Sub ClearObjects(ParamArray objList() As Variant)
    Dim i As Integer

    For i = LBound(objList) To UBound(objList())
        DoEvents
        If VarType(objList(i)) = vbObject Then
            Set objList(i) = Nothing
        End If
    Next i
    
End Sub


Public Function IsWorksheetEmpty(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    Set rng = ws.UsedRange
    IsWorksheetEmpty = Application.WorksheetFunction.CountA(rng) = 0
    
    ClearObjects ws, rng
    
End Function
