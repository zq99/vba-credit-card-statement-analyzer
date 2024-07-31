VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCSVImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsCSVImporter

Option Explicit

Private mblnScreenUpdating As Boolean
Private mintCalc As Integer
Private mblnDisplayAlerts As Boolean


Public Sub CaptureState()
    mblnScreenUpdating = Application.ScreenUpdating
    mintCalc = Application.Calculation
    mblnDisplayAlerts = Application.DisplayAlerts
End Sub

Public Sub RestoreState()
    Application.ScreenUpdating = mblnScreenUpdating
    Application.Calculation = mintCalc
    Application.DisplayAlerts = mblnDisplayAlerts
End Sub


Private Sub ClearObjects(ParamArray objList() As Variant)
    Dim i As Integer

    For i = LBound(objList) To UBound(objList())
        DoEvents
        If VarType(objList(i)) = vbObject Then
            Set objList(i) = Nothing
        ElseIf VarType(objList(i)) = vbBoolean Then
            objList(i) = False
        End If
    Next i
    
End Sub


Public Function ImportCSVtoSheet(ByVal csvFilePath As String, ByVal outputSheetName As String, ByVal startRow As Long, Optional ByVal excludeHeaderRow As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tempWorkbook As Workbook
    Dim outputWorksheet As Worksheet
    Dim tempWorksheet As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim csvDataRange As Range
    Dim firstDataRow As Long
    Dim cell As Range

    Call CaptureState

    ClearObjects Application.ScreenUpdating, Application.DisplayAlerts

    Set tempWorkbook = Workbooks.Open(Filename:=csvFilePath, ReadOnly:=True)
    Set tempWorksheet = tempWorkbook.Sheets(1)
    
    With tempWorksheet
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
        firstDataRow = IIf(excludeHeaderRow, 2, 1) ' Start from row 2 if header should be excluded
        Set csvDataRange = .Range(.Cells(firstDataRow, 1), .Cells(lastRow, lastColumn))
    End With
    
    Set outputWorksheet = ThisWorkbook.Sheets(outputSheetName)
  
    With outputWorksheet
        .Range(.Cells(startRow, 1), .Cells(startRow + csvDataRange.Rows.Count, 2)).NumberFormat = "@"
        .Cells(startRow, 1).Resize(csvDataRange.Rows.Count, csvDataRange.Columns.Count).value = csvDataRange.value
        
        For Each cell In .Range(.Cells(startRow, 1), .Cells(startRow + csvDataRange.Rows.Count, 2))
            If IsDate(cell.value) Then
                cell.value = Format(cell.value, "YYYY-MM-DD")
            End If
        Next cell
        
    End With
    
    tempWorkbook.Close False
    
    ImportCSVtoSheet = True

    GoTo ExitHere

ErrorHandler:

    If Not tempWorkbook Is Nothing Then tempWorkbook.Close False
    
    ImportCSVtoSheet = False

ExitHere:
    
    ClearObjects csvDataRange, tempWorksheet, tempWorkbook, outputWorksheet, cell

    Call RestoreState

End Function


