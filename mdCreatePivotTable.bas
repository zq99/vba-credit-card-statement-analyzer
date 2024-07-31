Attribute VB_Name = "mdCreatePivotTable"
Option Explicit


Public Sub RefreshReports()
    Call ClearPivotTables
    Call CreateReports
End Sub


Public Sub CreateReports()
    
    On Error GoTo ErrorHandler
    
    Dim wsDataSheet As Worksheet
    Dim rngdataRange As Range
    Dim lngLastRow As Long
    Dim lngLastColumn As Long
    Dim wsPivotSheet As Worksheet
    Dim rngPivotLocation As Range
    Dim oState As New clsMonitorApplicationState
    
    If IsWorksheetEmpty("data") Then
        MsgBox "No data found!", vbOKOnly + vbInformation, "Data"
        GoTo Exit_here
    End If
    
    oState.CaptureState
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsDataSheet = ThisWorkbook.Sheets("data")
    Set wsPivotSheet = ThisWorkbook.Sheets("import")
    
    lngLastRow = wsDataSheet.Cells(wsDataSheet.Rows.Count, "A").End(xlUp).Row
    lngLastColumn = wsDataSheet.Cells(INT_START_OUTPUT_ROW, wsDataSheet.Columns.Count).End(xlToLeft).Column
    Set rngdataRange = wsDataSheet.Range(wsDataSheet.Cells(INT_START_OUTPUT_ROW, 1), wsDataSheet.Cells(lngLastRow, lngLastColumn))
    
    Set rngPivotLocation = wsPivotSheet.Range("B10")
    Call CreateTransactionPivotTableReport(rngdataRange, wsPivotSheet, rngPivotLocation)
    
    Set rngPivotLocation = wsPivotSheet.Range("K10")
    Call CreateByDayPivotTableReport(rngdataRange, wsPivotSheet, rngPivotLocation)
    
    
Exit_here:

    oState.RestoreState

    ClearObjects wsDataSheet, rngdataRange, wsPivotSheet, rngPivotLocation, oState

    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation, "Pivot Table Creation Error"
    GoTo Exit_here

End Sub


Public Sub CreateTransactionPivotTableReport(ByVal rngdataRange As Range, ByVal wsPivotSheet As Worksheet, ByVal rngPivotLocation As Range)
    
    Dim pvtCreator As clsPivotTableCreator
    
    On Error GoTo ErrorHandler

    Set pvtCreator = New clsPivotTableCreator
    
    ' Set properties
    pvtCreator.dataRange = rngdataRange
    pvtCreator.PivotTableLocation = rngPivotLocation
    
    ' Add row labels
    pvtCreator.AddRowLabel "Description"
    
    
    ' Add value fields with aggregation functions
    pvtCreator.AddValueField "Description", xlCount, "Count"
    pvtCreator.AddValueField "Amount", xlSum, "Amount Sum"
    pvtCreator.AddValueField "Amount", xlMin, "Min Spent"
    pvtCreator.AddValueField "Amount", xlMax, "Max Spent"
    pvtCreator.AddValueField "Amount", xlAverage, "Avg Spent"
    pvtCreator.AddValueField "Greater_than_10", xlSum, "Large Spend Count"
    pvtCreator.AddFilterField "Filename"
    pvtCreator.AddFilterField "Type"
    
    ' Create the pivot table
    pvtCreator.CreatePivotTable
    
Exit_here:

    ClearObjects pvtCreator
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation, "Pivot Table Creation Error"
    GoTo Exit_here
End Sub


Public Sub CreateByDayPivotTableReport(ByVal rngdataRange As Range, ByVal wsPivotSheet As Worksheet, ByVal rngPivotLocation As Range)
    
    Dim pvtCreator As clsPivotTableCreator
    
    On Error GoTo ErrorHandler

    Set pvtCreator = New clsPivotTableCreator
    
    ' Set properties
    pvtCreator.dataRange = rngdataRange
    pvtCreator.PivotTableLocation = rngPivotLocation
    
    ' Add row labels
    pvtCreator.AddRowLabel "Date"
    
    
    ' Add value fields with aggregation functions
    pvtCreator.AddValueField "Date", xlCount, "Count"
    pvtCreator.AddValueField "Amount", xlSum, "Amount Sum"
    pvtCreator.AddValueField "Amount", xlMin, "Min Spent"
    pvtCreator.AddValueField "Amount", xlMax, "Max Spent"
    pvtCreator.AddValueField "Amount", xlAverage, "Avg Spent"
    pvtCreator.AddValueField "Greater_than_10", xlSum, "Large Spend Count"
    pvtCreator.AddFilterField "Filename"
    pvtCreator.AddFilterField "Type"
    
    ' Create the pivot table
    pvtCreator.CreatePivotTable
    
Exit_here:

    ClearObjects pvtCreator
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation, "Pivot Table Creation Error"
    GoTo Exit_here
End Sub


Public Sub ClearPivotTables()
    
    Dim oPivot As New clsPivotTableCreator
    oPivot.ClearPivotTables ("import")
    ClearObjects oPivot
    
End Sub

