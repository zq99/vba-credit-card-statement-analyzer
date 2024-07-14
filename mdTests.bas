Attribute VB_Name = "mdTests"
' Test Module for all clsCSVImporter, clsMonitorApplicationState, clsPivotTableCreator

Option Explicit


Public Sub RunAllTests()
    Call Test_clsPivotTableCreator
    Debug.Print "All tests completed."
    Call Test_clsCSVImporter
    Debug.Print "All tests completed."
    Call Test_clsMonitorApplicationState
    Debug.Print "All tests completed."
End Sub


Private Sub Test_clsPivotTableCreator()
    Dim ws As Worksheet
    Dim testData As Range
    Dim pivotLocation As Range
    Dim ptCreator As clsPivotTableCreator
    Dim pivotTable As pivotTable
    Dim testPassed As Boolean
    
    ' Setup test environment
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TestSheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "TestSheet"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add test data
    ws.Cells(1, 1).value = "Category"
    ws.Cells(1, 2).value = "Amount"
    ws.Cells(2, 1).value = "A"
    ws.Cells(2, 2).value = 10
    ws.Cells(3, 1).value = "B"
    ws.Cells(3, 2).value = 20
    ws.Cells(4, 1).value = "A"
    ws.Cells(4, 2).value = 30
    ws.Cells(5, 1).value = "B"
    ws.Cells(5, 2).value = 40
    
    Set testData = ws.Range("A1:B5")
    Set pivotLocation = ws.Range("D1")
    
    ' Initialize the class
    Set ptCreator = New clsPivotTableCreator
    ptCreator.dataRange = testData
    ptCreator.PivotTableLocation = pivotLocation
    
    ' Add Row Labels, Values, and Filters
    ptCreator.AddRowLabel "Category"
    ptCreator.AddValueField "Amount", xlSum, "Total Amount"
    
    ' Create Pivot Table
    ptCreator.CreatePivotTable
    
    ' Verify Pivot Table Creation
    Set pivotTable = Nothing
    On Error Resume Next
    Set pivotTable = ws.PivotTables(1)
    On Error GoTo 0
    
    ' Assert that the pivot table was created
    Debug.Assert Not pivotTable Is Nothing
    
    ' Additional assertions to verify the pivot table structure
    If Not pivotTable Is Nothing Then
        Debug.Assert pivotTable.PivotFields("Category").Orientation = xlRowField
        Debug.Assert pivotTable.PivotFields("Total Amount").Function = xlSum
    End If
    
    ' Cleanup
    On Error Resume Next
    ws.Cells.Clear
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Set ws = Nothing
    Set testData = Nothing
    Set pivotLocation = Nothing
    Set ptCreator = Nothing
    Set pivotTable = Nothing
    On Error GoTo 0
End Sub


Private Sub Test_clsCSVImporter()
    Dim importer As clsCSVImporter
    Dim ws As Worksheet
    Dim csvFilePath As String
    Dim outputSheetName As String
    Dim startRow As Long
    Dim excludeHeaderRow As Boolean
    Dim importSuccess As Boolean
    Dim expectedData As Variant
    Dim importedData As Variant
    
    ' Setup test environment
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TestCSVImport")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "TestCSVImport"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Prepare test CSV file
    csvFilePath = ThisWorkbook.Path & "\test_data.csv"
    CreateTestCSVFile csvFilePath
    
    ' Set parameters for ImportCSVtoSheet
    outputSheetName = "TestCSVImport"
    startRow = 1
    excludeHeaderRow = False
    
    ' Initialize the class
    Set importer = New clsCSVImporter
    
    ' Call ImportCSVtoSheet
    importSuccess = importer.ImportCSVtoSheet(csvFilePath, outputSheetName, startRow, excludeHeaderRow)
    
    ' Assert the import was successful
    Debug.Assert importSuccess
    
    ' Define expected data (should match the contents of test_data.csv)
    expectedData = Array(Array("Name", "Age", "Country"), _
                         Array("John", "30", "USA"), _
                         Array("Jane", "25", "Canada"), _
                         Array("Tom", "40", "UK"))
    
    ' Get imported data
    importedData = ws.Range("A1:C4").value
    
    ' Assert the imported data matches the expected data
    Dim i As Long, j As Long
    For i = LBound(expectedData) To UBound(expectedData)
        For j = LBound(expectedData(i)) To UBound(expectedData(i))
            Debug.Assert CStr(importedData(i + 1, j + 1)) = CStr(expectedData(i)(j))
        Next j
    Next i
    
    ' Cleanup
    On Error Resume Next
    ws.Cells.Clear
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Kill csvFilePath
    Set ws = Nothing
    Set importer = Nothing
    On Error GoTo 0
End Sub


Private Sub CreateTestCSVFile(filePath As String)
    Dim fileNum As Integer
    Dim csvContent As String
    
    csvContent = "Name,Age,Country" & vbCrLf & _
                 "John,30,USA" & vbCrLf & _
                 "Jane,25,Canada" & vbCrLf & _
                 "Tom,40,UK"
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent
    Close #fileNum
End Sub


Private Sub Test_clsMonitorApplicationState()
    Dim monitor As clsMonitorApplicationState
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalCalculation As XlCalculation
    Dim testPassed As Boolean
    
    ' Capture the original application state
    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalCalculation = Application.Calculation
    
    ' Initialize the class and capture the state
    Set monitor = New clsMonitorApplicationState
    monitor.CaptureState
    
    ' Assert that the captured state matches the original state
    Debug.Assert monitor.ScreenUpdating = originalScreenUpdating
    Debug.Assert monitor.Calculation = originalCalculation
    Debug.Assert monitor.DisplayAlerts = originalDisplayAlerts
    
    ' Change the application state
    Application.ScreenUpdating = Not originalScreenUpdating
    Application.DisplayAlerts = Not originalDisplayAlerts
    Application.Calculation = IIf(originalCalculation = xlCalculationAutomatic, xlCalculationManual, xlCalculationAutomatic)
    
    ' Restore the state using the class
    monitor.RestoreState
    
    ' Assert that the restored state matches the original state
    Debug.Assert Application.ScreenUpdating = originalScreenUpdating
    Debug.Assert Application.DisplayAlerts = originalDisplayAlerts
    Debug.Assert Application.Calculation = originalCalculation
    
    ' Cleanup
    Set monitor = Nothing
End Sub


