Attribute VB_Name = "mdImport"
Option Explicit

Private Const STR_SHEET_NAME As String = "data"
Public Const INT_START_OUTPUT_ROW As Integer = 4
Private Const STR_CELL_ADDRESS As String = "A" & INT_START_OUTPUT_ROW

Private Const STR_DATA_COLUMN_TRANSACTION_DATE As String = "A"
Private Const STR_DATA_COLUMN_DATE_ENTERED As String = "B"
Private Const STR_DATA_COLUMN_REFERENCE As String = "C"
Private Const STR_DATA_COLUMN_DESCRIPTION As String = "D"
Private Const STR_DATA_COLUMN_AMOUNT As String = "E"
Private Const STR_METRIC_COLUMN_FILENAME As String = "F"
Private Const STR_METRIC_COLUMN_DEBIT As String = "G"
Private Const STR_METRIC_COLUMN_CREDIT As String = "H"
Private Const STR_METRIC_COLUMN_TRANSACTION_TYPE As String = "I"
Private Const STR_METRIC_COLUMN_ACCUMLATED_DEBIT As String = "J"
Private Const STR_METRIC_COLUMN_GTR_THAN_10 As String = "K"

Private Type udtMetrics
    Label As String
    ColumnHeader As String
    Formula As String
End Type

Private arrFormattedSheets As Variant


Private Function InitFormattedSheetsList()
    arrFormattedSheets = Array("import")
End Function


Public Sub OpenFile()

    Dim strFilePath As String
    Dim sFile As String
    
    strFilePath = ThisWorkbook.Path
    sFile = fnOpenFileDialog(strFilePath)
    
    Call LoadFile(sFile)
    Call AddMetrics(sFile)
    Call ClearPivotTables
    Call CreateReports
    
End Sub


Public Sub ClearFile()
    Call ClearSheet(STR_SHEET_NAME)
    Call ClearPivotTables
End Sub


Private Sub AddMetrics(ByVal sFile As String)

    Dim ws As Worksheet
    Dim lngLastRow As Long
    Dim udtDataFields() As udtMetrics
    Dim i As Integer
    Dim rngColumn As Range
    Dim intArrayCount As Integer
    
    If IsWorksheetEmpty(STR_SHEET_NAME) Then
        GoTo Exit_here
    End If

    Set ws = ThisWorkbook.Sheets(STR_SHEET_NAME)
        
    lngLastRow = ws.Cells(ws.Rows.Count, STR_DATA_COLUMN_TRANSACTION_DATE).End(xlUp).Row
    
    intArrayCount = 1
    
    ReDim Preserve udtDataFields(1 To intArrayCount)
    udtDataFields(UBound(udtDataFields)).Formula = "=IF(RC[-2]>0, RC[-2], 0)"
    udtDataFields(UBound(udtDataFields)).ColumnHeader = STR_METRIC_COLUMN_DEBIT
    udtDataFields(UBound(udtDataFields)).Label = "Debit"
    
    intArrayCount = intArrayCount + 1
    ReDim Preserve udtDataFields(1 To intArrayCount)

    udtDataFields(UBound(udtDataFields)).Formula = "=IF(RC[-3]<0, RC[-3], 0)"
    udtDataFields(UBound(udtDataFields)).ColumnHeader = STR_METRIC_COLUMN_CREDIT
    udtDataFields(UBound(udtDataFields)).Label = "Credit"
    
    intArrayCount = intArrayCount + 1
    ReDim Preserve udtDataFields(1 To intArrayCount)

    udtDataFields(UBound(udtDataFields)).Formula = "=IF(RC[-4]<0, ""CREDIT"", ""DEBIT"")"
    udtDataFields(UBound(udtDataFields)).ColumnHeader = STR_METRIC_COLUMN_TRANSACTION_TYPE
    udtDataFields(UBound(udtDataFields)).Label = "Type"
    
    intArrayCount = intArrayCount + 1
    ReDim Preserve udtDataFields(1 To intArrayCount)
    
    udtDataFields(UBound(udtDataFields)).Formula = "=SUM(R5C[-3]:RC[-3])"
    udtDataFields(UBound(udtDataFields)).ColumnHeader = STR_METRIC_COLUMN_ACCUMLATED_DEBIT
    udtDataFields(UBound(udtDataFields)).Label = "Accumulated_Debit"
    
    intArrayCount = intArrayCount + 1
    ReDim Preserve udtDataFields(1 To intArrayCount)
    
    udtDataFields(UBound(udtDataFields)).Formula = "=IF(RC[-4]>10, 1, 0)"
    udtDataFields(UBound(udtDataFields)).ColumnHeader = STR_METRIC_COLUMN_GTR_THAN_10
    udtDataFields(UBound(udtDataFields)).Label = "Greater_than_10"
    
    For i = LBound(udtDataFields) To UBound(udtDataFields)
        ws.Cells(INT_START_OUTPUT_ROW, udtDataFields(i).ColumnHeader).value = udtDataFields(i).Label
        Set rngColumn = ws.Range(udtDataFields(i).ColumnHeader & INT_START_OUTPUT_ROW + 1 & ":" & udtDataFields(i).ColumnHeader & lngLastRow)
        rngColumn.FormulaR1C1 = udtDataFields(i).Formula
        rngColumn.value = rngColumn.value
    Next i

Exit_here:
    
    ClearObjects ws, rngColumn

End Sub


Private Function fnOpenFileDialog(ByVal strFolderPath As String) As String

    Dim fd As FileDialog
    Dim strFileName As String
    
    fnOpenFileDialog = Empty
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = strFolderPath
    
    With fd
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Title = "Select a File"
    End With
    
    If fd.Show = -1 Then
    
        strFileName = fd.SelectedItems(1)
        fnOpenFileDialog = strFileName
        
    End If
    
    ClearObjects fd
    
End Function


Private Sub LoadFile(ByVal sFile As String)
    
    Dim oState As New clsMonitorApplicationState
    Dim intUserResponse As Integer
    Dim blnExcludeHeader As Boolean
    
    oState.CaptureState

    If Len(Trim(sFile)) > 0 Then
    
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        If Not IsWorksheetEmpty(STR_SHEET_NAME) Then
            intUserResponse = MsgBox("Existing data detected" & vbCrLf & "Do you want to clear this?", vbYesNoCancel + vbQuestion, "Import")
            If intUserResponse = vbCancel Then
                GoTo Exit_here
            ElseIf intUserResponse = vbYes Then
                Call ClearSheet(STR_SHEET_NAME)
                blnExcludeHeader = False
            Else
                blnExcludeHeader = True
            End If
        End If
        
        Call AddCSVToSheet(STR_SHEET_NAME, sFile, blnExcludeHeader)
        
        Call SortData
    
    End If

Exit_here:

    oState.RestoreState
    
    ClearObjects oState
    
End Sub


Private Sub SortData()
    
On Error GoTo err_handler:

    Dim ws As Worksheet
    Dim lngLastRow As Long
    Dim strSortKeyColumnAddress As String
    Dim strSortRangeAddress As String
    
    Set ws = ThisWorkbook.Worksheets(STR_SHEET_NAME)
    
    lngLastRow = ws.Cells(ws.Rows.Count, STR_DATA_COLUMN_TRANSACTION_DATE).End(xlUp).Row
    
    strSortKeyColumnAddress = "A5:A" & lngLastRow
    strSortRangeAddress = "A4:K" & lngLastRow

    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add Key:=ws.Range(strSortKeyColumnAddress), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(strSortRangeAddress)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Exit_here:

    ClearObjects ws
    Exit Sub

err_handler:

    GoTo Exit_here
    
End Sub



Private Sub AddCSVToSheet(ByVal strShtName As String, ByVal file As String, ByVal blnExcludeHeader As Boolean, Optional ByVal adjustColumns As Boolean = True)

On Error GoTo err_handler:

    Dim ws As Worksheet
    Dim lngNextRow As Long
    Dim lngEndRow As Long
    Dim rngFilenameColumn As Range
    Dim oImporter As clsCSVImporter
    Dim blnResult As Boolean

    Set ws = ThisWorkbook.Sheets(strShtName)

    If ws.AutoFilterMode = True Then
        ws.AutoFilterMode = False
    End If

    Call UnfreezePanesOnSheet(strShtName)
    
    If IsWorksheetEmpty(strShtName) Then
        lngNextRow = INT_START_OUTPUT_ROW
    Else
        lngNextRow = ws.Cells(ws.Rows.Count, STR_DATA_COLUMN_TRANSACTION_DATE).End(xlUp).Row + 1
    End If
            
    Set oImporter = New clsCSVImporter
    blnResult = oImporter.ImportCSVtoSheet(file, strShtName, lngNextRow, blnExcludeHeader)
        
    If Not blnResult Then
        GoTo Exit_here
    End If
    
    If Not blnExcludeHeader Then
        If lngNextRow = INT_START_OUTPUT_ROW Then
            ws.Cells(INT_START_OUTPUT_ROW, STR_METRIC_COLUMN_FILENAME).value = "Filename"
        End If
    End If

    lngEndRow = ws.Cells(ws.Rows.Count, STR_DATA_COLUMN_TRANSACTION_DATE).End(xlUp).Row
    
    Set rngFilenameColumn = ws.Range(STR_METRIC_COLUMN_FILENAME & IIf(blnExcludeHeader, lngNextRow, INT_START_OUTPUT_ROW + 1) & ":" & STR_METRIC_COLUMN_FILENAME & lngEndRow)
    rngFilenameColumn.value = Dir(file)
    
    ws.AutoFilterMode = True
    

Exit_here:

    ClearObjects ws, rngFilenameColumn, oImporter
    Exit Sub

err_handler:
    Debug.Print "Error in  'AddCSVToSheet' : (Error number: " & Err.Number & ") " & Err.Description
    GoTo Exit_here
End Sub


Private Sub UnfreezePanesOnSheet(ByVal shtName As String)
    
    Dim w As Window
    Dim wsv
    Dim activews As Worksheet
    Dim ws As Worksheet
    
    For Each w In ThisWorkbook.Windows
        w.Activate
        If activews Is Nothing Then
            Set activews = w.ActiveSheet
        End If
        For Each wsv In w.SheetViews
            If wsv.Sheet.Name = shtName Then
                wsv.Sheet.Activate
                w.FreezePanes = False
                Exit For
            End If
        Next
        activews.Activate
        Set activews = Nothing
    Next
    
    ClearObjects w, wsv, activews, ws
    
End Sub


Private Sub ClearSheet(ByVal strSheet As String, Optional ByVal strInputRange As String = "")
    
    Dim wkb As Workbook
    Dim wsSht As Worksheet
    Dim rng As Range
    
    Set wkb = ThisWorkbook
    Set wsSht = wkb.Sheets(strSheet)
    
    wsSht.AutoFilterMode = False
    If strInputRange <> Empty Then
        On Error Resume Next
        wsSht.Range(strInputRange).value = Empty
        On Error GoTo 0
    End If
    Set rng = wsSht.Range(STR_CELL_ADDRESS).CurrentRegion
    rng.ClearContents
    
    ClearObjects rng, wkb, wsSht

End Sub
 
 
Private Function IsFormattedSheet(ByVal shtName As String) As Boolean
    
    Debug.Print ("checking sheet " & shtName & " is protected...")
    Dim i As Integer
    For i = LBound(arrFormattedSheets) To UBound(arrFormattedSheets)
        If shtName = arrFormattedSheets(i) Then
            IsFormattedSheet = True
            Exit Function
        End If
    Next
    IsFormattedSheet = False
    
End Function


Private Sub RemoveAutoFilters()

    Dim sht As Worksheet

    For Each sht In ThisWorkbook.Worksheets
        If Not IsFormattedSheet(sht.Name) Then
            If sht.AutoFilterMode = True Then
                sht.AutoFilterMode = False
            End If
        End If
    Next

    ClearObjects sht

End Sub







