VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPivotTableCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsPivotTableCreator

' Purpose: Dynamically creates pivot tables

Option Explicit

Private mrngDataRange As Range
Private mrngPivotTableLocation As Range
Private mcollRowLabels As Collection
Private mcollValues As Collection
Private mcollFilterFields As Collection
Private mblnScreenUpdating As Boolean
Private mintCalc As Integer
Private mblnDisplayAlerts As Boolean

Private Sub CaptureState()
    mblnScreenUpdating = Application.ScreenUpdating
    mintCalc = Application.Calculation
    mblnDisplayAlerts = Application.DisplayAlerts
End Sub

Private Sub RestoreState()
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


Public Property Get dataRange() As Range
    Set dataRange = mrngDataRange
End Property

Public Property Let dataRange(ByVal value As Range)
    If IsValidDataRange(value) Then
        Set mrngDataRange = value
    Else
        Err.Raise vbObjectError + 1000, "PivotTableCreator", "Invalid data range."
    End If
End Property

Public Property Get PivotTableLocation() As Range
    Set PivotTableLocation = mrngPivotTableLocation
End Property

Public Property Let PivotTableLocation(ByVal value As Range)
    Set mrngPivotTableLocation = value
End Property

Public Sub AddRowLabel(ByVal fieldName As String)
    If mcollRowLabels Is Nothing Then Set mcollRowLabels = New Collection
    If FieldExists(fieldName) Then
        mcollRowLabels.Add fieldName
    Else
        Err.Raise vbObjectError + 1001, "PivotTableCreator", "Row label field '" & fieldName & "' does not exist in the data range."
    End If
End Sub

Public Sub AddValueField(ByVal fieldName As String, ByVal aggregationFunction As XlConsolidationFunction, Optional ByVal customName As String = "")
    Dim valueField As Variant
    If FieldExists(fieldName) Then
        Set valueField = CreateObject("Scripting.Dictionary")
        valueField("FieldName") = fieldName
        valueField("Function") = aggregationFunction
        If customName <> "" Then
            valueField("CustomName") = customName
        Else
            valueField("CustomName") = fieldName & " " & GetFunctionName(aggregationFunction)
        End If
        
        If mcollValues Is Nothing Then Set mcollValues = New Collection
        mcollValues.Add valueField
    Else
        Err.Raise vbObjectError + 1002, "PivotTableCreator", "Value field '" & fieldName & "' does not exist in the data range."
    End If
End Sub

Public Sub AddFilterField(ByVal fieldName As String)
    If mcollFilterFields Is Nothing Then Set mcollFilterFields = New Collection
    If FieldExists(fieldName) Then
        mcollFilterFields.Add fieldName
    Else
        Err.Raise vbObjectError + 1003, "PivotTableCreator", "Filter field '" & fieldName & "' does not exist in the data range."
    End If
End Sub

Public Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As pivotTable
    Dim pvtField As PivotField
    Dim i As Integer
    Dim pvtTableExists As Boolean
    
    Call CaptureState

    ClearObjects Application.ScreenUpdating, Application.DisplayAlerts
    
    On Error Resume Next
    For Each pvtTable In mrngPivotTableLocation.Parent.PivotTables
        If Not Intersect(pvtTable.TableRange2, mrngPivotTableLocation) Is Nothing Then
            pvtTable.TableRange2.Clear
            pvtTableExists = True
            Exit For
        End If
    Next pvtTable
    On Error GoTo 0
    
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mrngDataRange)
    
    
    Dim iCount As Integer
    iCount = mrngPivotTableLocation.Parent.PivotTables.Count + 1

    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=mrngPivotTableLocation, TableName:="PivotTable" & iCount)

    If Not mcollRowLabels Is Nothing Then
        For i = 1 To mcollRowLabels.Count
            Set pvtField = pvtTable.PivotFields(mcollRowLabels(i))
            pvtField.Orientation = xlRowField
            pvtField.Position = i
        Next i
    End If

    If Not mcollValues Is Nothing Then
        For i = 1 To mcollValues.Count
            Set pvtField = pvtTable.PivotFields(mcollValues(i)("FieldName"))
            pvtField.Orientation = xlDataField
            pvtField.Function = mcollValues(i)("Function")
            pvtField.Position = i
            
            On Error Resume Next
            pvtField.Caption = mcollValues(i)("CustomName")
            If Err.Number <> 0 Then
                Err.Clear
                pvtField.Caption = mcollValues(i)("CustomName") & " (" & i & ")"
            End If
            On Error GoTo 0
        Next i
    End If

    If Not mcollFilterFields Is Nothing Then
        For i = 1 To mcollFilterFields.Count
            Set pvtField = pvtTable.PivotFields(mcollFilterFields(i))
            pvtField.Orientation = xlPageField
            pvtField.Position = i
        Next i
    End If
    
    Call RestoreState
    
    ClearObjects ws, pvtCache, pvtTable, pvtField
    
End Sub

Private Function IsValidDataRange(ByVal value As Range) As Boolean
    If value Is Nothing Then
        IsValidDataRange = False
    ElseIf value.Rows.Count < 1 Or value.Columns.Count < 1 Then
        IsValidDataRange = False
    ElseIf value.Parent Is Nothing Then
        IsValidDataRange = False
    ElseIf value.Address = "" Then
        IsValidDataRange = False
    Else
        IsValidDataRange = True
    End If
End Function

Private Function FieldExists(ByVal fieldName As String) As Boolean
    Dim header As Range
    FieldExists = False
    For Each header In mrngDataRange.Rows(1).Cells
        If header.value = fieldName Then
            FieldExists = True
            Exit Function
        End If
    Next header
End Function

Private Function GetFunctionName(ByVal aggregationFunction As XlConsolidationFunction) As String
    Select Case aggregationFunction
        Case xlSum
            GetFunctionName = "Sum"
        Case xlCount
            GetFunctionName = "Count"
        Case xlAverage
            GetFunctionName = "Average"
        Case Else
            GetFunctionName = "Custom"
    End Select
End Function


Private Sub Class_Terminate()

    ClearObjects mrngDataRange, mrngPivotTableLocation, mcollRowLabels, mcollValues, mcollFilterFields

End Sub


Public Sub RefreshPivots()
    
    Dim Sheet As Worksheet, Pivot As pivotTable
    
    For Each Sheet In ThisWorkbook.Worksheets
        For Each Pivot In Sheet.PivotTables
            On Error Resume Next
            Pivot.RefreshTable
            Pivot.Update
        Next
    Next
    
    ClearObjects Sheet, Pivot
    
End Sub


Public Sub ClearPivotTables(ByVal shtName As String)

    Dim pvtTable
    Dim ws As Worksheet
    
    Call CaptureState
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Sheets(shtName)
    For Each pvtTable In ws.PivotTables
        pvtTable.TableRange2.Clear
    Next pvtTable
    
    Call RestoreState
    
    ClearObjects pvtTable, ws
    
End Sub

