VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMonitorApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsMonitorApplicationState

Option Explicit

' Purpose: This class allows you to record the state of the current application
'          before changing any of the settings in a function or subroutine.
'          After a routine is complete, you can restore the settings back to what they
'          were before the routine ran

Private mblnScreenUpdating As Boolean
Private mblnDisplayAlerts As Boolean
Private mintCalc As Integer


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

Public Property Get ScreenUpdating() As Boolean
    ScreenUpdating = mblnScreenUpdating
End Property

Public Property Get DisplayAlerts() As Boolean
    DisplayAlerts = mblnDisplayAlerts
End Property

Public Property Get Calculation() As XlCalculation
    Calculation = mintCalc
End Property
