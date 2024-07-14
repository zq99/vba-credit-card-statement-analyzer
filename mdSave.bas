Attribute VB_Name = "mdSave"
Option Explicit


Public Sub SaveReport()
    
    Dim strPath As String
    Dim strName As String
    Dim strNewName As String
    Dim strMsg As String
    Dim intAnswer As Integer
        
    ThisWorkbook.Sheets("import").Calculate
    strPath = ThisWorkbook.Path
    strName = "credit_check"
    strName = strName & "_" + Format(Now, "YYYY-MM-DD") + ".xlsm"
    strNewName = strPath & "\" & strName
    
    If Dir(strNewName) <> "" Then
        intAnswer = MsgBox("Delete existing file?" & vbCrLf & strNewName, vbYesNo + vbQuestion, "Already Exists!")
        If intAnswer = vbNo Then
            GoTo ExitHere
        End If
        Kill strNewName
    End If
    
    On Error GoTo errHandler:
        ThisWorkbook.SaveAs strNewName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
ExitHere:
    Exit Sub
errHandler:
    strMsg = "Unable to save the tool to: "
    strMsg = strMsg & vbCrLf & strNewName
    strMsg = strMsg & vbCrLf & Err.Description & "[" & Err.Number & "]"
    MsgBox strMsg, vbCritical, "Archive"
    GoTo ExitHere
End Sub
