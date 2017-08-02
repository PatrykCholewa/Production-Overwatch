Attribute VB_Name = "LIB_ErrHandler"
Option Explicit

Sub Handle(number As Integer, source As String, description As String, helpFile As String, helpContext As String)
    Select Case Err.number
        Case 10000
            MsgBox description, vbExclamation, "Warning!"
        Case Is >= 10002
            MsgBox description, vbExclamation, "Warning!"
        Case 10001
            DEV_Tools.DEV_ProtectEverything
            MsgBox "Critical Error!" & vbNewLine & _
                "Contact developer for help!", _
                vbCritical, "Error"
        Case Else
            DEV_Tools.DEV_ProtectEverything
            MsgBox "Critical Error!" & vbNewLine & _
                "Contact developer for help!", _
                vbCritical, "Error"
    End Select
End Sub
