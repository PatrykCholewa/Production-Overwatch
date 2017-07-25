Attribute VB_Name = "LIB_ErrHandler"
Option Explicit

Sub Handle(number As Integer, source As String, description As String, helpFile As String, helpContext As String)
    If Err.number = 10000 Then
        MsgBox description, vbExclamation, "Warning!"
    Else
        MsgBox "Critical Error!" & vbNewLine & _
                "Contact developer for help!", _
                vbCritical, "Error"
    End If
End Sub
