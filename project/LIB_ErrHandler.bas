Attribute VB_Name = "LIB_ErrHandler"
Option Explicit

Sub Handle(number As Integer, source As String, description As String, helpFile As String, helpContext As String)
    Select Case Err.number
        Case 10000
            MsgBox description, vbExclamation, "Warning!"
        Case Is >= 10002
            MsgBox description, vbExclamation, "Warning!"
        Case 10001
            ProtectEverything
            MsgBox "Critical Error!" & vbNewLine & _
                "Contact developer for help!", _
                vbCritical, "Error"
        Case Else
            ProtectEverything
            MsgBox "Critical Error!" & vbNewLine & _
                "Contact developer for help!", _
                vbCritical, "Error"
    End Select
End Sub


Sub ProtectEverything()
    WS_Objects.Protect UserInterfaceOnly:=True
    WS_Plan.Protect UserInterfaceOnly:=True
    WS_Planner.Protect UserInterfaceOnly:=True
    WS_Report.Protect UserInterfaceOnly:=True
    WS_Reporter.Protect UserInterfaceOnly:=True
End Sub
