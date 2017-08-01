Attribute VB_Name = "MOD_User"
Option Explicit

Sub MOD_User_Create()

    UF_UserSetting.CreateUser

End Sub

Sub MOD_User_ModifyFunction()

    UF_UserSetting.ModifyFunction

End Sub

Sub MOD_User_ResetPassword()

    UF_UserSetting.ResetPassword
    
End Sub

Sub MOD_User_RemoveUser()

    Dim login As String
    Dim user As CM_User
    Set user = New CM_User
    
    On Error GoTo ErrHandler:
    
    login = InputBox("Insert username", "Remove user")
    
    If login = "" Then
        Err.Raise 10000, "MOD_User_RemoveUser", "There is no argument!"
    End If
    
    user.Constructor login
    
    If Not user.CheckExistence Then
        Err.Raise 10000, "MOD_User_RemoveUser", "User does not exists!"
    Else
        user.Remove
    End If
     
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub

Sub MOD_User_HideWS()

    WS_User.Visible = xlSheetHidden
End Sub
