VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_UserSetting 
   OleObjectBlob   =   "UF_UserSetting.frx":0000
   Caption         =   "Create/Modify User"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3120
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   25
End
Attribute VB_Name = "UF_UserSetting"
Attribute VB_Base = "0{BA16E11F-3B79-4796-8B7B-E6C16763910C}{FEE82CFE-ADA8-4F65-8326-11CA8E14BC34}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const WS_Pass_firstDataRow As Long = 1
Const WS_Pass_loginColumn As Long = 1
Const WS_Pass_passColumn As Long = 2
Const WS_Pass_numberOfColumns As Integer = 2

Const WS_User_firstDataRow As Long = 4
Const WS_User_loginColumn As Long = 1
Const WS_User_functionColumn As Long = 2
Const WS_User_numberOfColumns As Integer = 2

Private Sub CBT_Set_Click()

    On Error GoTo ErrHandler
    
    Dim confirm As Integer
    Dim user As CM_User
    Set user = New CM_User
    user.Constructor Me.TBX_Login
    
    If Not user.CheckExistence Then
        
        If Me.TBX_Pass1.Visible = False Then
            Err.Raise 10000, "UF_UserSetting.CBT_Set_Click", "Lack of Password Textbox!"
        End If
        
        If Me.TBX_Pass1 <> Me.TBX_Pass2 Then
            Err.Raise 10000, "UF_UserSetting.CBT_Set_Click", "Passwords do not match!"
        End If
        
        confirm = MsgBox("Create user?", vbYesNo)
        If confirm = vbNo Then
            Err.Raise 10000, "MsgBox", "Aborted!"
        End If
        
        user.Create Me.TBX_Pass1, Me.CBX_Function
        
    End If
    
    If Me.TBX_Pass1.Visible = True Then
        If Not user.Authorize(Me.TBX_Pass1) Then
            Err.Raise 10000, "UF_UserSetting.CBT_Set_Click()", "Wrong password!"
        End If
    End If
    
    confirm = MsgBox("Modify user?", vbYesNo)
    If confirm = vbNo Then
        Err.Raise 10000, "MsgBox", "Aborted!"
    End If
    
    If Me.CBX_Function.Visible = True Then
            user.ChangeFunct Me.CBX_Function
    End If
    
    If Me.TBX_PassN.Visible Then
        If Me.TBX_PassN = Me.TBX_Pass2 Then
            user.ChangePass Me.TBX_PassN
        Else
            Err.Raise 10000, "UF_UserSetting.CBT_Set_Click", "Passwords do not match!"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub

Private Sub TBX_Function_Change()

End Sub

Private Sub UserForm_Activate()
 
    Me.CBX_Function.AddItem "Worker"
    Me.CBX_Function.AddItem "Manager"
    Me.CBX_Function.AddItem "Admin"
 
End Sub

Public Sub CreateUser()

    Me.LAB_PassN.Visible = False
    Me.TBX_PassN.Visible = False
    
    Me.Show
    
End Sub

Public Sub ModifyFunction()

    Me.LAB_Pass1.Visible = False
    Me.TBX_Pass1.Visible = False
    
    Me.LAB_PassN.Visible = False
    Me.TBX_PassN.Visible = False
    
    Me.LAB_Pass2.Visible = False
    Me.TBX_Pass2.Visible = False
    
    Me.Show
    
End Sub

Public Sub ResetPassword()

    Me.LAB_Function.Visible = False
    Me.CBX_Function.Visible = False
    
    Me.LAB_Pass1.Visible = False
    Me.TBX_Pass1.Visible = False
    
    Me.Show
    
End Sub

Public Sub ChangePassword(login As String)

    Dim user As CM_User
    Set user = New CM_User
    user.Constructor login

    Me.CBX_Function.Enabled = False
    Me.CBX_Function.value = user.Funct
    
    Me.TBX_Login.Enabled = False
    Me.TBX_Login.value = login
    
    Me.Show

End Sub
