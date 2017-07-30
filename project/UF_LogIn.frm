VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_LogIn 
   OleObjectBlob   =   "UF_LogIn.frx":0000
   Caption         =   "Log In"
   ClientHeight    =   1605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2955
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   8
End
Attribute VB_Name = "UF_LogIn"
Attribute VB_Base = "0{03782E00-6BDD-4845-9138-D7D0E439BA27}{3237546C-19C0-4F97-9DBA-352AB8CFBA43}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const loginCell As String = "D2"
Const functionCell As String = "H2"

Private Sub CBT_Login_Click()

    On Error GoTo ErrHandler
    
    Dim user As CM_User
    Set user = New CM_User
    user.Constructor Me.TBX_Login.value
    
    If Not user.Authorize(Me.TBX_Pass) Then
        Err.Raise 10000, "UF_LogIn.CBT_LogIn_Click()", "Wrong password!"
    End If
    
    WS_Start.Range(loginCell) = Me.TBX_Login
    WS_Start.Range(functionCell) = user.Funct
    
    Select Case user.Funct
        Case "Admin"
            MOD_Start.MOD_Start_Admin
        Case "Manager"
            MOD_Start.MOD_Start_Manager
        Case "Worker"
            MOD_Start.MOD_Start_Worker
        Case Else
            Err.Raise 10001, "UF_LogIn.CBT_LogIn_Click()", "Unknown user function!"
    End Select
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_Activate()
 
End Sub

Private Sub UserForm_Click()

End Sub
