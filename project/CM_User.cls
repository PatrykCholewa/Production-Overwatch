VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CM_User"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const loginColumn As Long = 1
Const passColumn As Long = 2
Const functionColumn As Long = 2
Const numberOfColumns As Integer = 2
Const firstUserRow As Long = 4
Const firstPassRow As Long = 1

Const loginKey As Long = 107
Const passKey As Long = 97

Dim userTable As CM_RecordTable
Dim passTable As CM_RecordTable
Dim login As String

Private Property Get LoginHash() As String

    LoginHash = LIB_Security.Encrypt(login, loginKey)

End Property

Public Property Get Funct() As String

    Dim ind As Long
    ind = userTable.GetRowIndexOf(login, loginColumn)
    
    Funct = userTable.GetCellValueOf(login, functionColumn)
    
End Property

Private Property Get PassHash() As String
    
    PassHash = passTable.GetCellValueOf(LoginHash, passColumn)

End Property

Public Function Authorize(pass As String) As Boolean

    If LIB_Security.Encrypt(pass, passKey) = PassHash Then
        Authorize = True
    Else
        Authorize = False
    End If
    
End Function

Public Sub Constructor(username As String)

Set userTable = New CM_RecordTable
Set passTable = New CM_RecordTable
userTable.Constructor WS_User, loginColumn, numberOfColumns, firstUserRow
passTable.Constructor WS_Pass, loginColumn, numberOfColumns, firstPassRow

login = username

End Sub

Public Sub Create(pass As String, fun As String)

    Dim userValues() As String
    Dim passValues() As String
    ReDim userValues(2)
    ReDim passValues(2)
    
    passValues(1) = LIB_Security.Encrypt(login, loginKey)
    passValues(2) = LIB_Security.Encrypt(pass, passKey)
    WS_Pass.Activate
    passTable.AddRecord passValues
    
    userValues(1) = login
    userValues(2) = fun
    WS_User.Activate
    userTable.AddRecord userValues

End Sub

Public Sub ChangePass(newPass As String)

    passTable.SetCellValueOf LoginHash, passColumn, LIB_Security.Encrypt(newPass, passKey)

End Sub

Public Sub ChangeFunct(fun As String)

    userTable.SetCellValueOf login, functionColumn, fun

End Sub

Public Sub Remove()

    WS_Pass.Activate
    passTable.RemoveRecord LoginHash
    
    WS_User.Activate
    userTable.RemoveRecord login
    
End Sub

Public Function CheckExistence() As Boolean

    On Error GoTo ErrHandler
    
    CheckExistence = True
    
    userTable.GetRowIndexOf login, loginColumn
    
    Exit Function
    
ErrHandler:
    If Err.number = 10003 Then
        CheckExistence = False
        Exit Function
    Else
        LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    End If
End Function
