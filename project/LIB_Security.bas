Attribute VB_Name = "LIB_Security"
Option Explicit

Const WS_Pass_firstDataRow As Long = 1
Const WS_Pass_loginColumn As Long = 1
Const WS_Pass_passColumn As Long = 2
Const WS_Pass_numberOfColumns As Integer = 2

Const WS_User_firstDataRow As Long = 4
Const WS_User_loginColumn As Long = 1
Const WS_User_functionColumn As Long = 2
Const WS_User_numberOfColumns As Integer = 2

Public Function Encrypt(pass As String, key As Long) As String
'Inpired by https://access-programmers.co.uk/forums/showthread.php?t=172530 , 28.07.2017 , 0:33

    Dim i As Integer

    For i = 1 To Len(pass)
        Encrypt = Encrypt & CStr(Asc(Mid(pass, i, 1)) Xor key)
    Next i
    
End Function
