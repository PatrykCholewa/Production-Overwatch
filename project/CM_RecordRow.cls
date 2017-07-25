VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CM_RecordRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private vValue() As String
Private vNumberOfValues As Integer

Public Property Get Value() As String()

    Value = vValue

End Property

Private Property Let NumberOfValues(num As Integer)

    ReDim Preserve vValue(num + 1)
    vNumberOfValues = num

End Property

Public Property Get NumberOfValues() As Integer

    NumberOfValues = vNumberOfValues

End Property

Public Sub ClearValues()

    ReDim vValue(0)
    NumberOfValues = 0

End Sub

Public Sub AddValue(val As String)

    NumberOfValues = NumberOfValues + 1
    vValue(NumberOfValues) = val
    
End Sub

Public Sub CopyValues(values() As String, numbOfValues As Integer)

    Dim i As Integer
    
    ClearValues
    
    For i = 1 To numbOfValues
        AddValue (values(i))
    Next i

End Sub

Public Sub CopyFromWorksheet(ws As Worksheet, FirstColumn As Long, NumberOfColumns, row As Long)

    Dim i As Long
    
    ClearValues

    For i = 1 To NumberOfColumns
        AddValue ws.Cells(row, FirstColumn + i - 1)
    Next i

End Sub

Public Sub WriteToWorksheet(ws As Worksheet, FirstColumn As Long, row As Long)

    Dim i As Long
    For i = 1 To vNumberOfValues
        ws.Cells(row, FirstColumn + i - 1) = vValue(i)
    Next i

End Sub

Public Sub SwitchOnWorksheet(ws As Worksheet, FirstColumn As Long, row As Long)

    Dim i As Long
    Dim tmp As String
    
    For i = 1 To NumberOfValues
        tmp = vValue(i)
        vValue(i) = Cells(row, FirstColumn + i - 1)
        ws.Cells(row, FirstColumn + i - 1) = tmp
    Next i

End Sub
