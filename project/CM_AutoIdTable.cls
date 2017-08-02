VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CM_AutoIdTable"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private vTable As CM_RecordTable
Private vMaxId As Long

Public Property Get LastDataRow() As Long
    LastDataRow = vTable.LastDataRow
End Property

Public Property Get RecordTable() As CM_RecordTable
    Set RecordTable = vTable
End Property

Public Property Get FirstDataRow() As Long
    FirstDataRow = vTable.FirstDataRow
End Property

Public Property Get NumberOfColumns() As Integer
    NumberOfColumns = vTable.NumberOfColumns
End Property

Public Property Get FirstColumn() As Long
    FirstColumn = vTable.FirstColumn
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = vTable.Worksheet
End Property

Public Sub Constructor(ws As Worksheet, FirstColumn As Long, NumberOfColumns As Integer, FirstDataRow As Long, lastId As Long)
    Set vTable = New CM_RecordTable
    vTable.Constructor ws, FirstColumn, NumberOfColumns, FirstDataRow
    MaxId = lastId
End Sub

Public Property Get MaxId() As Long
        MaxId = vMaxId
End Property

Private Property Let MaxId(num As Long)
    vMaxId = num
End Property

Public Sub AddRecord(values() As String)
    
    Dim i As Long
    Dim CM_AIT_numberOfColumns As Integer
    Dim realValues() As String
    Dim ws As Worksheet
    
    CM_AIT_numberOfColumns = vTable.NumberOfColumns
    
    ReDim realValues(CM_AIT_numberOfColumns + 1)
    
    For i = 1 To CM_AIT_numberOfColumns
       realValues(i + 1) = values(i)
    Next i
    
    vMaxId = vMaxId + 1
    realValues(1) = vMaxId
    
    Set ws = vTable.Worksheet
    
    i = vTable.FirstDataRow
    Do While ws.Cells(i, vTable.FirstColumn) <> ""
        i = i + 1
    Loop
    
    vTable.PushRecord realValues, i
    
End Sub

Public Sub RemoveRecordById(id As Long)

    vTable.RemoveRecord CStr(id)

End Sub

Public Sub LostObjectAlert(name As String, column As Long, ws As Worksheet)

    vTable.LostObjectAlert name, column

End Sub

Public Sub Clear()

    vTable.Clear

End Sub

Public Sub CopyTo(copyTable As CM_AutoIdTable, copyMaxIdCell As String)

    Dim copyWs As Worksheet
    Set copyWs = copyTable.Worksheet
    
    vTable.CopyTo copyTable.RecordTable
    
    copyWs.Range(copyMaxIdCell) = vMaxId
    
End Sub

Public Sub CopyToIfColumnCellNotEmpty(copyTable As CM_AutoIdTable, copyMaxIdCell As String, ifColumn As Long)

    Dim copyWs As Worksheet
    Set copyWs = copyTable.Worksheet
    
    vTable.CopyToIfColumnCellNotEmpty copyTable.RecordTable, ifColumn
    
    copyWs.Range(copyMaxIdCell) = vMaxId
    
End Sub
