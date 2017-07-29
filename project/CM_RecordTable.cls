VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CM_RecordTable"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private vWorksheet As Worksheet
Private vFirstColumn As Long
Private vNumberOfColumns As Integer
Private vFirstDataRow As Long

Public Property Get Worksheet() As Worksheet
    Set Worksheet = vWorksheet
End Property

Public Property Get NumberOfColumns() As Integer
    NumberOfColumns = vNumberOfColumns
End Property

Public Property Get FirstColumn() As Long
    FirstColumn = vFirstColumn
End Property

Public Property Get FirstDataRow() As Long
    FirstDataRow = vFirstDataRow
End Property

Public Property Let Worksheet(ws As Worksheet)
    vWorksheet = ws
End Property

Private Property Let NumberOfColumns(number As Integer)
    vNumberOfColumns = number
End Property

Private Property Let FirstColumn(number As Long)
    vFirstColumn = number
End Property

Private Property Let FirstDataRow(number As Long)
    vFirstDataRow = number
End Property

Public Sub Constructor(ws As Worksheet, firstCol As Long, numbOfColumns As Integer, FirstRow As Long)

    Set vWorksheet = ws
    vFirstColumn = firstCol
    vNumberOfColumns = numbOfColumns
    vFirstDataRow = FirstRow

End Sub

Private Sub PushRecordInConcreteLine(record As CM_RecordRow, row As Long)
    
    Dim i As Long
    
    i = row
    
    Do While Not IsEmpty(vWorksheet.Cells(i, FirstColumn))
        record.SwitchOnWorksheet vWorksheet, FirstColumn, i
        i = i + 1
    Loop
    record.WriteToWorksheet vWorksheet, FirstColumn, i
    
End Sub

Public Sub PushRecord(values() As String, row As Long)

    Dim rec As CM_RecordRow
    Set rec = New CM_RecordRow
    rec.CopyValues values, NumberOfColumns
    
    PushRecordInConcreteLine rec, row

End Sub

Public Sub AddRecord(values() As String)

    Dim i As Long
    Dim rec As CM_RecordRow
    
    i = FirstDataRow
    Set rec = New CM_RecordRow
    rec.CopyValues values, NumberOfColumns
    
    Do While (StrComp(vWorksheet.Cells(i, FirstColumn), values(1), vbTextCompare) = -1)
        If IsEmpty(vWorksheet.Cells(i, FirstColumn)) Then
            rec.WriteToWorksheet vWorksheet, FirstColumn, i
            Exit Sub
        End If
        i = i + 1
    Loop
    If StrComp(vWorksheet.Cells(i, FirstColumn), values(1), vbTextCompare) = 0 Then
        Err.Raise 10002, "LIB_Records", "Object already exists."
    End If
    
    PushRecordInConcreteLine rec, i
    
End Sub

Private Sub RemoveConcreteRecordLine(row As Long)

    Dim i As Long
    Dim rec As CM_RecordRow
    Set rec = New CM_RecordRow
    
    i = row
    Do While Cells(i, FirstColumn) <> ""
        rec.CopyFromWorksheet vWorksheet, FirstColumn, NumberOfColumns, i + 1
        rec.WriteToWorksheet vWorksheet, FirstColumn, i
        i = i + 1
    Loop

End Sub

Public Sub RemoveRecord(name As String)

    Dim i As Long
    
    i = FirstDataRow
    Do While vWorksheet.Cells(i, FirstColumn) <> ""
        If StrComp(vWorksheet.Cells(i, FirstColumn), name, vbTextCompare) = 0 Then
            RemoveConcreteRecordLine i
            Exit Sub
        End If
        i = i + 1
    Loop
    
    Err.Raise 10003, "LIB_Records", "Record does not exist."

End Sub

'Returns True if finds lost object in column
Public Sub LostObjectAlert(name As String, column As Long)

    Dim row As Long
    row = FirstDataRow
    Do While (vWorksheet.Cells(row, column) <> "")
        If vWorksheet.Cells(row, column) = name Then
            Dim ret As Integer
            vWorksheet.Activate
            vWorksheet.Cells(row, column).Select
            ret = MsgBox("Delete that record?", 51, "Question")
            Select Case ret
                Case vbYes
                    RemoveConcreteRecordLine row
                    row = row - 1
                Case vbNo
                    'Do Nothing
                Case vbCancel
                    Err.Raise 10000, "CM_RecordTable.LostObjectAlert", "Action abort!"
            End Select
        End If
        row = row + 1
    Loop

End Sub

Public Function GetRowIndexOf(name As String, column As Long) As Long
    
    Dim i As Long
    i = FirstDataRow
    
    Do While vWorksheet.Cells(i, column) <> ""
        If vWorksheet.Cells(i, column) = name Then
            GetRowIndexOf = i
            Exit Function
        End If
        i = i + 1
    Loop
    
    Err.Raise 10003, "CM_RecordTable.GetRowIndexOf", "There is no such a record!"
    
End Function

Public Sub Clear()

    Dim i As Long
    Dim j As Integer
    
    i = FirstDataRow
    Do While vWorksheet.Cells(i, vFirstColumn) <> ""
        For j = 1 To vNumberOfColumns
            vWorksheet.Cells(i, vFirstColumn + j - 1) = ""
        Next j
        i = i + 1
    Loop

End Sub

Public Sub CopyTo(copyTable As CM_RecordTable)

    Dim row As Long
    Dim column As Long
    Dim columnDiff As Long
    Dim rowDiff As Long
    Dim copyWs As Worksheet
    
    Set copyWs = copyTable.Worksheet

    row = vFirstDataRow
    rowDiff = copyTable.FirstDataRow - vFirstDataRow
    column = vFirstColumn
    columnDiff = copyTable.FirstColumn - vFirstColumn
    
    copyTable.Clear
    
    Do While vWorksheet.Cells(row, vFirstColumn) <> ""
        For column = vFirstColumn To vFirstColumn + vNumberOfColumns - 1
            copyWs.Cells(row + rowDiff, column + columnDiff) = vWorksheet.Cells(row, column)
        Next column
        row = row + 1
    Loop
        
End Sub

Public Function GetCellValueOf(id As String, column As Integer) As String

    Dim ind As Long
    ind = GetRowIndexOf(id, vFirstColumn)

    GetCellValueOf = vWorksheet.Cells(ind, column)

End Function

Public Sub SetCellValueOf(id As String, column As Integer, value As String)

    Dim ind As Long
    ind = GetRowIndexOf(id, vFirstColumn)
    
    vWorksheet.Cells(ind, column) = value
    
End Sub
