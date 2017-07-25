Attribute VB_Name = "LIB_PriorityTable"
Option Explicit

Public Sub ChangePriority(howMuch As Long, row As Long, table As CM_AutoIdTable)

    Dim tmp As String
    Dim i As Long
    Dim ws As Worksheet
    Dim numOfColumns As Integer
    Dim firstCol As Long
    
    Set ws = table.Worksheet
    firstCol = table.FirstColumn
    
    If (row + howMuch < table.FirstDataRow) Or ws.Cells(row + howMuch, firstCol) = "" Then
        Err.Raise 10000, "LIB_PriorityTable.ChangePriority", "Getting beyond table!"
    End If
    
    For i = 1 To table.NumberOfColumns - 1
        tmp = ws.Cells(row, firstCol + i)
        ws.Cells(row, firstCol + i) = ws.Cells(row + howMuch, firstCol + i)
        ws.Cells(row + howMuch, firstCol + i) = tmp
    Next i

End Sub
