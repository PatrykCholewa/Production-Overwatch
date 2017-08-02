Attribute VB_Name = "MOD_Archives"
Option Explicit

Const WS_Reporter_firstDataRow As Long = 4
Const WS_Reporter_idColumn As Long = 1
Const WS_Reporter_amountColumn As Long = 3
Const WS_Reporter_numberOfColumns As Integer = 8
Const WS_Reporter_maxIdCell As String = "A1"
Const WS_Reporter_statusCell As String = "E1"

Const WS_Archives_firstDataRow As Long = 4
Const WS_Archives_idColumn As Long = 1
Const WS_Archives_numberOfColumns As Integer = 9
Const WS_Archives_maxIdCell As String = "A1"
Const WS_Archives_dateColumn As Long = 9

Sub MOD_Archives_HideWS()
    WS_Archives.Visible = xlSheetHidden
End Sub

Sub MOD_Archives_AddReport()

    On Error GoTo ErrHandler

    Dim repTable As CM_AutoIdTable
    Dim archTable As CM_AutoIdTable
    Dim copyArchTableFront As CM_AutoIdTable
    Dim archTableLastRow As Long
    Dim row As Long

    Set repTable = New CM_AutoIdTable
    repTable.Constructor WS_Reporter, WS_Reporter_idColumn, WS_Reporter_numberOfColumns, _
                        WS_Reporter_firstDataRow, WS_Reporter.Range(WS_Reporter_maxIdCell)

    Set archTable = New CM_AutoIdTable
    archTable.Constructor WS_Archives, WS_Archives_idColumn, WS_Archives_numberOfColumns, _
                        WS_Archives_firstDataRow, WS_Archives.Range(WS_Archives_maxIdCell)
                        
    archTableLastRow = archTable.LastDataRow
                        
    Set copyArchTableFront = New CM_AutoIdTable
    copyArchTableFront.Constructor WS_Archives, WS_Archives_idColumn, WS_Archives_numberOfColumns - 1, _
                        archTableLastRow + 1, WS_Archives.Range(WS_Archives_maxIdCell)
                        
    repTable.CopyToIfColumnCellNotEmpty copyArchTableFront, WS_Archives_maxIdCell, WS_Reporter_amountColumn
    
    For row = copyArchTableFront.FirstDataRow To copyArchTableFront.LastDataRow
        WS_Archives.Cells(row, WS_Archives_dateColumn) = Date
    Next row
    
    Exit Sub
    
ErrHandler:
    If Err.number = 10005 Then
        archTableLastRow = archTable.FirstDataRow - 1
        Resume Next
    Else
        LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    End If
End Sub
