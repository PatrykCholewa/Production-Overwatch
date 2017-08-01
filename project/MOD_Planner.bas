Attribute VB_Name = "MOD_Planner"
Option Explicit

Const WS_Planner_firstDataRow As Long = 4
Const WS_Planner_idColumn As Long = 1
Const WS_Planner_quantityColumn As Long = 2
Const WS_Planner_productColumn As Long = 3
Const WS_Planner_kitColumn As Long = 4
Const WS_Planner_materialColumn As Long = 5
Const WS_Planner_numberOfColumns As Integer = 5
Const WS_Planner_maxIdRow As Long = 2
Const WS_Planner_maxIdColumn As Long = 5
Const WS_Planner_statusCell As String = "E1"

Const WS_Plan_firstDataRow As Long = 4
Const WS_Plan_idColumn As Long = 1
Const WS_Plan_productColumn As Long = 6
Const WS_Plan_numberOfFirstColumns As Long = 2
Const WS_Plan_numberOfSecondColumns As Long = 3
Const WS_Plan_maxIdCell As String = "C1"

Const WS_Reporter_statusCell As String = "E1"

Sub MOD_Planner_AddRecord()
    
    On Error GoTo ErrHandler:

    UF_AddPlanerRecord.Show
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Planner_RemoveSelectedRecord()

    On Error GoTo ErrHandler:

    Dim id As Long
    Dim table As CM_AutoIdTable
    Dim lastId As Long
    Dim row As Long
    
    row = ActiveCell.row
    
    If row < WS_Planner_firstDataRow Then
        Err.Raise 10000, "MOD_Planner_RemoveSelectedRow", "There is no data!"
    End If
    
    id = WS_Planner.Cells(row, WS_Planner_idColumn)
    
    lastId = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)

    Set table = New CM_AutoIdTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow, lastId
    table.RemoveRecordById id
    
    WS_Planner.Range(WS_Planner_statusCell) = "Not sent"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub
    
Sub MOD_Planner_SelectedRecordPriorityUp()
    
    On Error GoTo ErrHandler:

    Dim id As Long
    Dim table As CM_AutoIdTable
    Dim lastId As Long
    Dim row As Long
    
    row = ActiveCell.row
    lastId = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)
    
    If row < WS_Planner_firstDataRow Then
        Err.Raise 10000, "MOD_Planner_RemoveSelectedRow", "There is no data!"
    End If
    
    If WS_Planner.Cells(row, WS_Planner_idColumn) = "" Then
        Err.Raise 10000, "MOD_Planner_RemoveSelectedRow", "There is no data!"
    End If
    
    Set table = New CM_AutoIdTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow, lastId
    LIB_PriorityTable.ChangePriority -1, row, table
    
    WS_Planner.Range(WS_Planner_statusCell) = "Not sent"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Planner_SelectedRecordPriorityDown()
    
    On Error GoTo ErrHandler:

    Dim id As Long
    Dim table As CM_AutoIdTable
    Dim lastId As Long
    Dim row As Long
    
    row = ActiveCell.row
    lastId = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)
    
    If row < WS_Planner_firstDataRow Then
        Err.Raise 10000, "MOD_Planner_RemoveSelectedRow", "There is no data!"
    End If
    
    If WS_Planner.Cells(row, WS_Planner_idColumn) = "" Then
        Err.Raise 10000, "MOD_Planner_RemoveSelectedRow", "There is no data!"
    End If
    
    Set table = New CM_AutoIdTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow, lastId
    LIB_PriorityTable.ChangePriority 1, row, table
    
    WS_Planner.Range(WS_Planner_statusCell) = "Not sent"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub


Sub MOD_Planner_ClearTable()
    
    On Error GoTo ErrHandler:

    Dim table As CM_AutoIdTable
    Dim lastId As Long
    
    lastId = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)
    
    Set table = New CM_AutoIdTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow, lastId
    table.Clear
    
    WS_Planner.Range(WS_Planner_statusCell) = "Not sent"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Planner_SendTable()

    On Error GoTo ErrHandler:

    Dim plannerFirstTable As CM_AutoIdTable
    Dim plannerSecondTable As CM_RecordTable
    Dim planFirstTable As CM_AutoIdTable
    Dim planSecondTable As CM_RecordTable
    Dim lastId As Long
    
    lastId = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)
    
    Set plannerFirstTable = New CM_AutoIdTable
    Set plannerSecondTable = New CM_RecordTable
    Set planFirstTable = New CM_AutoIdTable
    Set planSecondTable = New CM_RecordTable
    
    plannerFirstTable.Constructor WS_Planner, WS_Planner_idColumn, 2, WS_Planner_firstDataRow, lastId
    plannerSecondTable.Constructor WS_Planner, WS_Planner_productColumn, WS_Planner_numberOfColumns - 2, WS_Planner_firstDataRow
    planFirstTable.Constructor WS_Plan, WS_Plan_idColumn, WS_Plan_numberOfFirstColumns, WS_Plan_firstDataRow, WS_Plan.Range(WS_Plan_maxIdCell).Value
    planSecondTable.Constructor WS_Plan, WS_Plan_productColumn, WS_Plan_numberOfSecondColumns, WS_Plan_firstDataRow
    
    plannerFirstTable.CopyTo planFirstTable, WS_Plan_maxIdCell
    plannerSecondTable.CopyTo planSecondTable
    
    WS_Planner.Range(WS_Planner_statusCell) = "Sent"
    WS_Reporter.Range(WS_Reporter_statusCell) = "Not Current"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub

Sub MOD_Planner_HideWS()

    WS_Planner.Visible = xlSheetHidden
    WS_Plan.Visible = xlSheetHidden
    WS_Report.Visible = xlSheetHidden

End Sub
