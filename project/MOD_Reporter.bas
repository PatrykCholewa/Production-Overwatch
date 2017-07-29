Attribute VB_Name = "MOD_Reporter"
Option Explicit

Const WS_Planner_statusCell As String = "E1"

Const WS_Plan_firstDataRow As Long = 4
Const WS_Plan_idColumn As Long = 1
Const WS_Plan_numberOfColumns As Integer = 8
Const WS_Plan_maxIdCell As String = "C1"

Const WS_Reporter_firstDataRow As Long = 4
Const WS_Reporter_idColumn As Long = 1
Const WS_Reporter_numberOfColumns As Integer = 8
Const WS_Reporter_maxIdCell As String = "A1"
Const WS_Reporter_statusCell As String = "E1"

Const WS_Report_firstDataRow As Long = 2
Const WS_Report_firstColumn As Long = 1
Const WS_Report_numberOfColumns As Integer = 7

Sub MOD_Reporter_ActualizePlan()

    On Error GoTo ErrHandler

    Dim planTable As CM_AutoIdTable
    Dim reporterTable As CM_AutoIdTable
    
    Set planTable = New CM_AutoIdTable
    Set reporterTable = New CM_AutoIdTable
    
    If WS_Reporter.Range(WS_Reporter_statusCell) = "Current" Then
        Exit Sub
    End If
    
    planTable.Constructor WS_Plan, WS_Plan_idColumn, WS_Plan_numberOfColumns, WS_Plan_firstDataRow, WS_Plan.Range(WS_Plan_maxIdCell)
    reporterTable.Constructor WS_Reporter, WS_Reporter_idColumn, WS_Reporter_numberOfColumns, WS_Reporter_firstDataRow, WS_Reporter.Range(WS_Reporter_maxIdCell)

    planTable.CopyTo reporterTable, WS_Reporter_maxIdCell
    
    WS_Reporter.Range(WS_Reporter_statusCell) = "Current"
    WS_Planner.Range(WS_Planner_statusCell) = "Got"
    
    WS_Reporter.CBT_ChangeInformation.Visible = True
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub

Sub MOD_Reporter_SendReport()

    On Error GoTo ErrHandler

    If WS_Reporter.Range(WS_Reporter_statusCell) = "Reported" Then
    
        Err.Raise 10000, "MOD_Reporter_SendReport", "Already reported!"
    
    End If

    Dim reporterTable As CM_RecordTable
    Dim reportTable As CM_RecordTable
    
    Set reporterTable = New CM_RecordTable
    Set reportTable = New CM_RecordTable
    
    reporterTable.Constructor WS_Reporter, WS_Reporter_idColumn + 1, WS_Reporter_numberOfColumns - 1, WS_Reporter_firstDataRow
    reportTable.Constructor WS_Report, WS_Report_firstColumn, WS_Report_numberOfColumns, WS_Report_firstDataRow

    reporterTable.CopyTo reportTable
    
    If WS_Reporter.Range(WS_Reporter_statusCell) = "Current" Then
    
        WS_Reporter.Range(WS_Reporter_statusCell) = "Reported"
        WS_Planner.Range(WS_Planner_statusCell) = "Reported"
        
    End If
    
    WS_Reporter.CBT_ChangeInformation.Visible = False
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub

Sub MOD_Reporter_ChangeInformation()

    On Error GoTo ErrHandler:

    UF_ChangeReporterInformation.Show
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub
