VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddPlanerRecord 
   OleObjectBlob   =   "UF_AddPlanerRecord.frx":0000
   Caption         =   "Add Planer Record"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3075
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   8
End
Attribute VB_Name = "UF_AddPlanerRecord"
Attribute VB_Base = "0{FC53C30C-F40E-4A95-B3AB-35D80B822E14}{A926EB3A-0958-4AD4-9D7F-51274497271D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const WS_Objects_firstDataRow As Long = 5
Const WS_Objects_productColumn As Long = 1
Const WS_Objects_numberOfProductColumns As Long = 3

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

Private Sub UserForm_Activate()

    Dim row As Long
    
    With Me.CBX_Product
        .Clear
        .ColumnWidths = "70 pt;70 pt;70 pt" ' WS_Objects_numberOfProductColumns = 3
        .ListWidth = "220 pt"
    End With
    
    row = WS_Objects_firstDataRow
    Do While WS_Objects.Cells(row, WS_Objects_productColumn) <> ""
        Dim i As Integer
        Me.CBX_Product.AddItem WS_Objects.Cells(row, WS_Objects_productColumn).Value
        For i = 1 To WS_Objects_numberOfProductColumns - 1
            Me.CBX_Product.List(row - WS_Objects_firstDataRow, i) = WS_Objects.Cells(row, WS_Objects_productColumn + i)
        Next i
        row = row + 1
    Loop
    
 
End Sub

Private Sub CBT_AddRecord_Click()
    On Error GoTo ErrHandler
    
    Dim UF_APR_table As CM_AutoIdTable
    Dim UF_APR_productTable As CM_RecordTable
    Dim UF_APR_values(1 To WS_Planner_numberOfColumns) As String
    Dim UF_APR_productRow As Long
    
    Set UF_APR_table = New CM_AutoIdTable
    Set UF_APR_productTable = New CM_RecordTable
    
    If Not IsNumeric(Me.TBX_Quantity.Value) Then
        Err.Raise 10000, "UF_AddPlanerRecord.CBT_AddRecord_Click.CBT_AddRecord", "Quantity is not numeric."
    End If
    
    UF_APR_productTable.Constructor WS_Objects, WS_Objects_productColumn, WS_Objects_numberOfProductColumns, WS_Objects_firstDataRow
    UF_APR_productRow = UF_APR_productTable.GetRowIndexOf(Me.CBX_Product.Value, WS_Objects_productColumn)
    
    UF_APR_values(1) = Me.TBX_Quantity.Value
    UF_APR_values(2) = Me.CBX_Product.Value
    UF_APR_values(3) = WS_Objects.Cells(UF_APR_productRow, WS_Objects_productColumn + 1)
    UF_APR_values(4) = WS_Objects.Cells(UF_APR_productRow, WS_Objects_productColumn + 2)
    
    UF_APR_table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow, _
                        WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn)
    UF_APR_table.AddRecord UF_APR_values
    
    WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn) = WS_Planner.Cells(WS_Planner_maxIdRow, WS_Planner_maxIdColumn) + 1
    
    WS_Planner.Range(WS_Planner_statusCell) = "Not sent"
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub
