Attribute VB_Name = "MOD_Objects"
Option Explicit

Const FirstDataRow As Long = 5
Const productNameColumn As Long = 1
Const productKitColumn As Long = 2
Const productMaterialColumn As Long = 3
Const kitColumn = 4
Const materialColumn As Long = 5
Const numberOfMaterialColumns As Long = 1
Const numberOfKitColumns As Long = 1
Const numberOfProductColumns As Long = 3

Const WS_Planner_firstDataRow As Long = 4
Const WS_Planner_idColumn As Long = 1
Const WS_Planner_numberOfColumns As Integer = 5
Const WS_Planner_productColumn As Long = 3
Const WS_Planner_kitColumn As Long = 4
Const WS_Planner_materialColumn As Long = 5

Sub MOD_Objects_AddMaterial()

    Dim Value(1 To numberOfMaterialColumns) As String
    Dim table As CM_RecordTable
    Set table = New CM_RecordTable
    
    On Error GoTo ErrHandler:
    Value(1) = InputBox("Add Material")
    If Value(1) = "" Then
        Err.Raise 10000, "InputBox", "There is no argument!"
    End If
    
    table.Constructor WS_Objects, materialColumn, numberOfMaterialColumns, FirstDataRow
    table.AddRecord Value

    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Objects_AddKit()

    Dim Value(1 To numberOfKitColumns) As String
    Dim table As CM_RecordTable
    Set table = New CM_RecordTable
    
    On Error GoTo ErrHandler:
    Value(1) = InputBox("Add Kit")
    If Value(1) = "" Then
        Err.Raise 10000, "InputBox", "There is no argument!"
    End If

    table.Constructor WS_Objects, kitColumn, numberOfKitColumns, FirstDataRow
    table.AddRecord Value
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Objects_AddProduct()
    
    On Error GoTo ErrHandler:

    UF_AddProduct.Show
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
    
End Sub

Sub MOD_Objects_RemoveMaterial()

    Dim name As String
    Dim table As CM_RecordTable
    On Error GoTo ErrHandler:
    
    name = InputBox("Remove Material")
    
    If name = "" Then
        Err.Raise 10000, "MOD_Objects_RemoveMaterial.InputBox", "There is no argument!"
    End If
    
    Set table = New CM_RecordTable
    table.Constructor WS_Objects, productNameColumn, numberOfProductColumns, FirstDataRow
    table.LostObjectAlert name, productMaterialColumn
    
    Set table = New CM_RecordTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow
    table.LostObjectAlert name, WS_Planner_materialColumn

    WS_Objects.Activate
    Set table = New CM_RecordTable
    table.Constructor WS_Objects, materialColumn, numberOfMaterialColumns, FirstDataRow
    table.RemoveRecord name
        
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub

Sub MOD_Objects_RemoveKit()

    Dim name As String
    Dim table As CM_RecordTable
    On Error GoTo ErrHandler:
    
    name = InputBox("Remove Material")
    
    If name = "" Then
        Err.Raise 10000, "MOD_Objects_RemoveKit.InputBox", "There is no argument!"
    End If
    
    Set table = New CM_RecordTable
    table.Constructor WS_Objects, productNameColumn, numberOfProductColumns, FirstDataRow
    table.LostObjectAlert name, productKitColumn
    
    Set table = New CM_RecordTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow
    table.LostObjectAlert name, WS_Planner_kitColumn

    WS_Objects.Activate
    Set table = New CM_RecordTable
    table.Constructor WS_Objects, kitColumn, numberOfKitColumns, FirstDataRow
    table.RemoveRecord name
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub

Sub MOD_Objects_RemoveProduct()

    Dim name As String
    Dim table As CM_RecordTable
    On Error GoTo ErrHandler:
    
    name = InputBox("Remove Product")
    
    If name = "" Then
        Err.Raise 10000, "MOD_Objects_RemoveProduct.InputBox", "There is no argument!"
    End If
    
    Set table = New CM_RecordTable
    table.Constructor WS_Planner, WS_Planner_idColumn, WS_Planner_numberOfColumns, WS_Planner_firstDataRow
    table.LostObjectAlert name, WS_Planner_productColumn

    WS_Objects.Activate
    Set table = New CM_RecordTable
    table.Constructor WS_Objects, productNameColumn, numberOfProductColumns, FirstDataRow
    table.RemoveRecord name
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub
