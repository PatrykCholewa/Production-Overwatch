VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddProduct 
   OleObjectBlob   =   "UF_AddProduct.frx":0000
   Caption         =   "Add Product"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3510
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   30
End
Attribute VB_Name = "UF_AddProduct"
Attribute VB_Base = "0{3F7C7A68-D609-4ECE-A289-565D534B2CAA}{ABD93449-2FB0-4978-8672-A994C9723700}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const FirstDataRow As Long = 5
Const productNameColumn As Long = 1
Const kitNameColumn As Long = 4
Const materialNameColumn As Long = 5
Const numberOfProductColumns As Long = 3


Private Sub CBX_Kit_Change()

End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_Activate()

    Dim row As Long
    
    Me.CBX_Kit.Clear
    Me.CBX_Material.Clear
    
    row = FirstDataRow
    Do While WS_Objects.Cells(row, kitNameColumn) <> ""
        Me.CBX_Kit.AddItem Cells(row, kitNameColumn).Value
        row = row + 1
    Loop
    
    row = FirstDataRow
    Do While WS_Objects.Cells(row, materialNameColumn) <> ""
        Me.CBX_Material.AddItem Cells(row, materialNameColumn).Value
        row = row + 1
    Loop
    
End Sub

Private Sub CBT_AddProduct_Click()
    
    On Error GoTo ErrHandler

    Dim table As CM_RecordTable
    Dim values(1 To numberOfProductColumns) As String
    Dim i As Integer
    Set table = New CM_RecordTable
    values(1) = TBX_Name.text
    values(2) = CBX_Kit.Value
    values(3) = CBX_Material.Value
    table.Constructor WS_Objects, productNameColumn, numberOfProductColumns, FirstDataRow
    
    For i = 1 To numberOfProductColumns
        If values(i) = "" Then
            Err.Raise 10000, "UF_AddProduct.CBT_AddProduct_Click()", "There is ungiven value."
        End If
    Next i
    
    table.AddRecord values
    
    Exit Sub
    
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext

End Sub

Private Sub UserForm_Click()

End Sub
