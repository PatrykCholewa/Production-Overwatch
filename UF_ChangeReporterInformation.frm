VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ChangeReporterInformation 
   OleObjectBlob   =   "UF_ChangeReporterInformation.frx":0000
   Caption         =   "Change Information"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   57
End
Attribute VB_Name = "UF_ChangeReporterInformation"
Attribute VB_Base = "0{8395AB63-8B88-40D6-A2C0-E7F8185D6E8B}{C367BA50-B097-4AF9-936A-4BCB9E011C7E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const WS_Reporter_firstDataRow As Long = 4
Const WS_Reporter_idColumn As Long = 1
Const WS_Reporter_requiredAmountColumn As Long = 2
Const WS_Reporter_producedAmountColumn As Long = 3
Const WS_Reporter_nokColumn As Long = 4
Const WS_Reporter_timeColumn As Long = 5
Const WS_Reporter_productColumn As Long = 6
Const WS_Reporter_kitColumn As Long = 7
Const WS_Reporter_materialColumn As Long = 8
Const WS_Reporter_numberOfColumns As Integer = 8
Const WS_Reporter_maxIdCell As String = "C1"
Const WS_Reporter_statusCell As String = "E1"

Private Sub CBT_Apply_Click()

    On Error GoTo ErrHandler

    Dim row As Long
    
    row = WS_Reporter_firstDataRow + CBX_RProduct.ListIndex
    
    If Not IsNumeric(TBX_Quantity.text) Then
        Err.Raise 10000, "UF_ChangeReporterInformation.CBT_Apply_Click()", "Produced quantity is not numeric!"
    End If
    
    If Not IsNumeric(TBX_Noks.text) Then
        Err.Raise 10000, "UF_ChangeReporterInformation.CBT_Apply_Click()", "NOK is not numeric!"
    End If
    
    WS_Reporter.Cells(row, WS_Reporter_producedAmountColumn) = Me.TBX_Quantity.text
    WS_Reporter.Cells(row, WS_Reporter_nokColumn) = Me.TBX_Noks.text
    WS_Reporter.Cells(row, WS_Reporter_timeColumn) = Format(Me.TBX_Time.text, "HH:MM")
    
    Exit Sub
        
ErrHandler:
    LIB_ErrHandler.Handle Err.number, Err.source, Err.description, Err.helpFile, Err.helpContext
End Sub

Private Sub CBX_RProduct_Change()

    Dim row As Long

    TBX_RQuantity.text = Me.CBX_RProduct.List(CBX_RProduct.ListIndex, 1)
    
    row = WS_Reporter_firstDataRow + CBX_RProduct.ListIndex
    
    Me.TBX_Quantity.text = WS_Reporter.Cells(row, WS_Reporter_producedAmountColumn)
    Me.TBX_Noks.text = WS_Reporter.Cells(row, WS_Reporter_nokColumn)
    Me.TBX_Time.text = Format(WS_Reporter.Cells(row, WS_Reporter_timeColumn), "HH:MM")

End Sub

Private Sub TBX_Quantity_Change()

End Sub

Private Sub TBX_Time_Change()

End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_Activate()

    Dim row As Long
    
    With Me.CBX_RProduct
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "60 pt;25 pt;35 pt;35 pt" ' WS_Reporter architecture
        .ListWidth = "165 pt"
    End With
    
    Me.TBX_Time.Value = Format(Me.TBX_Time.Value, "HH:MM")
    
    row = WS_Reporter_firstDataRow
    Do While WS_Reporter.Cells(row, WS_Reporter_idColumn) <> ""
        Dim i As Integer
        Me.CBX_RProduct.AddItem WS_Reporter.Cells(row, WS_Reporter_productColumn)
        For i = 1 To 3 ' 3 because quantity, kit and material
            Me.CBX_RProduct.List(row - WS_Reporter_firstDataRow, 1) = WS_Reporter.Cells(row, WS_Reporter_requiredAmountColumn)
            Me.CBX_RProduct.List(row - WS_Reporter_firstDataRow, 2) = WS_Reporter.Cells(row, WS_Reporter_kitColumn)
            Me.CBX_RProduct.List(row - WS_Reporter_firstDataRow, 3) = WS_Reporter.Cells(row, WS_Reporter_materialColumn)
        Next i
        row = row + 1
    Loop
 
End Sub

Private Sub UserForm_Click()

End Sub
