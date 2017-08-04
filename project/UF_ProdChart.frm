VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ProdChart 
   OleObjectBlob   =   "UF_ProdChart.frx":0000
   Caption         =   "Production Charter"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   25
End
Attribute VB_Name = "UF_ProdChart"
Attribute VB_Base = "0{15BF64C7-25CC-44DC-A092-738DE2960B6C}{C453BA68-EDDF-4238-817E-E568E489B3BE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const WS_Objects_firstDataRow As Long = 5
Const WS_Objects_productColumn As Long = 1
Const WS_Objects_kitColumn = 4
Const WS_Objects_materialColumn As Long = 5

Const WS_Archives_firstDataRow As Long = 4
Const WS_Archives_idColumn As Long = 1
Const WS_Archives_reqQuantityColumn As Long = 2
Const WS_Archives_prodQuantityColumn As Long = 3
Const WS_Archives_nokColumn As Long = 4
Const WS_Archives_timeColumn As Long = 5
Const WS_Archives_productColumn As Long = 6
Const WS_Archives_kitColumn As Long = 7
Const WS_Archives_materialColumn As Long = 8
Const WS_Archives_dateColumn As Long = 9
Const WS_Archives_numberOfColumns As Integer = 9

Private Sub CBX_ChtType_Change()

    Select Case Me.CBX_ChtType.value
        Case "Bar"
            CHT_Production.ChartType = xlBarClustered
        Case "Line"
            CHT_Production.ChartType = xlLine
        Case Else
            Err.Raise 10001, "UF_ProdCharter.CBX_Subject_Change", "Not planned value!"
    End Select
    
    CHT_Production.Activate

End Sub

Private Sub CBX_Subject_Change()

    Dim row As Long
    Dim column As Long

    Select Case Me.CBX_Subject.value
        Case "Product"
            column = WS_Objects_productColumn
        Case "Kit"
            column = WS_Objects_kitColumn
        Case "Material"
            column = WS_Objects_materialColumn
        Case Else
            Err.Raise 10001, "UF_ProdCharter.CBX_Subject_Change", "Not planned value!"
    End Select
    
    Me.CBX_ConSubject.Clear
    Me.CBX_ConSubject.AddItem "ALL"
    Me.CBX_ConSubject.value = "ALL"
    Me.CBX_ConSubject.Enabled = True
    
    row = WS_Objects_firstDataRow
    Do While WS_Objects.Cells(row, column) <> ""
        Me.CBX_ConSubject.AddItem WS_Objects.Cells(row, column)
        row = row + 1
    Loop
                    

End Sub

Private Sub UserForm_Activate()
 
    CHT_Production.Activate
 
    Me.CBX_ChtType.AddItem "Bar"
    Me.CBX_ChtType.AddItem "Line"
    
    Me.CBX_Date.AddItem "ALL"
    Me.CBX_Date.AddItem "Last month"
    Me.CBX_Date.value = "ALL"
    
    Me.CBX_Subject.AddItem "Product"
    Me.CBX_Subject.AddItem "Kit"
    Me.CBX_Subject.AddItem "Material"
    Me.CBX_Subject.value = "Product"
    
    Me.CBX_ConSubject.AddItem "ALL"
    Me.CBX_ConSubject.value = "ALL"
    Me.CBX_ConSubject.Enabled = False
    
    Me.CBX_Object.AddItem "NOK"
    Me.CBX_Object.AddItem "Time"
    Me.CBX_Object.value = "Time"
    
    Me.CBX_What.AddItem "Average"
    Me.CBX_What.AddItem "Progress"
    Me.CBX_What.value = "Progress"
 
 
End Sub

Private Sub CBT_Draw_Click()

    MOD_ProdCharter.MOD_ProdCharter_DrawChart Me.CBX_Date, _
                                                Me.CBX_Subject, _
                                                Me.CBX_ConSubject, _
                                                Me.CBX_Object, _
                                                Me.CBX_What
                                            

End Sub
