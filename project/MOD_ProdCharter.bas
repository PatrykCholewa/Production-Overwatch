Attribute VB_Name = "MOD_ProdCharter"
Option Explicit

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

Public Sub MOD_ProdCharter_HideWS()

    CHT_Production.Visible = xlSheetHidden

End Sub

Private Function SubjectColumn(name As String) As Long

    Select Case name
        Case "Product"
            SubjectColumn = WS_Archives_productColumn
        Case "Kit"
            SubjectColumn = WS_Archives_kitColumn
        Case "Material"
            SubjectColumn = WS_Archives_materialColumn
        Case Else
            Err.Raise 10001, "MOD_ProdCharter.GetSubjectColumn", "Value not permitted!"
    End Select

End Function

Private Function ObjectColumn(name As String) As Long

    Select Case name
        Case "NOK"
            ObjectColumn = WS_Archives_nokColumn
        Case "Time"
            ObjectColumn = WS_Archives_timeColumn
        Case Else
            Err.Raise 10001, "MOD_ProdCharter.GetSubjectColumn", "Value not permitted!"
    End Select

End Function

Private Function AnalysisTypeNumber(name As String) As Integer

    Select Case name
        Case "Average"
            'Not completed
        Case "Progress"
            'Not completed
        Case Else
            Err.Raise 10001, "MOD_ProdCharter.GetSubjectColumn", "Value not permitted!"
    End Select
    
End Function

Private Function MinDate(minDateString As String) As Date

    Select Case s
        Case "ALL"
            'Not completed
        Case "Last month"
            'Not completed
        Case Else
            Err.Raise 10001, "MOD_ProdCharter.GetSubjectColumn", "Value not permitted!"
    End Select

End Function

Public Sub MOD_ProdCharter_DrawChart(minDateString As String, subjectCategory As String, subjectName As String, objectName As String, analysisType As String)

    

End Sub
    

  
