Attribute VB_Name = "MOD_Start"
Option Explicit

Const loginCell As String = "D2"
Const functionCell As String = "H2"

Sub MOD_Start_Logout()

    With WS_Start
        .CBT_Archives.Visible = False
        .CBT_LogIn.Visible = True
        .CBT_Logout.Visible = False
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = False
        .CBT_Planner.Visible = False
        .CBT_Reporter.Visible = False
        .CBT_Users.Visible = False
        .CBT_ProdCharter.Visible = False
        .CBT_HideProdCharter.Visible = False
        
        .Range(loginCell) = ""
        .Range(functionCell) = ""
    End With

    DEV_Tools.DEV_HideEverything

End Sub

Sub MOD_Start_Login()

    UF_LogIn.Show

End Sub

Sub MOD_Start_WSObjects()
    WS_Objects.Visible = xlSheetVisible
    WS_Objects.Activate
End Sub

Sub MOD_Start_WSPlanner()
    WS_Planner.Visible = xlSheetVisible
    WS_Planner.Activate
End Sub

Sub MOD_Start_WSPlan()
    WS_Plan.Visible = xlSheetVisible
    WS_Plan.Activate
End Sub

Sub MOD_Start_WSReporter()
    WS_Reporter.Visible = xlSheetVisible
    WS_Reporter.Activate
End Sub

Sub MOD_Start_WSReport()
    WS_Report.Visible = xlSheetVisible
    WS_Report.Activate
End Sub

Sub MOD_Start_WSUser()
    WS_User.Visible = xlSheetVisible
    WS_User.Activate
End Sub

Sub MOD_Start_ChangePass()

    UF_UserSetting.ChangePassword WS_Start.Range(loginCell)

End Sub

Sub MOD_Start_WSArchives()
    WS_Archives.Visible = xlSheetVisible
    WS_Archives.Activate
End Sub

Sub MOD_Start_ProdCharter()

    CHT_Production.Visible = xlSheetVisible

    With WS_Start
        .CBT_HideProdCharter.Visible = True
        .CBT_ProdCharter.Visible = False
    End With
    
    UF_ProdChart.Show
    
End Sub

Sub MOD_Start_HideProdCharter()

    CHT_Production.Visible = xlSheetHidden

    With WS_Start
        .CBT_HideProdCharter.Visible = False
        .CBT_ProdCharter.Visible = True
    End With
    
    UF_ProdChart.Hide

End Sub

Sub MOD_Start_Admin()

    With WS_Start
        .CBT_Archives.Visible = True
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_Objects.Visible = True
        .CBT_PassChange.Visible = True
        .CBT_Planner.Visible = True
        .CBT_Reporter.Visible = True
        .CBT_Users.Visible = True
        .CBT_HideProdCharter.Visible = False
        .CBT_ProdCharter.Visible = True
    End With

End Sub

Sub MOD_Start_Worker()

    With WS_Start
        .CBT_Archives.Visible = False
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = True
        .CBT_Planner.Visible = False
        .CBT_Reporter.Visible = True
        .CBT_Users.Visible = False
        .CBT_HideProdCharter.Visible = False
        .CBT_ProdCharter.Visible = False
    End With

End Sub

Sub MOD_Start_Manager()

    With WS_Start
        .CBT_Archives.Visible = True
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = True
        .CBT_Planner.Visible = True
        .CBT_Reporter.Visible = False
        .CBT_Users.Visible = False
        .CBT_HideProdCharter.Visible = False
        .CBT_ProdCharter.Visible = True
    End With

End Sub
