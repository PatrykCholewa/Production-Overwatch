Attribute VB_Name = "MOD_Start"
Option Explicit

Const loginCell As String = "D2"
Const functionCell As String = "H2"

Sub MOD_Start_Logout()

    With WS_Start
        .CBT_LogIn.Visible = True
        .CBT_Logout.Visible = False
        .CBT_NewReport.Visible = False
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = False
        .CBT_Plan.Visible = False
        .CBT_Planner.Visible = False
        .CBT_Reporter.Visible = False
        .CBT_Users.Visible = False
        
        .Range(loginCell) = ""
        .Range(functionCell) = ""
    End With

End Sub

Sub MOD_Start_Login()

    UF_LogIn.Show

End Sub

Sub MOD_Start_Admin()

    With WS_Start
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_NewReport.Visible = True
        .CBT_Objects.Visible = True
        .CBT_PassChange.Visible = True
        .CBT_Plan.Visible = True
        .CBT_Planner.Visible = True
        .CBT_Reporter.Visible = True
        .CBT_Users.Visible = True
    End With

End Sub

Sub MOD_Start_Worker()

    With WS_Start
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_NewReport.Visible = False
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = True
        .CBT_Plan.Visible = True
        .CBT_Planner.Visible = False
        .CBT_Reporter.Visible = True
        .CBT_Users.Visible = False
    End With

End Sub

Sub MOD_Start_Manager()

    With WS_Start
        .CBT_LogIn.Visible = False
        .CBT_Logout.Visible = True
        .CBT_NewReport.Visible = True
        .CBT_Objects.Visible = False
        .CBT_PassChange.Visible = True
        .CBT_Plan.Visible = True
        .CBT_Planner.Visible = True
        .CBT_Reporter.Visible = False
        .CBT_Users.Visible = False
    End With

End Sub
