VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WS_Start"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CBT_Login, 2, 0, MSForms, CommandButton"
Attribute VB_Control = "CBT_Logout, 3, 1, MSForms, CommandButton"
Attribute VB_Control = "CBT_Objects, 4, 2, MSForms, CommandButton"
Attribute VB_Control = "CBT_Planner, 5, 3, MSForms, CommandButton"
Attribute VB_Control = "CBT_Plan, 6, 4, MSForms, CommandButton"
Attribute VB_Control = "CBT_Users, 7, 5, MSForms, CommandButton"
Attribute VB_Control = "CBT_PassChange, 8, 6, MSForms, CommandButton"
Attribute VB_Control = "CBT_Reporter, 9, 7, MSForms, CommandButton"
Attribute VB_Control = "CBT_NewReport, 10, 8, MSForms, CommandButton"
Option Explicit

Private Sub CBT_Logout_Click()

    MOD_Start_Logout

End Sub

Private Sub CBT_PassChange_Click()

End Sub

Private Sub CBT_Login_Click()

    MOD_Start.MOD_Start_Login

End Sub
