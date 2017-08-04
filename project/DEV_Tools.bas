Attribute VB_Name = "DEV_Tools"
Option Explicit

Sub DEV_ProtectEverything()
    WS_Archives.Protect UserInterfaceOnly:=True
    WS_Objects.Protect UserInterfaceOnly:=True
    WS_Plan.Protect UserInterfaceOnly:=True
    WS_Planner.Protect UserInterfaceOnly:=True
    WS_Report.Protect UserInterfaceOnly:=True
    WS_Reporter.Protect UserInterfaceOnly:=True
    WS_User.Protect UserInterfaceOnly:=True
    WS_Pass.Protect UserInterfaceOnly:=True
    WS_Start.Protect UserInterfaceOnly:=True
End Sub

Sub DEV_HideEverything()
    CHT_Production.Visible = xlSheetVeryHidden
    WS_Archives.Visible = xlSheetVeryHidden
    WS_Objects.Visible = xlSheetVeryHidden
    WS_Plan.Visible = xlSheetVeryHidden
    WS_Planner.Visible = xlSheetVeryHidden
    WS_Report.Visible = xlSheetVeryHidden
    WS_Reporter.Visible = xlSheetVeryHidden
    WS_User.Visible = xlSheetVeryHidden
    WS_Pass.Visible = xlSheetVeryHidden
    WS_Start.Visible = xlSheetVisible
End Sub

Sub DEV_UnprotectEverything()
    WS_Archives.Unprotect
    WS_Objects.Unprotect
    WS_Plan.Unprotect
    WS_Planner.Unprotect
    WS_Report.Unprotect
    WS_Reporter.Unprotect
    WS_User.Unprotect
    WS_Start.Unprotect
End Sub

'Sub WorkbookProtect()
'    ActiveWorkbook.Protect UserInterfaceOnly:=True
'End Sub
