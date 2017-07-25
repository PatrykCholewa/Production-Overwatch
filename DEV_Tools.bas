Attribute VB_Name = "DEV_Tools"
Option Explicit


Sub DEV_ObjectsSheetProtect()
    WS_Objects.Protect UserInterfaceOnly:=True
    WS_Plan.Protect UserInterfaceOnly:=True
    WS_Planner.Protect UserInterfaceOnly:=True
    WS_Report.Protect UserInterfaceOnly:=True
    WS_Reporter.Protect UserInterfaceOnly:=True
End Sub

Sub DEV_ObjectsSheetUnprotect()
    WS_Objects.Unprotect
    WS_Plan.Unprotect
    WS_Planner.Unprotect
    WS_Report.Unprotect
    WS_Reporter.Unprotect
End Sub

'Sub WorkbookProtect()
'    ActiveWorkbook.Protect UserInterfaceOnly:=True
'End Sub
