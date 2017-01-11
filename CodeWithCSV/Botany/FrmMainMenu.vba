Option Compare Database
Option Explicit



Private Sub cmdBasicData_Click()
' Open Basic data form
On Error GoTo err_Basic

    DoCmd.OpenForm "FrmBasicData", acNormal, , , acFormEdit
Exit Sub

err_Basic:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdHistoric_Click()
' Open Basic data form
On Error GoTo err_Historic

    DoCmd.OpenForm "Bots98: Main", acNormal, , , acFormReadOnly
Exit Sub

err_Historic:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPriorityReport_Click()
' Open Basic data form
On Error GoTo err_PR

    DoCmd.OpenForm "FrmPriorityReport", acNormal, , , acFormEdit
Exit Sub

err_PR:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
' Quit app
On Error GoTo err_Quit
    DoCmd.Quit acQuitSaveAll

Exit Sub

err_Quit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdScan_Click()
' open scanning form
On Error GoTo err_Scan

    DoCmd.OpenForm "FrmSampleScan", acNormal, , , acFormEdit
Exit Sub

err_Scan:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSieveScan_Click()
' open scanning form
On Error GoTo err_Scan

    DoCmd.OpenForm "FrmSieveScan", acNormal, , , acFormEdit
Exit Sub

err_Scan:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnit_Click()
' new on site 2006
On Error GoTo err_Unit

    DoCmd.OpenForm "FrmBotUnitDescription", acNormal, , , acFormEdit
Exit Sub

err_Unit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command3_Click()
' open priority form
On Error GoTo err_Priority

    DoCmd.OpenForm "FrmPriority", acNormal, , , acFormEdit
Exit Sub

err_Priority:
    Call General_Error_Trap
    Exit Sub
End Sub
