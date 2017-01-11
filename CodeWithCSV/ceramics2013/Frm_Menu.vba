Option Compare Database
Option Explicit

Private Sub cmdAdmin_Click()
On Error GoTo err_cmdAdmin

    DoCmd.OpenForm "Frm_AdminMenu"
    
    
Exit Sub

err_cmdAdmin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdBody_Click()
On Error GoTo err_cmdBody

    DoCmd.OpenForm "Frm_BodySherd"
    DoCmd.Maximize
    
Exit Sub

err_cmdBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDiagnostic_Click()
On Error GoTo err_cmdDiag

    DoCmd.OpenForm "Frm_Diagnostic"
    DoCmd.Maximize
Exit Sub

err_cmdDiag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOldSys_Click()
DoCmd.OpenForm "Frm_Main"

End Sub

Private Sub cmdQuit_Click()
On Error GoTo err_quit

    DoCmd.Quit acQuitSaveAll
    
Exit Sub

err_quit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnIdBody_Click()
On Error GoTo err_cmdUnIdDiag

    DoCmd.OpenForm "Frm_NonNeolithic_Sherds"
    DoCmd.Maximize
Exit Sub

err_cmdUnIdDiag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnIdDiag_Click()
On Error GoTo err_cmdUnIdDiag

    DoCmd.OpenForm "Frm_Unidentified_Diagnostic"
    DoCmd.Maximize
    
Exit Sub

err_cmdUnIdDiag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnitsOverview_Click()
On Error GoTo err_cmdUnit

    DoCmd.OpenForm "Frm_UnitOverview"
    DoCmd.Maximize
Exit Sub

err_cmdUnit:
    Call General_Error_Trap
    Exit Sub
End Sub
