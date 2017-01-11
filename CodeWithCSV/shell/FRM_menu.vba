Option Compare Database
Option Explicit

Private Sub cmdAdmin_Click()
On Error GoTo err_cmdAdmin
    DoCmd.OpenForm "FRM_ADMIN_SPECIES_LOV", acNormal
Exit Sub

err_cmdAdmin:
    MsgBox "An error has occured: " & Err.DESCRIPTION
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Quit acQuitSaveAll
Exit Sub

err_close:
    MsgBox "An error has occured: " & Err.DESCRIPTION
    Exit Sub
End Sub

Private Sub cmdData_Click()
On Error GoTo err_data

    DoCmd.OpenForm "FRM_SHELL_LEVEL_ONE", acNormal
Exit Sub

err_data:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cmdField_Click()
'open data entry form
On Error GoTo err_cmdField

    DoCmd.OpenForm "FRM_SHELL_LEVEL_ONE"
    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdField:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdTypeLetter_Click()
On Error GoTo err_tl

    DoCmd.OpenForm "FRM_ADMIN_Type_Letter", acNormal
Exit Sub

err_tl:
    Call General_Error_Trap
    Exit Sub
End Sub
