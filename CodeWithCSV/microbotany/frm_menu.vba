Option Compare Database
Option Explicit

Private Sub cmd_Register_Click()
On Error GoTo err_Register_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Phyto_SampleRegister"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit Sub

err_Register_Click:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub

End Sub

Private Sub cmdAdmin_Click()
On Error GoTo err_cmdAdmin
    DoCmd.OpenForm "frm_admin_phytoname_LOV", acNormal
Exit Sub

err_cmdAdmin:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Quit acQuitSaveAll
Exit Sub

err_close:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdData_Click()
On Error GoTo err_data

    DoCmd.OpenForm "frm_Phyto_Data_Entry", acNormal
Exit Sub

err_data:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdField_Click()
On Error GoTo err_field

    DoCmd.OpenForm "frm_Phyto_FieldAnalysis", acNormal
Exit Sub

err_field:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub
