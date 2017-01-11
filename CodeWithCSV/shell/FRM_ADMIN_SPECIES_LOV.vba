Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
'find combo
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then

    DoCmd.GoToControl Me![FRM_SUB_ADMIN_Species].Name
    DoCmd.GoToControl "type number"
    DoCmd.FindRecord Me![cboFind]
    Me![cboFind] = ""
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "This type number has not been entered. Add it to the bottom of the list and next time you open this form it will be placed in the correct position", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
    DoCmd.GoToControl Me![FRM_SUB_ADMIN_Species].Name
    'DoCmd.GoToRecord acDataForm, Me![FRM_SUB_ADMIN_Species], acLast
    
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close

    DoCmd.OpenForm "Frm_menu"
    DoCmd.Restore
    
    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'maximise
On Error GoTo err_open

    DoCmd.Maximize

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
