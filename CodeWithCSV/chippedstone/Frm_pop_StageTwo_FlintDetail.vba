Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    If Me.FilterOn = True Then Me.FilterOn = False
    Me![GID].Enabled = True
    DoCmd.GoToControl Me![GID].Name
    DoCmd.FindRecord Me![cboFind]
    DoCmd.GoToControl Me![Weight].Name
    Me![GID].Enabled = False
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
cmdClose_Click

End Sub

Private Sub cmdClose_Click()
'close this new 2009 form
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
On Error GoTo Err_gofirst_Click


    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_gofirst_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdGoLast_Click()
On Error GoTo Err_goLast_Click


    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_goLast_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdGoNext_Click()
On Error GoTo Err_goNext_Click


    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_goNext_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdGoPrev_Click()
On Error GoTo Err_goPrev_Click


    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_goPrev_Click:
    Call General_Error_Trap
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open

DoCmd.GoToControl "Weight"

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
