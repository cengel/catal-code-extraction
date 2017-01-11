Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo err_cancel

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cancel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOpen_Click()
On Error GoTo err_open

    If Me![frmWhich] = 1 Then
        DoCmd.OpenForm "Frm_MNI", acNormal, , , , , 1
        DoCmd.Close acForm, Me.Name
        
    ElseIf Me![frmWhich] = 2 Then
        DoCmd.OpenForm "Frm_MNI", acNormal, , , , , 2
        DoCmd.Close acForm, Me.Name
    ElseIf Me![frmWhich] = 3 Then
        DoCmd.OpenForm "Frm_MNI", acNormal, , , , , 3
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "No MNI option selected", vbInformation, "Choose MNI"
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub
