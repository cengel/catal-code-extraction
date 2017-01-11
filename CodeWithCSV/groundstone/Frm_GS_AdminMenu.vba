Option Compare Database
Option Explicit

Private Sub cmdOpenData_Click()
'open the main data screen
On Error GoTo err_OpenData

    DoCmd.OpenForm "Frm_GS_Main", acNormal

Exit Sub

err_OpenData:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdArteClass_Click()
On Error GoTo err_cmdArteClass

    DoCmd.OpenForm "Frm_Admin_ArtefactClass", acNormal

Exit Sub

err_cmdArteClass:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdArteType_Click()
On Error GoTo err_cmdArteType

    DoCmd.OpenForm "Frm_Admin_ArtefactTypeSubTypeLOV", acNormal

Exit Sub

err_cmdArteType:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdFraction_Click()
On Error GoTo err_cmdPercent

    DoCmd.OpenForm "Frm_Admin_FractionLOV", acNormal

Exit Sub

err_cmdPercent:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdPercent_Click()

On Error GoTo err_cmdPercent

    DoCmd.OpenForm "Frm_Admin_FlotPercent", acNormal

Exit Sub

err_cmdPercent:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
'quit system
On Error GoTo err_Quit

    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_Quit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
