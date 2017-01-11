Option Compare Database
Option Explicit

Private Sub cmdAdmin_Click()
'open the pub data screen
On Error GoTo err_Admin

    DoCmd.OpenForm "Frm_GS_AdminMenu", acNormal

Exit Sub

err_Admin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOpenData_Click()
'open the main data screen
On Error GoTo err_OpenData

    DoCmd.OpenForm "Frm_MainData", acNormal, , , acFormPropertySettings
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_OpenData:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPub_Click()
'open the pub data screen
On Error GoTo err_pub

    DoCmd.OpenForm "Frm_GS_Publications", acNormal

Exit Sub

err_pub:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
'quit system
On Error GoTo err_Quit

    DoCmd.Quit acQuitSaveAll
    

Exit Sub

err_Quit:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdStageTwo_Click()
'open the sample data screen
On Error GoTo err_Sample

    DoCmd.OpenForm "Frm_Search", acNormal, , , acFormPropertySettings
      DoCmd.Close acForm, Me.Name

Exit Sub

err_Sample:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
'    permiss = GetGeneralPermissions
'    If permiss <> "ADMIN" Then
'        me![cmdAdmin].visible = true
'        DoCmd.close acForm, Me.Name
'    else
'       me![cmdAdmin].visible = false
'    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
