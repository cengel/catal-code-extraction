Option Compare Database
Option Explicit

Private Sub cmdAdmin_Click()
'open the pub data screen
On Error GoTo err_Admin

    DoCmd.OpenForm "Frm_GS_AdminMenu", acNormal
    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_Admin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOpenData_Click()
'open the main data screen
On Error GoTo err_OpenData

    DoCmd.OpenForm "Frm_Level1", acNormal
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

Private Sub cmdOpenData2_Click()
'open the main data screen
On Error GoTo err_OpenData2

    DoCmd.OpenForm "Frm_Level2", acNormal
    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_OpenData2:
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

Private Sub cmdSample_Click()
'open the sample data screen
On Error GoTo err_Sample

    DoCmd.OpenForm "Frm_GS_Samples", acNormal

Exit Sub

err_Sample:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    'Dim permiss
    'permiss = GetGeneralPermissions
    'If permiss <> "ADMIN" Then
    '    Me![cmdAdmin].Visible = False
    'Else
    '   Me![cmdAdmin].Visible = True
    'End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub openOldDB_Click()

On Error GoTo err_openOldDB

    DoCmd.OpenForm "Frm_Basic_Data", acNormal
    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_openOldDB:
    Call General_Error_Trap
    Exit Sub
End Sub
