Option Compare Database

Private Sub Command44_Click()
On Error GoTo err_Command44_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    DoCmd.Close acForm, "frm_Admin_Packing"
    
Exit_Command44_Click:
    Exit Sub

err_Command44_Click:
    MsgBox Err.Description
    Resume Exit_Command44_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'new for season 2006
'must only allow admins in
On Error GoTo err_open

    Dim permiss
    permiss = GetGeneralPermissions

    If permiss <> "ADMIN" Then
        MsgBox "Only administrators can view this form", vbInformation, "Access Denied"
        DoCmd.Close acForm, "frm_admin_packing"
        
    End If
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
