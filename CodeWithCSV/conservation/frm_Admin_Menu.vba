Option Compare Database
Option Explicit

Private Sub cboConserv_AfterUpdate()
'new season 2006, if pick conservator name bring up quick view
'list of their records
On Error GoTo err_cboConserv

    DoCmd.OpenForm "frm_ConservationRef_ByConservator", acFormDS, , "[NameID] = " & Me![cboConserv], acFormReadOnly
    

Exit Sub

err_cboConserv:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
'new season 2006
'open conservators list admin screen
On Error GoTo err_Close

    DoCmd.Close acForm, "frm_Admin_Menu", acSaveYes
Exit Sub

err_Close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdConserList_Click()
'new season 2006
'open conservators list admin screen
On Error GoTo err_ConsList

    DoCmd.OpenForm "frm_Admin_conservators", acNormal
Exit Sub

err_ConsList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPacking_Click()
'new season 2006
'open packing list admin screen
On Error GoTo err_packing

    DoCmd.OpenForm "frm_Admin_packing", acNormal
Exit Sub

err_packing:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTreatments_Click()
'new season 2006
'open treatments list admin screen
On Error GoTo err_Treatments

    DoCmd.OpenForm "frm_Admin_treatments", acNormal
Exit Sub

err_Treatments:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'new for season 2006
'must only allow admins in
On Error GoTo err_open

    Dim permiss
    permiss = GetGeneralPermissions

    If permiss <> "ADMIN" Then
        MsgBox "Only administrators can view this form", vbInformation, "Access Denied"
        DoCmd.Close acForm, "frm_admin_menu"
        
    End If
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
