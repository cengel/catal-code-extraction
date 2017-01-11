Option Compare Database
Option Explicit

Private Sub cmdLocate_Click()
'allow user to locate this artefact in  crate
'this worked for GS but not for figurines as can't seem to alter the visibility of this button on current
'so have put button on main form -  would like to return to this off site and crack it - 9/8/10
On Error GoTo err_locate

    'get find number from main form to pass as openargs
''    Dim current
''    current = Forms![Frm_Basic_Data]![frm_subform_basic].Form![Unit] & ":" & Forms![Frm_Basic_Data]![frm_subform_basic].Form![Lettercode] & ":" & Forms![Frm_Basic_Data]![frm_subform_basic].Form![FindNumber]
''    DoCmd.OpenForm "frm_subform_newlocation", acNormal, , , acFormPropertySettings, acDialog, current
    
Exit Sub

err_locate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMove_Click()
'moving something is more complicated than a simple locate - if the find is in more than one crate I think
'this is an issue that should be raised with the finds officer as is that correct or is that a mistake - so this is the
'first check this code carries out
On Error GoTo err_move

    If Me.RecordsetClone.RecordCount > 1 Then
        MsgBox "This particular find is listed as being located in " & Me.RecordsetClone.RecordCount & " crates. This maybe because " & _
                "it was comprised of more than one material but this needs to be checked. Please take this issue to the Finds Officer " & _
                "straight away who can resolve it and ensure the Crate Register is updated correctly." & Chr(13) & Chr(13) & "The Finds Officer will then deal with " & _
                "updating the location for you.", vbExclamation, "Raise with Finds Officer"
    Else
        'ok only one location so now is it a FG location?
        If Me![CrateLetter] <> "FG" Then
            MsgBox "This particular find is listed as being located in non Figurine Crate. This may mean it has been mis-assigned or " & _
                "that the find was comprised of more than one material. This needs to be checked. Please take this issue to the Finds Officer " & _
                "straight away who can resolve it and ensure the Crate Register is updated correctly." & Chr(13) & Chr(13) & "The Finds Officer will then deal with " & _
                "updating the location for you.", vbExclamation, "Raise with Finds Officer"
        Else
            'ok so allow it to be moved within the FG crates
            'get find number from main form to pass as openargs
            Dim current
            current = Forms![Frm_MainData]![ID number] & ":" & Me![CrateLetter] & Me![CrateNumber]
            DoCmd.OpenForm "frm_subform_changelocation", acNormal, , , acFormPropertySettings, acDialog, current
        End If
    End If

Exit Sub

err_move:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'check if there is any known crate location for the selected record (this form is only opened via frm_sub_basic)
On Error GoTo err_current

    ''MsgBox "current:" & Me.RecordsetClone.RecordCount
    If Me.RecordsetClone.RecordCount > 0 Then
        Me![cmdMove].Visible = True
        Me![cmdLocate].Visible = False
    Else
        Me![cmdMove].Visible = False
        Me![cmdLocate].Visible = True
    End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'check if there is any known crate location for the selected record (this form is only opened via frm_sub_basic)
'On Error GoTo err_open
'
'    MsgBox Me.RecordsetClone.RecordCount
'
'
'Exit Sub
'
'err_open:
'    Call General_Error_Trap
'    Exit Sub
    
End Sub

