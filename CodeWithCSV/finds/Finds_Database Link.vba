Option Compare Database
Option Explicit 'saj


Private Sub Update_GID()
'sub used by gid fields written by anja adapted by saj to error trap and include letter code fld
On Error GoTo err_updategid

'Me![GID] = Me![Unit] & "." & Me![Find Number]

Me![GID] = Me![txtUnit] & "." & Me![cboFindLetter] & Me![txtFindNumber]
If Me![txtUnit] <> "" And Me![cboFindLetter] <> "" And Me![txtFindNumber] <> "" Then
    Me.Refresh
End If
Exit Sub

err_updategid:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected gid from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'update subform
         If Me![cboFindUnit] = "Archaeobots" Then
            Me![frm_subform_dblink].SourceObject = "frm_subform_bots"
         ElseIf Me![cboFindUnit] = "Chipped Stone" Then
             Me![frm_subform_dblink].SourceObject = "frm_subform_chippedstone"
         ElseIf Me![cboFindUnit] = "Faunal" Then
             Me![frm_subform_dblink].SourceObject = "frm_subform_faunal"
         ElseIf Me![cboFindUnit] = "Groundstone" Then
             Me![frm_subform_dblink].SourceObject = "frm_subform_Groundstone"
         ElseIf Me![cboFindUnit] = "Excavation X finds" Then
             Me![frm_subform_dblink].SourceObject = "frm_subform_ExcaXFinds"
         
         End If
    End If
    
    If Me![cboUnit] <> "" Then
        Call cboUnit_AfterUpdate
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUnit_AfterUpdate()
On Error GoTo err_cboUnit

    If Me![cboUnit] <> "" Then
        DoCmd.GoToControl "frm_subform_dblink"
        Me![frm_subform_dblink].Form.FilterOn = True
        If Me![cboFindUnit] = "Faunal" Or Me![cboFindUnit] = "Archaeobots" Or Me![cboFindUnit] = "Excavation X finds" Then
            
            Me![frm_subform_dblink].Form.Filter = "[unit number] =" & Me![cboUnit]
        
        Else
            Me![frm_subform_dblink].Form.Filter = "[unit] =" & Me![cboUnit]
        End If
        
        
    End If


Exit Sub

err_cboUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub













Private Sub cmdAll_Click()
On Error GoTo err_cmdAll

    DoCmd.GoToControl "frm_subform_dblink"
    Me![frm_subform_dblink].Form.FilterOn = False
    Me![cboUnit] = ""
Exit Sub

err_cmdAll:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open

    Me![frm_subform_dblink].SourceObject = ""
    Me![cboUnit] = ""
    Me![cboFindUnit] = ""
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
