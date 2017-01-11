Option Compare Database
Option Explicit


Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
'        Me![txtUnit] = Me![cboFindUnit]
        Me![frm_sub_X-finds].Requery
        Me![fmr_sub_Exca_Samples].Requery
        Me![frm_subform_UnitDetails].Requery
        Me![Frm_Sub_Diagnostic_Totals].Requery
        Me![Frm_Sub_Unidentified_Sherds_Totals].Requery
        Me![Frm_Sub_BodySherds_Totals].Requery
        Me![Frm_Sub_NonNeolithics_Totals].Requery
 
        'if a filter is on - turn off
        If Me.FilterOn = True Then Me.FilterOn = False
        'for existing number the field will be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![txtUnit].Enabled = False Then Me![txtUnit].Enabled = True
        DoCmd.GoToControl "txtUnit"
        DoCmd.FindRecord Me![cboFindUnit], , , , True
        Me![cboFindUnit] = ""
        DoCmd.GoToControl "cboFindUnit"
        Me![txtUnit].Enabled = False
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboViewBodySherds_AfterUpdate()
On Error GoTo err_cboViewBodysherds
    Dim where
    where = "[Unit] = " & Me![cboViewBodySherds].Column(1) & " AND [Ware Code] = '" & Me![cboViewBodySherds].Column(2) & "'"
    
    DoCmd.OpenForm "Frm_BodySherd", acNormal, , where
    Me![cboViewBodySherds] = ""
Exit Sub

err_cboViewBodysherds:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboViewBody_AfterUpdate()
On Error GoTo err_cboViewBody
    Dim where
    where = "[Unit] = " & Me![cboViewBody].Column(1) & " AND [WareGroup] = '" & Me![cboViewBody].Column(2) & "'"
    DoCmd.OpenForm "Frm_BodySherd", acNormal, , where
    Me![cboViewBody] = ""
Exit Sub

err_cboViewBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboViewDiagnostic_AfterUpdate()
On Error GoTo err_cboViewDiag
    Dim where
    where = "[Unit] = " & Me![cboViewDiagnostic].Column(1) & " AND [LetterCode] = '" & Me![cboViewDiagnostic].Column(2) & "' AND [FindNumber] = " & Me![cboViewDiagnostic].Column(3)
    
    DoCmd.OpenForm "Frm_Diagnostic", acNormal, , where
    Me![cboViewDiagnostic] = ""
Exit Sub

err_cboViewDiag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboViewUnIdBodySherds_AfterUpdate()
On Error GoTo err_cboViewUnIDBody
    Dim where
    where = "[Unit] = " & Me![cboViewUnIdBodySherds].Column(1) & " AND [Period] = '" & Me![cboViewUnIdBodySherds].Column(2) & "'"
    
    DoCmd.OpenForm "Frm_Unidentified_Body_Sherds", acNormal, , where
    Me![cboViewUnIdBodySherds] = ""
Exit Sub

err_cboViewUnIDBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboViewUnIDDiag_AfterUpdate()
On Error GoTo err_cboViewUnIDDiag
    Dim where
    where = "[Unit] = " & Me![cboViewUnIDDiag].Column(1) & " AND [Element] = '" & Me![cboViewUnIDDiag].Column(2) & "'"
    
    DoCmd.OpenForm "Frm_Unidentified_Diagnostic", acNormal, , where
    Me![cboViewUnIDDiag] = ""
Exit Sub

err_cboViewUnIDDiag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.Close acForm, Me.Name
    DoCmd.Restore
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAdd_Click()
On Error GoTo err_cmdAdd
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_Diagnostic", acNormal, , , acFormAdd
        Forms![Frm_Diagnostic]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
Exit Sub

err_cmdAdd:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewUnID_Click()
On Error GoTo err_cmdAddUnID
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_Unidentified_Diagnostic", acNormal, , , acFormAdd
        Forms![Frm_Unidentified_Diagnostic]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
Exit Sub

err_cmdAddUnID:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewUnIDBody_Click()
On Error GoTo err_cmdAddUnIDBody
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_Unidentified_Body_Sherds", acNormal, , , acFormAdd
        Forms![Frm_Unidentified_Body_Sherds]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
Exit Sub

err_cmdAddUnIDBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdNewBodySherd_Click()
On Error GoTo err_cmdNewBodySherd
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_BodySherd", acNormal, , , acFormAdd
        Forms![Frm_BodySherd]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
Exit Sub

err_cmdNewBodySherd:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()

End Sub

Private Sub cmdAddNewBody_Click()
On Error GoTo err_cmdAddNewBody
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_BodySherd", acNormal, , , acFormAdd
        Forms![Frm_BodySherd]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
    Exit Sub
err_cmdAddNewBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewNonNeolithic_Click()
On Error GoTo err_cmdNewNonNeolithic
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_NonNeolithic_Sherds", acNormal, , , acFormAdd
        Forms![Frm_NonNeolithic_Sherds]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
    Exit Sub
err_cmdNewNonNeolithic:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewUnidSherd_Click()
On Error GoTo err_cmdAddUnidSherd
    If Not IsNull(Me![txtUnit]) Then
        DoCmd.OpenForm "Frm_Unidentified_Diagnostic", acNormal, , , acFormAdd
        Forms![Frm_Unidentified_Diagnostic]![txtUnit] = Me![txtUnit]
    Else
        MsgBox "Please select a unit number first", vbInformation
    End If
Exit Sub

err_cmdAddUnidSherd:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
On Error GoTo err_current

    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
    'don't include find number as defaults to x
    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
        'new record allow GID entry
        Me![txtUnit].Enabled = True
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
    Else
        'existing entry lock
        Me![txtUnit].Enabled = False
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = Me.Section(0).BackColor
    End If
    'Me![cboViewDiagnostic].Requery
    'Me![cboViewUnIdBodySherds].Requery
    'Me![cboViewUnIDDiag].Requery
    'Me![cboViewBodySherds].Requery
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

