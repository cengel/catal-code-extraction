Option Compare Database

Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'for existing number the field will be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![txtUnit].Enabled = False Then Me![txtUnit].Enabled = True
        DoCmd.GoToControl "txtUnit"
        DoCmd.FindRecord Me![cboFindUnit]
        Me![cboFindUnit] = ""
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboFindUnit_NotInList(NewData As String, Response As Integer)
On Error GoTo err_cbo

MsgBox "This Unit has not been entered into the Ceramics Sheet yet", vbExclamation, "Unit not in lisr"
Response = acDataErrContinue
Me![cboFindUnit].Undo

Exit Sub

err_cbo:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_cmdAddNew_Click

    'DoCmd.Close acForm, Me.Name
    DoCmd.Quit acQuitSaveAll
    
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDiag_Click()
On Error GoTo err_diag


    Dim sql, cd
    cd = InputBox("Please enter the new S number:", "Sherd Number")
    If cd <> "" Then
        sql = "INSERT INTO [Ceramics_Stratified_Diagnostic_Sherds_NY06] ([Unit], [LetterCode], [FindNumber]) VALUES (" & Me![txtUnit] & ", 'S'," & cd & ");"
        DoCmd.RunSQL sql
        'me![frm_subform_strat_body_Sherds].Forms![WARE CODE] = Me![txtWareCode]
        Me![frm_subform_Strat_Diagnostic_Sherds].Requery
        DoCmd.GoToControl "frm_subform_Strat_Diagnostic_Sherds"
        DoCmd.GoToControl "Weight gr"
    End If
Exit Sub

err_diag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdNewBody_Click()
On Error GoTo err_txtWare


    Dim sql, cd
    cd = InputBox("Please enter the new ware code:", "Ware Code")
    If cd <> "" Then
        sql = "INSERT INTO [Ceramics_Stratified_Body_Sherds_NY06] ([Unit], [WARE CODE]) VALUES (" & Me![txtUnit] & ", '" & cd & "');"
        DoCmd.RunSQL sql
        'me![frm_subform_strat_body_Sherds].Forms![WARE CODE] = Me![txtWareCode]
        Me![frm_subform_strat_body_Sherds].Requery
        DoCmd.GoToControl "frm_subform_strat_body_sherds"
        DoCmd.GoToControl "total body sherds"
    End If
Exit Sub

err_txtWare:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnI_Click()
On Error GoTo err_uni


    Dim sql, cd, dec
    cd = InputBox("Please enter the new Element:", "Element")
    If cd <> "" Then
        dec = InputBox("Please enter the new Decoration:", "Decoration")
        If dec <> "" Then
            sql = "INSERT INTO [Ceramics_Stratified_Diagnostic_UnIdentified_Sherds_NY06] ([Unit], [Element], [Decoration]) VALUES (" & Me![txtUnit] & ", '" & cd & "', '" & dec & "');"
            DoCmd.RunSQL sql
            'me![frm_subform_strat_body_Sherds].Forms![WARE CODE] = Me![txtWareCode]
            Me![frm_subform_strat_Diag_Unid_Sherds].Requery
            DoCmd.GoToControl "frm_subform_strat_Diag_Unid_Sherds"
            DoCmd.GoToControl "Total Mineral Tempered"
        End If
    End If
Exit Sub

err_uni:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'new code for 2006
On Error GoTo err_current

    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![cboFindLetter] = "" Or IsNull(Me![cboFindLetter])) And (Me![txtFindNumber] = "" Or IsNull(Me![txtFindNumber])) Then
    'don't include find number as defaults to x
    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
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
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub

Private Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    Call General_Error_Trap
    Resume Exit_go_previous2_Click
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    Call General_Error_Trap
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    Call General_Error_Trap
    Resume Exit_go_last_Click
End Sub

Private Sub txtUnit_AfterUpdate()
'Me![frm_subform_UnitDetails].Requery
Me.Refresh
End Sub

Private Sub txtWareCode_AfterUpdate()
On Error GoTo err_txtWare

If Me![txtWareCode] <> "" Then
    Dim sql
    sql = "INSERT INTO [Ceramics_Stratified_Body_Sherds_NY06] ([Unit], [WARE CODE]) VALUES (" & Me![txtUnit] & ", '" & Me![txtWareCode] & "');"
    DoCmd.RunSQL sql
    'me![frm_subform_strat_body_Sherds].Forms![WARE CODE] = Me![txtWareCode]
    Me![frm_subform_strat_body_Sherds].Requery
    DoCmd.GoToControl "frm_subform_strat_body_sherds"
    DoCmd.GoToControl "total body sherds"
    Me![txtWareCode] = ""
End If
Exit Sub

err_txtWare:
    Call General_Error_Trap
    Exit Sub
End Sub
