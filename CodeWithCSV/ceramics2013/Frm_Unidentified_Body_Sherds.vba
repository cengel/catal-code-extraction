Option Compare Database
Option Explicit



Private Sub cboFilterUnit_AfterUpdate()
'new 2010 filter for unit
On Error GoTo err_filterunit

If Me![cboFilterUnit] <> "" Then
    Me.Filter = "[Unit] = " & Me![cboFilterUnit]
    Me.FilterOn = True
    Me![cmdRemoveFilter].Visible = True
End If

Exit Sub

err_filterunit:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'if a filter is on - turn off
         If Me.FilterOn = True Then
            Me.FilterOn = False
            Me![cmdRemoveFilter].Visible = False
            Me![cboFilterUnit] = ""
        End If
    
         'for existing number the field will be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![txtShowUnit].Enabled = False Then Me![txtShowUnit].Enabled = True
        DoCmd.GoToControl "txtShowUnit"
        DoCmd.FindRecord Me![cboFindUnit], , , , True
        Me![cboFindUnit] = ""
        DoCmd.GoToControl "cboFindUnit"
        Me![txtShowUnit].Enabled = False
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
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

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
   ' DoCmd.GoToControl "Period" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![cboPeriod].Enabled = True
    Me![cboPeriod].Locked = False
    Me![cboPeriod].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddSameUnit_Click()
On Error GoTo err_cmdAddSameUnit_Click

    Dim thisunit
    thisunit = Me![txtUnit]
    
    'DoCmd.GoToControl "Period" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    If Not IsNull(thisunit) Then Me![txtUnit] = thisunit
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![cboPeriod].Enabled = True
    Me![cboPeriod].Locked = False
    Me![cboPeriod].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddSameUnit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRemoveFilter_Click()
On Error GoTo Err_cmdRemoveFilter

    Me.Filter = ""
    Me.FilterOn = False
    Me![cboFilterUnit] = ""
    DoCmd.GoToControl "cboFindUnit"
    Me![cmdRemoveFilter].Visible = False

    Exit Sub

Err_cmdRemoveFilter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()

On Error GoTo err_current

    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![cboPeriod] = "" Or IsNull(Me![cboPeriod])) Then
    'don't include find number as defaults to x
    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
        'new record allow GID entry
        Me![txtUnit].Enabled = True
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
        Me![cboPeriod].Enabled = True
        Me![cboPeriod].Locked = False
        Me![cboPeriod].BackColor = 16777215
    Else
        'existing entry lock
        Me![txtUnit].Enabled = False
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = Me.Section(0).BackColor
        Me![cboPeriod].Enabled = False
        Me![cboPeriod].Locked = True
        Me![cboPeriod].BackColor = Me.Section(0).BackColor
    End If

If Me.FilterOn = True Then
    Me![cmdRemoveFilter].Visible = True
End If
Exit Sub

err_current:
    'when setting a new record the focus can be in the tab element and then the code above fails - trap this
    If Err.Number = 2165 Then 'you can't hide a control that has the focus.
        Resume Next
    Else
        Call General_Error_Trap
    End If
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
Call CheckUnitDescript(Me![txtUnit])

End Sub
