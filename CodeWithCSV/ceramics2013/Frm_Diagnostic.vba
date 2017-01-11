Option Compare Database
Option Explicit

Sub DealWithElement(ctrl As Control)
'this sets the tabs when the element tick boxes

Select Case ctrl.Name
Case "chkRim"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_rim].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(0).Visible = False
        Else
            MsgBox "There is rim data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(0).Visible = True
        Me!tabcontrolElements.Pages(0).SetFocus
        
    End If
Case "chkBase"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_base].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(1).Visible = False
        Else
            MsgBox "There is base data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(1).Visible = True
        Me!tabcontrolElements.Pages(1).SetFocus
    End If
Case "chkLug"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_lug].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(2).Visible = False
        Else
            MsgBox "There is lug data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(2).Visible = True
        Me!tabcontrolElements.Pages(2).SetFocus
    End If
Case "chkPedestal"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_pedestal].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(3).Visible = False
        Else
            MsgBox "There is pedestal data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(3).Visible = True
        Me!tabcontrolElements.Pages(3).SetFocus
    End If
Case "chkLid"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_lid].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(4).Visible = False
        Else
            MsgBox "There is lid data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(4).Visible = True
        Me!tabcontrolElements.Pages(4).SetFocus
    End If
Case "chkFoot"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_foot].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(5).Visible = False
        Else
            MsgBox "There is foot data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(5).Visible = True
        Me!tabcontrolElements.Pages(5).SetFocus
    End If
Case "chkHandle"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_handle].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(6).Visible = False
        Else
            MsgBox "There is handle data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(6).Visible = True
        Me!tabcontrolElements.Pages(6).SetFocus
    End If
Case "chkKnob"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_knob].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(7).Visible = False
        Else
            MsgBox "There is knob data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(7).Visible = True
        Me!tabcontrolElements.Pages(7).SetFocus
    End If
Case "chkcarin"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_carination].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(9).Visible = False
        Else
            MsgBox "There is carination data recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(9).Visible = True
        Me!tabcontrolElements.Pages(9).SetFocus
        
    End If
Case "chkDeco"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_decoration].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(8).Visible = False
        Else
            MsgBox "There is decoration information recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(8).Visible = True
        Me!tabcontrolElements.Pages(8).SetFocus
        
    End If
Case "chkRareform"
    If ctrl = False Then
        If Me![frm_sub_diagnostic_rareform].Form.RecordsetClone.RecordCount = 0 Then
            Me!tabcontrolElements.Pages(10).Visible = False
        Else
            MsgBox "There is rare form information recorded, please delete this first"
            ctrl = True
        End If
    Else
        Me!tabcontrolElements.Pages(10).Visible = True
        Me!tabcontrolElements.Pages(10).SetFocus
        
    End If
End Select
'MsgBox ctrl.Name


End Sub
Private Sub Command23_Click()
On Error GoTo Err_Command23_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command23_Click:
    Exit Sub

Err_Command23_Click:
    MsgBox Err.Description
    Resume Exit_Command23_Click
    
End Sub

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

Private Sub chkBase_Click()
DealWithElement Me!chkBase

End Sub

Private Sub chkCarin_Click()
DealWithElement Me!chkCarin
End Sub

Private Sub chkDeco_Click()
DealWithElement Me!chkDeco
End Sub

Private Sub chkFoot_Click()
DealWithElement Me!chkFoot

End Sub

Private Sub chkHandle_Click()
DealWithElement Me!chkHandle

End Sub

Private Sub chkKnob_Click()
DealWithElement Me!chkKnob

End Sub

Private Sub chkLid_Click()
DealWithElement Me!chkLid

End Sub

Private Sub chkLug_Click()
DealWithElement Me!chkLug

End Sub

Private Sub chkPedestal_Click()
DealWithElement Me!chkPedestal

End Sub

Private Sub chkRareform_Click()
DealWithElement Me!chkRareform
End Sub

Private Sub chkRim_Click()
DealWithElement Me!chkRim

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

    Dim thisunit
    thisunit = Me![txtUnit]
    
    DoCmd.GoToControl "Phase" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    If Not IsNull(thisunit) Then Me![txtUnit] = thisunit
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![LetterCode].Enabled = True
    Me![LetterCode].Locked = False
    Me![LetterCode].BackColor = 16777215
    Me![FindNumber].Enabled = True
    Me![FindNumber].Locked = False
    Me![FindNumber].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewUnit_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToControl "Phase" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![LetterCode].Enabled = True
    Me![LetterCode].Locked = False
    Me![LetterCode].BackColor = 16777215
    Me![FindNumber].Enabled = True
    Me![FindNumber].Locked = False
    Me![FindNumber].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'allow deletion of entire record
On Error GoTo err_delete

Call DeleteDiagnosticRecord(Me![txtUnit], Me![LetterCode], Me![FindNumber])

Exit Sub

err_delete:
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

Private Sub cmdReNum_Click()
On Error GoTo err_ReNum
Dim val

    If Me![txtUnit] <> "" And Me!LetterCode <> "" And Me![FindNumber] <> "" Then
        val = ReNumberDiagnostic(Me![txtUnit], Me!LetterCode, Me![FindNumber])
        'new number if successful has been fed into find cbo so search to display.
        'if failed to update then cbofind will be blank so nothing happens
        cboFindUnit_AfterUpdate
        'MsgBox val
    Else
        MsgBox "Incomplete GID to process", vbInformation, "Action Cancelled"
    End If
Exit Sub

err_ReNum:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub FindNumber_AfterUpdate()
'This is the wierdest thing I have ever seen in Access - entering a new number eg: 999994.S1 after
'tabbing out of the findnumber the subforms were all grabbing data from 17017.S1 even though the main
'record were still the new number, this happened for all new numbers where the find number matched one for
'unit 17017.
'
'NO IDEA WHY but putting save here stops it. Wish I knew why though! SAJ 9th July 2009
On Error GoTo err_findnumber

DoCmd.RunCommand acCmdSaveRecord

Exit Sub

err_findnumber:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()

On Error GoTo err_current

    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![LetterCode] = "" Or IsNull(Me![LetterCode])) And (Me![FindNumber] = "" Or IsNull(Me![FindNumber])) Then
    'don't include find number as defaults to x
    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
        'new record allow GID entry
        Me![txtUnit].Enabled = True
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
        Me![LetterCode].Enabled = True
        Me![LetterCode].Locked = False
        Me![LetterCode].BackColor = 16777215
        Me![FindNumber].Enabled = True
        Me![FindNumber].Locked = False
        Me![FindNumber].BackColor = 16777215
    Else
        'existing entry lock
        Me![txtUnit].Enabled = False
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = Me.Section(0).BackColor
        Me![LetterCode].Enabled = False
        Me![LetterCode].Locked = True
        Me![LetterCode].BackColor = Me.Section(0).BackColor
        Me![FindNumber].Enabled = False
        Me![FindNumber].Locked = True
        Me![FindNumber].BackColor = Me.Section(0).BackColor
    End If

'set focus to top
If Me![txtUnit].Enabled = True Then
    DoCmd.GoToControl "txtUnit"
Else
    'DoCmd.GoToControl "cboFindUnit"
    ' added conditon below that this not be in datasheet view,
    ' otherwise it will produce an error message
    ' CE - June 2014
    If Me.CurrentView <> 2 Then
        DoCmd.GoToControl "cboFindUnit"
    End If
End If

If Me![Rim] = True Then
    'MsgBox "rim present"
    'Me!tabcontrolElements.Pages(0).Enabled = True
    Me!tabcontrolElements.Pages(0).Visible = True
Else
    'MsgBox "no rim"
    Me!tabcontrolElements.Pages(0).Visible = False
End If

If Me![Base] = True Then
    Me!tabcontrolElements.Pages(1).Visible = True
Else
    Me!tabcontrolElements.Pages(1).Visible = False
End If

If Me![Lug] = True Then
    Me!tabcontrolElements.Pages(2).Visible = True
Else
    Me!tabcontrolElements.Pages(2).Visible = False
End If

If Me![Pedestal] = True Then
    Me!tabcontrolElements.Pages(3).Visible = True
Else
    Me!tabcontrolElements.Pages(3).Visible = False
End If

If Me![Lid] = True Then
    Me!tabcontrolElements.Pages(4).Visible = True
Else
    Me!tabcontrolElements.Pages(4).Visible = False
End If

If Me![Foot] = True Then
    Me!tabcontrolElements.Pages(5).Visible = True
Else
    Me!tabcontrolElements.Pages(5).Visible = False
End If

If Me![Handle] = True Then
    Me!tabcontrolElements.Pages(6).Visible = True
Else
    Me!tabcontrolElements.Pages(6).Visible = False
End If

If Me![Knob] = True Then
    Me!tabcontrolElements.Pages(7).Visible = True
Else
    Me!tabcontrolElements.Pages(7).Visible = False
End If

If Me![Decoration] = True Then
    Me!tabcontrolElements.Pages(8).Visible = True
Else
    Me!tabcontrolElements.Pages(8).Visible = False
End If

If Me![Carination] = True Then
    Me!tabcontrolElements.Pages(9).Visible = True
Else
    Me!tabcontrolElements.Pages(9).Visible = False
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

Private Sub Form_Open(Cancel As Integer)
'new to disable admin features to non admin users
On Error GoTo er_open

If GetGeneralPermissions = "Admin" Then
    Me!cmdReNum.Enabled = True
    Me!cmdDelete.Enabled = True
Else
    Me!cmdReNum.Enabled = False
    Me!cmdDelete.Enabled = False
End If

If Me.FilterOn = True Then
    Me![cmdRemoveFilter].Visible = True
End If


Exit Sub

er_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub frm_sub_diagnostic_carination_Enter()

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

Private Sub WARE_CODE_NotInList(NewData As String, Response As Integer)
On Error GoTo err_warecode_NotInList

Dim retVal, sql
retVal = MsgBox("This Ware Code does not appear in the pre-defined list. Have you checked the list to make sure there is no match?", vbQuestion + vbYesNo, "New ware code")
If retVal = vbYes Then
    MsgBox "Ok this ware code will now be added to the list", vbInformation, "New Ware Code Allowed"
    'allow value,
     Response = acDataErrAdded
    
    Dim desc
    desc = InputBox("Please enter the description for this new code eg: DMS-fine", "Ware Code Description")
    If desc <> "" Then
        sql = "INSERT INTO [Ceramics_Code_Warecode_LOV] ([WareCode], [Description]) VALUES ('" & NewData & "', '" & desc & "');"
    Else
        sql = "INSERT INTO [Ceramics_Code_Warecode_LOV] ([WareCode]) VALUES ('" & NewData & "');"
    End If
    DoCmd.RunSQL sql
    
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub

err_warecode_NotInList:
    Call General_Error_Trap

    Exit Sub

End Sub
