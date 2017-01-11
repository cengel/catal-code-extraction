Option Compare Database
Option Explicit

Private Sub cboFilterUnit_AfterUpdate()
'put a where clause on the subform to only show that unit
On Error GoTo err_filter

    If Me![cboFilterUnit] <> "" Then
        Me![frm_subform_level2].Form.RecordSource = "SELECT * FROM Q_GS_Level2_with_Excavation WHERE [Unit] = " & Me![cboFilterUnit]
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFilterUnit_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofilterNot

    MsgBox "Sorry this Unit does not exist in this database yet", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboFilterUnit].Undo
Exit Sub

err_cbofilterNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    DoCmd.GoToControl "frm_subform_level2"
    DoCmd.GoToControl "GID"
    DoCmd.FindRecord Me![cboFind]
    DoCmd.GoToControl "cboAnalyst"
End If


Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this GID does not exist in the database", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub checkPriorityUnits_Click()
On Error GoTo err_checkPriorityUnits

If Me.checkPriorityUnits = True Then
    Me![frm_subform_level2].Form.RecordSource = "SELECT * FROM Q_GS_Level2_with_Excavation WHERE [Priority Unit] = 1"
Else
    Me![frm_subform_level2].Form.RecordSource = "Q_GS_Level2_with_Excavation"
End If
    
' check if there is a filter for units (ie if the reset button is visible)and for now just reset it.
If Me![cmdRemoveFilter].Visible = True Then
    Me![cboFilterUnit] = ""
    Me![cmdRemoveFilter].Visible = False
End If
    
Exit Sub

err_checkPriorityUnits:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_close

    DoCmd.OpenForm "Frm_GS_MainMenu", acNormal, , , acFormPropertySettings
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'********************************************************************
' Create new record
' SAJ
' added locks to disable entry of other fields until we control for duplicate GID
' in Level 1 and Level 2 tables when new GID is added
' we remove the locks later when we do the check after updating
' the three fields that allow for entry: Unit, Lettercode, Fieldnumber
' CE June 2014
'********************************************************************
On Error GoTo Err_cmdgonew_Click
Dim ctl As Control
    DoCmd.GoToControl Me![frm_subform_level2].Name
    DoCmd.GoToRecord , , acNewRec
    DoCmd.GoToControl Me![frm_subform_level2].Form![Unit].Name
    
    'lock all fields except Unit, Letter, FindNo - CE June 2014
    For Each ctl In Me.frm_subform_level2.Controls
        
        If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) And Not (ctl.Name = "Unit" Or ctl.Name = "cboLettercode" Or ctl.Name = "FindNumber") Then
            ctl.Locked = True
        End If
        
'        If ctl.Name = "Unit" Or ctl.Name = "cboLettercode" Or ctl.Name = "FindNumber" Then
'            ctl.Locked = False
'        End If
            
        'If (ctl.Name = "cboAnalyst" Or ctl.ControlType = acTextBox) And Not (ctl.Name = "Unit" Or ctl.Name = "Lettercode" Or ctl.Name = "FindNumber") Then
        '    ctl.Locked = True
        'End If
    
    Next ctl

    Exit Sub

Err_cmdgonew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    DoCmd.GoToControl Me![frm_subform_level2].Name
    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_cmdgofirst_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoLast_Click()
'********************************************************************
' Go to last record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoLast_Click

    DoCmd.GoToControl Me![frm_subform_level2].Name
    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_cmdgoLast_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoNext_Click()
'********************************************************************
' Go to next record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoNext_Click

    DoCmd.GoToControl Me![frm_subform_level2].Name
    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_cmdgoNext_Click:
    If Err.Number = 2105 Then
        MsgBox "No more records to show", vbInformation, "End of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdGoPrev_Click()
'********************************************************************
' Go to previous record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoprevious_Click

    DoCmd.GoToControl Me![frm_subform_level2].Name
    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_cmdgoprevious_Click:
    If Err.Number = 2105 Then
        MsgBox "Already at the begining of the recordset", vbInformation, "Begining of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![frm_subform_level2].Form.[GID] <> "" Then
        DoCmd.OpenForm "Frm_Pop_DataOutLevel2", acNormal, , , acFormPropertySettings, , Me![frm_subform_level2].Form![GID]
    Else
        MsgBox "The output options form cannot be shown when there is no record selected", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRemoveFilter_Click()
'remove the where clause on the subform acting as a unit filter
On Error GoTo err_Removefilter

    Me![cboFilterUnit] = ""
    Me![frm_subform_level2].Form.RecordSource = "Q_GS_Level2_with_Excavation"
    DoCmd.GoToControl "cboFind"
    Me![cmdRemoveFilter].Visible = False
   

Exit Sub

err_Removefilter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize

End Sub

Private Sub tglForm_Click()
On Error GoTo err_tglForm_Click

Debug.Print Me.frm_subform_level2.Form.CurrentView

    If Me.frm_subform_level2.Form.CurrentView = 1 Then
        Forms![Frm_Level2].[frm_subform_level2].SetFocus
        RunCommand acCmdSubformDatasheetView
    Else
        Forms![Frm_Level2].[frm_subform_level2].SetFocus
        RunCommand acCmdSubformFormView
    End If
Exit Sub

err_tglForm_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
