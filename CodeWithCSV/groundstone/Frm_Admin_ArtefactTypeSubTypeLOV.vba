Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




Private Sub cboClass_AfterUpdate()
'the class must be select for a new record, after it has been selected allow Type to be entered
On Error GoTo err_cboClass

If Me![cboClass] <> "" Then
    Me![txtToolTypes].Enabled = True
    DoCmd.GoToControl Me![txtToolTypes].Name
    Me![cmdEdit].Enabled = True
    Me![cmdDelete].Enabled = True
    Me![Frm_Admin_Subform_ArtefactSubType].Enabled = True
    Me![cboClass].Enabled = False
    Me![cboClass].BackStyle = 0
    Me![cboClass].Locked = True
End If
Exit Sub

err_cboClass:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindFeature_AfterUpdate()
'v9.2 saj find a feature from the list
On Error GoTo err_find
    If Me![cboFindFeature] <> "" Then
        DoCmd.GoToControl "txtToolTypes"
        DoCmd.FindRecord Me![cboFindFeature]
    End If
    Me.AllowEdits = False
Exit Sub

err_find:
    Call General_Error_Trap
    Exit Sub
End Sub







Private Sub cboFindFeature_GotFocus()
'v9.2 SAJ - as the form does not allow edits the list can't be selected from -
'when list gets focus turn on allow edits and off again when loses it
Me.AllowEdits = True
End Sub

Private Sub cboFindFeature_LostFocus()
'v9.2 SAJ - as the form does not allow edits the list can't be selected from -
'when list gets focus turn on allow edits and off again when loses it
Me.AllowEdits = False
End Sub

Private Sub cmdAddNew_Click()
'v9.2 SAJ - add a new record
On Error GoTo err_cmdAddNew_Click

    DoCmd.RunCommand acCmdRecordsGoToNew
    'to ensure that the user enters the artefact class associated with this type
    'enable the combo and disable all other controls
    Me![cboClass].Visible = True
    Me![cboClass].Enabled = True
    Me![cboClass].Locked = False
    Me![cboClass].BackStyle = 1
    DoCmd.GoToControl Me![cboClass].Name
    Me![txtToolTypes].Enabled = False
    Me![cmdEdit].Enabled = False
    Me![cmdDelete].Enabled = False
    Me![Frm_Admin_Subform_ArtefactSubType].Enabled = False
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
'v9.2 SAJ - close the form
    DoCmd.Close acForm, Me.Name
End Sub



Private Sub cmdEdit_Click()
'v9.2 SAJ - check if user can edit this record
' The artefact type value is used on the worked data
' so must check all this values to see if this type is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdEdit_Click
'
'Dim checkValidAction, retVal
'
'    'check worked Artefact Type field
'    checkValidAction = CheckIfLOVValueUsed("GroundStone List of Values: Artefact Type", "Tool Types", Me![txtToolTypes], "GroundStone 3: Worked Stone Basics", "GID", "Artefact Type", "edit", False)
'
'    If checkValidAction = "ok" Then
'        'edit action can go ahead - at present simply offer an input box for this
'        retVal = InputBox("No records refer to this Artefact Type (" & Me![txtToolTypes] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Artefact Type that you wish to replace this entry with:", "Enter edited Artefact Type")
'        If retVal <> "" Then
'            Me![txtToolTypes] = retVal
'        End If
'    ElseIf checkValidAction = "fail" Then
'        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'    Else
'        MsgBox checkValidAction, vbExclamation, "Action Report"
'    End If
'
'Exit_cmdEdit_Click:
'    Exit Sub
'
'Err_cmdEdit_Click:
'    Call General_Error_Trap
'    Exit Sub
'
End Sub

Private Sub cmdDelete_Click()
'v9.2 SAJ - check if user can delete this record
' The feature type value is used on the Feature Sheet
' so must check all this values to see if this type is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdDelete_Click
'
'Dim checkValidAction, retVal
'
'    'first check if this has any subtypes
'    If Not IsNull(Me![Frm_Admin_Subform_ArtefactSubType].Form![Code]) Then
'        MsgBox "You must delete the Sub types associated with this Artefact first", vbInformation, "Invalid Action"
'    Else
'
'        'check worked data artefact type field
'        checkValidAction = CheckIfLOVValueUsed("GroundStone List of Values: Artefact Type", "Tool Types", Me![txtToolTypes], "GroundStone 3: Worked Stone Basics", "GID", "Artefact Type", "delete", False)
'        If checkValidAction = "ok" Then
'        'delete action can go ahead - at present simply offer an input box for this
'            retVal = MsgBox("No records refer to this Artefact Type (" & Me![txtToolTypes] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtToolTypes] & " from the list of available Artefact Types?", vbExclamation + vbYesNo, "Confirm Deletion")
'            If retVal = vbYes Then
'                Me.AllowDeletions = True
'                DoCmd.RunCommand acCmdDeleteRecord
'                Me.AllowDeletions = False
'            End If
'
'        ElseIf checkValidAction = "fail" Then
'            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'        Else
'            'is error occured in CheckIfLOVValueUsed it will come back empty so err already displayed
'            If Not IsEmpty(checkValidAction) Then MsgBox checkValidAction, vbExclamation, "Action Report"
'        End If
'    End If
'Exit_cmdDelete_Click:
'    Exit Sub
'
'Err_cmdDelete_Click:
'    Call General_Error_Trap
'    Exit Sub
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
