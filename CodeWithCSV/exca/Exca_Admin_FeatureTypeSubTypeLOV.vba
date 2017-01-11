Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




Private Sub cboFindFeature_AfterUpdate()
'v9.2 saj find a feature from the list
On Error GoTo err_find
    If Me![cboFindFeature] <> "" Then
        DoCmd.GoToControl "txtFeatureType"
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
' The feature type value is used on the Feature Sheet
' so must check all this values to see if this type is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdEdit_Click

Dim checkValidAction, retval

    'check space sheet level field
    checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "edit")
    
    If checkValidAction = "ok" Then
        'edit action can go ahead - at present simply offer an input box for this
        retval = InputBox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Feature Type that you wish to replace this entry with:", "Enter edited Feature Type")
        If retval <> "" Then
            Me![txtFeatureType] = retval
        End If
    ElseIf checkValidAction = "fail" Then
        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
    Else
        MsgBox checkValidAction, vbExclamation, "Action Report"
    End If
    
Exit_cmdEdit_Click:
    Exit Sub

Err_cmdEdit_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdDelete_Click()
'v9.2 SAJ - check if user can delete this record
' The feature type value is used on the Feature Sheet
' so must check all this values to see if this type is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdDelete_Click

Dim checkValidAction, retval

    'first check if this has any subtypes
    If Not IsNull(Me![Exca: Admin_Subform_FeatureSubType].Form![FeatureTypeID]) Then
        MsgBox "You must delete the Sub types associated with this feature first", vbInformation, "Invalid Action"
    Else

        'check FEATURE sheet feature type field
        checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "delete")
        If checkValidAction = "ok" Then
        'delete action can go ahead - at present simply offer an input box for this
            retval = MsgBox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFeatureType] & " from the list of available Feature Types?", vbExclamation + vbYesNo, "Confirm Deletion")
            If retval = vbYes Then
                Me.AllowDeletions = True
                DoCmd.RunCommand acCmdDeleteRecord
                Me.AllowDeletions = False
            End If
        
        ElseIf checkValidAction = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            'is error occured in CheckIfLOVValueUsed it will come back empty so err already displayed
            If Not IsEmpty(checkValidAction) Then MsgBox checkValidAction, vbExclamation, "Action Report"
        End If
    End If
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
    
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
