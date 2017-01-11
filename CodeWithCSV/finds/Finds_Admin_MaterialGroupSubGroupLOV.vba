Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




Private Sub cboFindFeature_AfterUpdate()
'v9.2 saj find a feature from the list
On Error GoTo Err_find
    If Me![cboFindFeature] <> "" Then
        DoCmd.GoToControl "txtFeatureType"
        DoCmd.FindRecord Me![cboFindFeature]
        DoCmd.GoToControl "cboFindFeature"
    End If
    'Me.AllowEdits = False
Exit Sub

Err_find:
    'Call General_Error_Trap
    MsgBox Err.Description
    Exit Sub
End Sub







Private Sub cboFindFeature_GotFocus()
'v9.2 SAJ - as the form does not allow edits the list can't be selected from -
'when list gets focus turn on allow edits and off again when loses it
'Me.AllowEdits = True
End Sub

Private Sub cboFindFeature_LostFocus()
'v9.2 SAJ - as the form does not allow edits the list can't be selected from -
'when list gets focus turn on allow edits and off again when loses it
'Me.AllowEdits = False
End Sub

Private Sub cmdAddNew_Click()
'SAJ - add a new record
On Error GoTo err_cmdAddNew_Click

    'DoCmd.RunCommand acCmdRecordsGoToNew
    'v3.2 - get name of new group via input box
    Dim resp
    resp = InputBox("Please enter the name of the new Material Group", "New Material Group")
    If resp <> "" Then
        'check name doesn't exist already
        Dim checkit
        checkit = DLookup("[MaterialGroup]", "[Finds_Code_MaterialGroup]", "[MaterialGroup] = '" & resp & "'")
        If IsNull(checkit) Then
            'ok to add
            DoCmd.RunCommand acCmdRecordsGoToNew
            Me![MaterialGroup] = resp
            Me![Exca: Admin_Subform_FeatureSubType].Form![cboMaterialSubgroupText].Visible = True
            Me![Exca: Admin_Subform_FeatureSubType].Form![txtMaterialSubgroupText].Visible = False
            'DoCmd.GoToControl Me![Exca: Admin_Subform_FeatureSubType].Form![txtMaterialSubgroupText].Name
            DoCmd.GoToControl "txtFeatureType"
        Else
            'already exists
            MsgBox "This Material Group already exists, the system will display it now", vbInformation, "Material Group Exists"
            DoCmd.GoToControl "txtFeatureType"
            DoCmd.FindRecord resp
            Me!cboFindFeature = resp
            DoCmd.GoToControl "cboFindFeature"
        End If
    End If

Exit Sub

err_cmdAddNew_Click:
   ' Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
'new v4.3 2009 - all print out
On Error GoTo err_print

    DoCmd.OpenReport "R_materials", acViewPreview

Exit Sub

err_print:
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

'Dim checkValidAction, retVal
'
'    'check space sheet level field
'    checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "edit")
'
'    If checkValidAction = "ok" Then
'        'edit action can go ahead - at present simply offer an input box for this
'        retVal = InputBox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Feature Type that you wish to replace this entry with:", "Enter edited Feature Type")
'        If retVal <> "" Then
'            Me![txtFeatureType] = retVal
'        End If
'    ElseIf checkValidAction = "fail" Then
'        msgbox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'    Else
'        msgbox checkValidAction, vbExclamation, "Action Report"
 '   End If
    
Exit_cmdEdit_Click:
    Exit Sub

Err_cmdEdit_Click:
'    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdDelete_Click()
'v9.2 SAJ - check if user can delete this record
' The feature type value is used on the Feature Sheet
' so must check all this values to see if this type is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdDelete_Click

Dim checkValidAction, retVal

    'first check if this has any subtypes
'    If Not IsNull(Me![Exca: Admin_Subform_FeatureSubType].Form![FeatureTypeID]) Then
'        msgbox "You must delete the Sub types associated with this feature first", vbInformation, "Invalid Action"
'    Else
'
'        'check FEATURE sheet feature type field
'        checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "delete")
'        If checkValidAction = "ok" Then
'        'delete action can go ahead - at present simply offer an input box for this
'            retVal = msgbox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFeatureType] & " from the list of available Feature Types?", vbExclamation + vbYesNo, "Confirm Deletion")
'            If retVal = vbYes Then
'                Me.AllowDeletions = True
'                DoCmd.RunCommand acCmdDeleteRecord
'                Me.AllowDeletions = False
'            End If
'
'        ElseIf checkValidAction = "fail" Then
'            msgbox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'        Else
'            'is error occured in CheckIfLOVValueUsed it will come back empty so err already displayed
'            If Not IsEmpty(checkValidAction) Then msgbox checkValidAction, vbExclamation, "Action Report"
'        End If
'    End If
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
'    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

'    Dim permiss
'    permiss = GetGeneralPermissions
'    If permiss <> "ADMIN" Then
'        msgbox "Sorry but only Administrators have access to this form"
'        DoCmd.close acForm, Me.Name
'    End If
Exit Sub

err_Form_Open:
'    Call General_Error_Trap
    Exit Sub
End Sub
