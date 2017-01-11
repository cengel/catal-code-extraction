Option Compare Database
Option Explicit

Private Sub cmdDelete_Click()
'v9.2 SAJ - check if user can delete this record
' The feature sub type value is used on the Feature Sheet
' so must check all this value along with its associated feature type to see if
' this type is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdDelete_Click

Dim checkValidAction, retVal

       'check FEATURE sheet feature subtype field
'        checkValidAction = CheckIfLOVValueUsed("Exca:SubFeatureTypeLOV", "FeatureSubType", Me![txtFeatureSubType], "Exca: Features", "Feature Number", "FeatureSubType", "delete", " AND [Feature Type] = '" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![txtFeatureType] & "'")
'        If checkValidAction = "ok" Then
'        'delete action can go ahead - at present simply offer an input box for this
'            retVal = msgbox("No records refer to this Feature SubType (" & Me![txtFeatureSubType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFeatureSubType] & " from the list of available Feature Subtypes?", vbExclamation + vbYesNo, "Confirm Deletion")
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
'
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
 '   Call General_Error_Trap
    Exit Sub
    

End Sub

Private Sub cmdEdit_Click()
'v9.2 SAJ - check if user can edit this record
' The feature subtype value is used on the Feature Sheet
' so must check this value along with its associated feature type to see if this sub
' type is used before allow edit.
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdEdit_Click

Dim checkValidAction, retVal

    'check feature sheet feature and feature subtype field for this match (as the text of the subtype maybe used for a diff feature also and don't want that to stop edit)
'    checkValidAction = CheckIfLOVValueUsed("Exca:SubFeatureTypeLOV", "FeatureSubType", Me![txtFeatureSubType], "Exca: Features", "Feature Number", "FeatureSubType", "edit", " AND [Feature Type] = '" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![txtFeatureType] & "'")
'
'    If checkValidAction = "ok" Then
'        'edit action can go ahead - at present simply offer an input box for this
'        retVal = InputBox("No records refer to this Feature sub type (" & Me![txtFeatureSubType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Feature Sub Type that you wish to replace this entry with:", "Enter edited Feature Sub Type")
'        If retVal <> "" Then
'            Me![txtFeatureSubType] = retVal
'        End If
'    ElseIf checkValidAction = "fail" Then
'        msgbox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'    Else
'        msgbox checkValidAction, vbExclamation, "Action Report"
'    End If
    
Exit_cmdEdit_Click:
    Exit Sub

Err_cmdEdit_Click:
'    Call General_Error_Trap
    Exit Sub
    

End Sub

'******************************************************
' This subform is new with v9.2
' SAJ v9.2
'******************************************************
Private Sub cmdNewSubType_Click()
'v9.2 - allow new subtype to be added
On Error GoTo err_cmdNewSubType_Click

    If Forms![Finds: Admin_MaterialGroupSubGroupLOV]![MaterialID] <> "" Then
        Dim sql, retVal
        retVal = InputBox("Please enter the new object type for the material type '" & Forms![Finds: Admin_MaterialGroupSubGroupLOV]![MaterialGroup] & "': ", "Enter new subtype")
        If retVal <> "" Then
            sql = "INSERT INTO [MaterialSubGroup_ObjectTypes] ([MaterialSubGroupID], [ObjectTypeText]) VALUES (" & Forms![Finds: Admin_MaterialGroupSubGroupLOV]![Finds: Admin_Subform_MaterialSubGroup].Form![MaterialSubGroupID] & ", '" & retVal & "');"
            DoCmd.RunSQL sql
            Me.Requery
        End If
    Else
        MsgBox "Sorry not all the data necessary to make a new subtype is available.", vbExclamation, "Invalid Action"
    End If
Exit Sub

err_cmdNewSubType_Click:
 '   Call General_Error_Trap
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
'    permiss = GetGeneralPermissions
'    If permiss <> "ADMIN" Then
'        msgbox "Sorry but only Administrators have access to this form"
'        DoCmd.close acForm, Me.Name
'    End If
Exit Sub

err_Form_Open:
 '   Call General_Error_Trap
    Exit Sub
End Sub
