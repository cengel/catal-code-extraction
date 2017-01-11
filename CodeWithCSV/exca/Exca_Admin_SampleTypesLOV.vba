Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
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

Private Sub cmdDelete_Click()
' Season 2009 SAJ - check if user can delete this record
' The sample type is used in in the Exca: Samples table - so must check all this to see
' if this type is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdDelete_Click

Dim checkValidAction, checkValidAction2, checkValidAction3, retval

    'check samples type field
    checkValidAction = CheckIfLOVValueUsed("Exca:SampleTypeLOV", "SampleType", Me![txtSampleType], "Exca: Samples", "Unit Number", "Sample Type", "delete")
    
    If checkValidAction = "ok" Then
        
                'delete action can go ahead - at present simply offer an input box for this
                retval = MsgBox("No records refer to this Type (" & Me![txtSampleType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtSampleType] & " from the list of available Types?", vbExclamation + vbYesNo, "Confirm Deletion")
                If retval = vbYes Then
                    Me.AllowDeletions = True
                    DoCmd.RunCommand acCmdDeleteRecord
                    Me.AllowDeletions = False
                End If
                
    ElseIf checkValidAction = "fail" Then
        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
    Else
        MsgBox checkValidAction, vbExclamation, "Action Report"
    End If
    
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
' Season 2009 SAJ - check if user can edit this record
' The sample type is used in in the Exca: Samples table - so must check all this to see
' if this type is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdEdit_Click

Dim checkValidAction, checkValidAction2, checkValidAction3, retval

    'check samples type field
    checkValidAction = CheckIfLOVValueUsed("Exca:SampleTypeLOV", "SampleType", Me![txtSampleType], "Exca: Samples", "Unit Number", "Sample Type", "edit")
    
    If checkValidAction = "ok" Then
        
                'edit action can go ahead - at present simply offer an input box for this
                retval = InputBox("No records refer to this Sample Type (" & Me![txtSampleType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Sample Type that you wish to replace this entry with:", "Enter edited Type")
                If retval <> "" Then
                    Me![txtSampleType] = retval
                
                    retval = InputBox("If there a default amount for this Sample Type (" & Me![txtSampleType] & ")." & Chr(13) & Chr(13) & "Leave this blank if there is not and press ok", "Default Amount")
                    Me![txtSampleAmount] = retval
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

Private Sub Excavation_Click()
On Error GoTo err_close
    
    DoCmd.Close acForm, Me.Name

Exit Sub

err_close:
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
