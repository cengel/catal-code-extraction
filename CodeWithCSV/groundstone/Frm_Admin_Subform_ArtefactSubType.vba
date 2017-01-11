Option Compare Database
Option Explicit

Private Sub cmdDelete_Click()
'SAJ - check if user can delete this record
' The subtype value is used in the worked data table
' so must check all this to see if is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdDelete_Click
'
'Dim checkValidAction, retVal
'
'    'check basic data for fraction field
'    checkValidAction = CheckIfLOVValueUsed("Groundstone List of Values: Artefact SubType", "Tool SubType", Me![txtSubType], "GroundStone 3: Worked Stone Basics", "GID", "Artefact SubType", "delete", False)
'
'    If checkValidAction = "ok" Then
'                'delete action can go ahead - at present simply offer an input box for this
'                retVal = MsgBox("No records refer to this Sub Type (" & Me![txtSubType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtSubType] & " from the list of available SubTypes?", vbExclamation + vbYesNo, "Confirm Deletion")
'                If retVal = vbYes Then
'                    Me.AllowDeletions = True
'                    DoCmd.RunCommand acCmdDeleteRecord
'                    Me.AllowDeletions = False
'                End If
'
'    ElseIf checkValidAction = "fail" Then
'        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
'    Else
'        MsgBox checkValidAction, vbExclamation, "Action Report"
'    End If
'
'Exit_cmdDelete_Click:
'    Exit Sub
'
'Err_cmdDelete_Click:
'    Call General_Error_Trap
'    Exit Sub
    
End Sub

Private Sub cmdEdit_Click()
'v9.2 SAJ - check if user can edit this record
' The percent value is used in the basic data table
' so must check all this to see if is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdEdit_Click
'
'Dim checkValidAction, checkValidAction2, checkValidAction3, retVal
'
'    'check basic data for fraction field
'    checkValidAction = CheckIfLOVValueUsed("Groundstone List of Values: Artefact SubType", "Tool SubType", Me![txtSubType], "GroundStone 3: Worked Stone Basics", "GID", "Artefact SubType", "edit", False)
'
'    If checkValidAction = "ok" Then
'        'edit action can go ahead - at present simply offer an input box for this
'        retVal = InputBox("No records refer to this SubType (" & Me![txtSubType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Percent that you wish to replace this entry with:", "Enter edited SubType")
'        If retVal <> "" Then
'             Me![txtSubType] = retVal
'        End If
'
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
    
End Sub

'******************************************************
' This subform is new with v9.2
' SAJ v9.2
'******************************************************
Private Sub cmdNewSubType_Click()
'v9.2 - allow new subtype to be added
On Error GoTo err_cmdNewSubType_Click

    If Forms![Frm_Admin_ArtefactTypeSubTypeLOV]![Code] <> "" Then
        Dim sql, retVal
        retVal = InputBox("Please enter the new subtype for the artefact type '" & Forms![Frm_Admin_ArtefactTypeSubTypeLOV]![Tool Types] & "': ", "Enter new subtype")
        If retVal <> "" Then
            sql = "INSERT INTO [Groundstone List of Values: Artefact SubType] ([TypeCode], [Tool SubType]) VALUES (" & Forms![Frm_Admin_ArtefactTypeSubTypeLOV]![Code] & ", '" & retVal & "');"
            DoCmd.RunSQL sql
            Me.Requery
        End If
    Else
        MsgBox "Sorry not all the data necessary to make a new subtype is available.", vbExclamation, "Invalid Action"
    End If
Exit Sub

err_cmdNewSubType_Click:
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
