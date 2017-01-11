Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




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
' The level value is used in 3 potential places on the Space Sheet - certain level
' Uncertain level start and uncertain level end so must check all these values to see
' if this level is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdEdit_Click

Dim checkValidAction, checkValidAction2, checkValidAction3, retval

    'check space sheet level field
    checkValidAction = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "Level", "edit")
    
    If checkValidAction = "ok" Then
        ''check space sheet uncertain level start field
        checkValidAction2 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelStart", "edit")
        
        If checkValidAction2 = "ok" Then
        'check space sheet uncertain level end field
            checkValidAction3 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelEnd", "edit")
        
            If checkValidAction3 = "ok" Then
                'edit action can go ahead - at present simply offer an input box for this
                retval = InputBox("No records refer to this Level (" & Me![txtLevel] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Level that you wish to replace this entry with:", "Enter edited Level")
                If retval <> "" Then
                    Me![txtLevel] = retval
                End If
                
            ElseIf checkValidAction3 = "fail" Then
                MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
            Else
                MsgBox checkValidAction3, vbExclamation, "Action Report"
            End If
        ElseIf checkValidAction2 = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            MsgBox checkValidAction2, vbExclamation, "Action Report"
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
' The level value is used in 3 potential places on the Space Sheet - certain level
' Uncertain level start and uncertain level end so must check all these values to see
' if this level is used before allow deletion
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdDelete_Click

Dim checkValidAction, checkValidAction2, checkValidAction3, retval

    'check space sheet level field
    checkValidAction = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "Level", "delete")
    
    If checkValidAction = "ok" Then
        ''check space sheet uncertain level start field
        checkValidAction2 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelStart", "delete")
        
        If checkValidAction2 = "ok" Then
        'check space sheet uncertain level end field
            checkValidAction3 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelEnd", "delete")
        
            If checkValidAction3 = "ok" Then
                'delete action can go ahead - at present simply offer an input box for this
                retval = MsgBox("No records refer to this Level (" & Me![txtLevel] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtLevel] & " from the list of available levels?", vbExclamation + vbYesNo, "Confirm Deletion")
                If retval = vbYes Then
                    Me.AllowDeletions = True
                    DoCmd.RunCommand acCmdDeleteRecord
                    Me.AllowDeletions = False
                End If
                
            ElseIf checkValidAction3 = "fail" Then
                MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
            Else
                MsgBox checkValidAction3, vbExclamation, "Action Report"
            End If
        ElseIf checkValidAction2 = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            MsgBox checkValidAction2, vbExclamation, "Action Report"
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
