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

Private Sub cboFind_Click()
On Error GoTo err_cboFind
    
    If Me![cboFind] <> "" Then
        DoCmd.GoToControl "txtSpaceNumber"
        DoCmd.FindRecord Me![cboFind]
   
    End If
Exit Sub

err_cboFind:
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
'v9.2 SAJ - allow deletion of record with warning

On Error GoTo Err_cmdDelete_Click

Dim SUnit, SFeature
Dim retval, msg, msg1

retval = MsgBox("You have selected to delete Space number: " & Me![txtSpaceNumber] & ". The system will now check what additional data exists for this Space and will prompt you again before deleting it." & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbCritical + vbYesNo, "Confirm Action")
If retval = vbYes Then
    
    'check  feature to units
    SUnit = AdminDeletionCheck("Exca: Units in Spaces", "In_Space", Me![txtSpaceNumber], "Related to Unit", "Unit")
    
    'check Feature in spaces
    SFeature = AdminDeletionCheck("Exca: Features in Spaces", "In_Space", Me![txtSpaceNumber], "Related to Feature", "Feature")
    
    If SUnit <> "" Then msg = msg & SUnit & "; "
    If SFeature <> "" Then msg = msg & SFeature & "; "
    
    If msg = "" Then
        msg = "This Space can safely be deleted."
    Else
        msg1 = "This Space has the following relationships that will also be removed by the deletion - " & Chr(13) & Chr(13)
        msg = msg1 & msg
    End If
    
    msg = msg & Chr(13) & Chr(13) & "Are you quite sure that you want to permanently delete Space " & Me![txtSpaceNumber] & "?"
    retval = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retval = vbYes Then
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
        
        If SUnit <> "" Then Call DeleteARecord("Exca: Units in Spaces", "In_Space", Me![txtSpaceNumber], False, mydb)
        If SFeature <> "" Then Call DeleteARecord("Exca: Features in Spaces", "In_Space", Me![txtSpaceNumber], False, mydb)
        
        Call DeleteARecord("Exca: Space Sheet", "Space Number", Me![txtSpaceNumber], False, mydb)
    
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Deletion has been successful"
            Me.Requery
            Me![cboFind].Requery
        Else
            wrkdefault.Rollback
            MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    Else
        MsgBox "Deletion cancelled", vbInformation, "Action Cancelled"
    
    End If
End If
    
    
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Form_Delete(Cancel As Integer)
Call cmdDelete_Click
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
