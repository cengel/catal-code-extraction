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
'SAJ - close the form
    DoCmd.Close acForm, Me.Name
End Sub



Private Sub cmdEdit_Click()
'v9.2 SAJ - check if user can edit this record
' The fraction value is used in the basic data table
' so must check all this to see if is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdEdit_Click
'
'Dim checkValidAction, checkValidAction2, checkValidAction3, retVal
'
'    'check basic data for fraction field
'    checkValidAction = CheckIfLOVValueUsed("GroundStone List of Values: FlotFraction", "FlotFraction", Me![txtFraction], "GroundStone 1: Basic Data", "GID", "Fraction", "edit", False)
'
'    If checkValidAction = "ok" Then
'        'edit action can go ahead - at present simply offer an input box for this
'        retVal = InputBox("No records refer to this Fraction (" & Me![txtFraction] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Fraction that you wish to replace this entry with:", "Enter edited Fraction")
'        If retVal <> "" Then
'             Me![txtFraction] = retVal
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

'Err_cmdEdit_Click:
'    Call General_Error_Trap
'    Exit Sub
    
End Sub

Private Sub cmdDelete_Click()
'SAJ - check if user can delete this record
' The fraction value is used in the basic data table
' so must check all this to see if is used before allow delete
' At present not offering global edits - this can be extended to offer this if required

'On Error GoTo Err_cmdDelete_Click
'
'Dim checkValidAction, checkValidAction2, checkValidAction3, retVal
'
'    'check basic data for fraction field
'    checkValidAction = CheckIfLOVValueUsed("GroundStone List of Values: FlotFraction", "FlotFraction", Me![txtFraction], "GroundStone 1: Basic Data", "GID", "Fraction", "delete", False)
'
'    If checkValidAction = "ok" Then
'                'delete action can go ahead - at present simply offer an input box for this
'                retVal = MsgBox("No records refer to this Fraction (" & Me![txtFraction] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFraction] & " from the list of available fractions?", vbExclamation + vbYesNo, "Confirm Deletion")
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
