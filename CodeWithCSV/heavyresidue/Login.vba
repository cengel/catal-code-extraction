Option Compare Database
'************************************************************************
' This form is new to obtain user login to the database
'
' SAJ v1
'************************************************************************


Private Sub cmdCancel_Click()
'************************************************************************
' Without a user name and password the system cannot run so give the option
' to try again or to quit.
'
' SAJ v9
'************************************************************************
On Error GoTo cmdCancel_Click
Dim retVal

retVal = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
    If retVal = vbYes Then
        MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
        DoCmd.Quit acQuitSaveAll
    End If
    DoCmd.GoToControl "txtLogin"
Exit Sub

cmdCancel_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdOK_Click()
'************************************************************************
' Check both a user name and pwd provided if not prompt user to enter
' If provided feed off to procedure to refresh the link on one table
' with this info to a. check if links are ok, b. ensure access knows the
' users details to allow entry to rest of tables
'
' SAJ v1
'************************************************************************
On Error GoTo cmdOK_Click
Dim retVal

If IsNull(Me![txtLogin]) Or IsNull(Me![txtPwd]) Then
    'entered blank login or pwd double check user wants to cancel
    retVal = MsgBox("Sorry but the system cannot continue without both a login name and a password. Do you want to try again?", vbCritical + vbYesNo, "Login required")
    If retVal = vbYes Then 'try again
        DoCmd.GoToControl "txtLogin"
        Exit Sub
    Else 'no, don't try again so quit system
        retVal = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
        If retVal = vbYes Then
            MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
            DoCmd.Quit acQuitSaveAll
        Else 'no I don't want to quit system, ie: try again
            DoCmd.GoToControl "txtLogin"
            Exit Sub
        End If
    End If
        
Else
    'login and pwd provided
    If LogUserIn(Me![txtLogin], Me![txtPwd]) = True Then
        'function in module TableLinkingProcedures - validate user login
        'if true then login ok and tables accessible - ready to go
        DoCmd.Close acForm, "Login" 'shut form as modal
    Else
        ''OLD: if loginuser = false then the system should have quit by now, this is a catchall
        ''OLD: MsgBox "The system may not of logged into the database correctly. If you encounter any problems please restart the application"
        'ok the login as failed and the user either has asked to try again, or they haven't taken the option to quit so they
        'must remain here until they get it right or quit
    End If
    
End If


Exit Sub

cmdOK_Click:
    Call General_Error_Trap
    DoCmd.Close acForm, "Excavation_Login" 'this may be better as a simply quit the system, will see, however must shut form as modal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'***************************************************************************
' Capture the return key press and action it like all other login boxes
' ie: replicate pressing he ok button.
' Sussed it, you have to set the form method 'Key Preview' to yes to catch it here
'
' SAJ v9.1
'***************************************************************************
On Error Resume Next
'MsgBox KeyAscii
If KeyAscii = 13 Then
    ''MsgBox KeyAscii
    cmdOK_Click
End If
End Sub


