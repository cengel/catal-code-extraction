Option Compare Database
Option Explicit

'
Private Sub Command144_Click()
On Error GoTo Err_Command144_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
'refresh data
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
'go to form
    stDocName = "Bots: Heavy Residue Phase II"
    
    stLinkCriteria = "([Unit]=" & Me![Unit] & " And [Sample]=""" & Me![Sample] & """ And [Flot Number]=" & Me![Flot Number] & ")"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command144_Click:
    Exit Sub

Err_Command144_Click:
    MsgBox Err.Description
    Resume Exit_Command144_Click

End Sub


Private Sub Update_GID()
Me![GID] = Me![Unit] & "." & Me![Sample] & "." & Me![Flot Number]
End Sub


Private Sub Flot_Number_AfterUpdate()
Update_GID
End Sub

Private Sub Flot_Number_Change()
Update_GID
'Forms![Bots: Flot Sheet].Refresh
End Sub






Sub Open_HR_Phase_II_Click()
On Error GoTo Err_Open_HR_Phase_II_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    ' refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    
    'go to record:
    stDocName = "Bots98: Heavy Residue Phase 2"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
    'close this form
    DoCmd.Close acForm, "Bots98: Flot Sheet"
    
Exit_Open_HR_Phase_II_Click:
    Exit Sub

Err_Open_HR_Phase_II_Click:
    MsgBox Err.Description
    Resume Exit_Open_HR_Phase_II_Click
    
End Sub
Sub Go_to_button_Click()
On Error GoTo Err_Go_to_button_Click


    Screen.PreviousControl.SetFocus
    Unit.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Go_to_button_Click:
    Exit Sub

Err_Go_to_button_Click:
    MsgBox Err.Description
    Resume Exit_Go_to_button_Click
    
End Sub
Sub refresh_Click()
On Error GoTo Err_refresh_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_refresh_Click:
    Exit Sub

Err_refresh_Click:
    MsgBox Err.Description
    Resume Exit_refresh_Click
    
End Sub


Private Sub Open_LR_Phase_II_Click()
On Error GoTo Err_Open_LR_Phase_II_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
' refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    stDocName = "Bots98: Light Residue Phase 2"
'go to record
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
'close this form
    DoCmd.Close acForm, "Bots98: Flot Sheet"
    
Exit_Open_LR_Phase_II_Click:
    Exit Sub

Err_Open_LR_Phase_II_Click:
    MsgBox Err.Description
    Resume Exit_Open_LR_Phase_II_Click
End Sub


Sub go_Next_Click()
On Error GoTo Err_go_Next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_Next_Click:
    Exit Sub

Err_go_Next_Click:
    MsgBox Err.Description
    Resume Exit_go_Next_Click
    
End Sub
Sub go_Last_Click()
On Error GoTo Err_go_Last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_Last_Click:
    Exit Sub

Err_go_Last_Click:
    MsgBox Err.Description
    Resume Exit_go_Last_Click
    
End Sub
Sub go_to_new_Click()
On Error GoTo Err_go_to_new_Click


    DoCmd.GoToRecord , , acNewRec

Exit_go_to_new_Click:
    Exit Sub

Err_go_to_new_Click:
    MsgBox Err.Description
    Resume Exit_go_to_new_Click
    
End Sub
Sub go_previous_Click()
On Error GoTo Err_go_previous_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous_Click:
    Exit Sub

Err_go_previous_Click:
    MsgBox Err.Description
    Resume Exit_go_previous_Click
    
End Sub
Sub go_first_Click()
On Error GoTo Err_go_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_first_Click:
    Exit Sub

Err_go_first_Click:
    MsgBox Err.Description
    Resume Exit_go_first_Click
    
End Sub

Private Sub Sample_AfterUpdate()
Update_GID
End Sub

Private Sub Sample_Change()
Update_GID
End Sub


Private Sub Unit_AfterUpdate()

Update_GID
'Me![GID] = Me![Unit] & "." & Me![Sample] & "." & Me![Flot Number]

End Sub


Private Sub Unit_Change()
Update_GID
End Sub


Private Sub Unit_Enter()
'SAJ pre version numbers - related to security RO - calling
'the update on enter when nothing has changed is causing a SQL update
'permissions error which will really confuse the user! taken out
'11/01/06
'Update_GID
End Sub


Sub Exit_Click()
On Error GoTo Err_Exit_Click


    DoCmd.Close

Exit_Exit_Click:
    Exit Sub

Err_Exit_Click:
    MsgBox Err.Description
    Resume Exit_Exit_Click
    
End Sub
Sub light_sum_Click()
On Error GoTo Err_light_sum_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
' refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    
    stDocName = "Bots98: Light Summary"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_light_sum_Click:
    Exit Sub

Err_light_sum_Click:
    MsgBox Err.Description
    Resume Exit_light_sum_Click
    
End Sub
Sub Command159_Click()
On Error GoTo Err_Command159_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots: Heavy Residue Summary display"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command159_Click:
    Exit Sub

Err_Command159_Click:
    MsgBox Err.Description
    Resume Exit_Command159_Click
    
End Sub
Sub heavy_sum_Click()
On Error GoTo Err_heavy_sum_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
' refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    
    stDocName = "Bots98: Heavy Summary"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_heavy_sum_Click:
    Exit Sub

Err_heavy_sum_Click:
    MsgBox Err.Description
    Resume Exit_heavy_sum_Click
    
End Sub
Sub Command180_Click()
On Error GoTo Err_Command180_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Residue Phase 2"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command180_Click:
    Exit Sub

Err_Command180_Click:
    MsgBox Err.Description
    Resume Exit_Command180_Click
    
End Sub
Sub Screened_Click()
On Error GoTo Err_Screened_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    'go to record
    stDocName = "Bots98: Screened Bots"
    
    stLinkCriteria = "[Unit]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
        
    'close this form
    DoCmd.Close acForm, "Bots98: Flot Sheet"


Exit_Screened_Click:
    Exit Sub

Err_Screened_Click:
    MsgBox Err.Description
    Resume Exit_Screened_Click
    
End Sub
Sub Screened_Bots_Click()
On Error GoTo Err_Screened_Bots_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Screened Bots"
    
    stLinkCriteria = "[Unit]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Screened_Bots_Click:
    Exit Sub

Err_Screened_Bots_Click:
    MsgBox Err.Description
    Resume Exit_Screened_Bots_Click
    
End Sub
