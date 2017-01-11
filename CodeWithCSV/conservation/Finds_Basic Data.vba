Option Compare Database


Private Sub Update_GID()
Me![GID] = Me![Unit] & "." & Me![Find Number]
'Me.Refresh
End Sub

Private Sub Find_Number_AfterUpdate()
Update_GID
Forms![Finds: Basic Data].Refresh
End Sub

Private Sub Find_Number_Change()
Update_GID
End Sub

Private Sub Form_Current()
If IsNull([Finds: subform Conservation].Form![Conservation Ref].Value) Then
[Finds: subform Conservation].Form!Command4.Enabled = False
Else
[Finds: subform Conservation].Form!Command4.Enabled = True
End If

End Sub

Private Sub Unit_AfterUpdate()
Update_GID
Forms![Finds: Basic Data].Refresh
End Sub


Private Sub Unit_Change()
Update_GID
End Sub


Private Sub Unit_Enter()
'SAJ before versioning - this causes sql update error to be returned to user even
'they have not tried to edit anything, most confusing and unnecessary so removed
' 11/01/05
'Update_GID
End Sub


Sub first_Click()
On Error GoTo Err_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_first_Click:
    Exit Sub

Err_first_Click:
    MsgBox Err.Description
    Resume Exit_first_Click
    
End Sub
Sub prev_Click()
On Error GoTo Err_prev_Click


    DoCmd.GoToRecord , , acPrevious

Exit_prev_Click:
    Exit Sub

Err_prev_Click:
    MsgBox Err.Description
    Resume Exit_prev_Click
    
End Sub
Sub next_Click()
On Error GoTo Err_next_Click


    DoCmd.GoToRecord , , acNext

Exit_next_Click:
    Exit Sub

Err_next_Click:
    MsgBox Err.Description
    Resume Exit_next_Click
    
End Sub
Sub last_Click()
On Error GoTo Err_last_Click


    DoCmd.GoToRecord , , acLast

Exit_last_Click:
    Exit Sub

Err_last_Click:
    MsgBox Err.Description
    Resume Exit_last_Click
    
End Sub
Sub new_Click()
On Error GoTo Err_new_Click


    DoCmd.GoToRecord , , acNewRec

Exit_new_Click:
    Exit Sub

Err_new_Click:
    MsgBox Err.Description
    Resume Exit_new_Click
    
End Sub
Sub closeCommand45_Click()
On Error GoTo Err_closeCommand45_Click


    DoCmd.Close

Exit_closeCommand45_Click:
    Exit Sub

Err_closeCommand45_Click:
    MsgBox Err.Description
    Resume Exit_closeCommand45_Click
    
End Sub
Sub find_Click()
On Error GoTo err_find_Click


    Screen.PreviousControl.SetFocus
    GID.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_Click:
    Exit Sub

err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub
Sub cons_Click()
On Error GoTo Err_cons_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[Conserv: Basic Record.GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cons_Click:
    Exit Sub

Err_cons_Click:
    MsgBox Err.Description
    Resume Exit_cons_Click
    
End Sub
Sub conservation_Click()
On Error GoTo Err_conservation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[GID1]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord acForm, stDocName, acGoTo, acLast

Exit_conservation_Click:
    Exit Sub

Err_conservation_Click:
    MsgBox Err.Description
    Resume Exit_conservation_Click
    
End Sub
