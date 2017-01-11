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

Private Sub Unit_AfterUpdate()
Update_GID
Forms![Finds: Basic Data].Refresh
End Sub


Private Sub Unit_Change()
Update_GID
End Sub


Private Sub Unit_Enter()
Update_GID
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
Sub Close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
Sub find_Click()
On Error GoTo Err_find_Click

    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
   
Exit_find_Click:
    Exit Sub

Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
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
