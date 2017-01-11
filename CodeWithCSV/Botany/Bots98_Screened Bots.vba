Option Compare Database
Option Explicit

Private Sub go_first_Click()
On Error GoTo Err_go_first_Click

    DoCmd.GoToRecord , , acFirst

Exit_go_first_Click:
    Exit Sub

Err_go_first_Click:
    MsgBox Err.Description
    Resume Exit_go_first_Click
End Sub


Private Sub go_Last_Click()
On Error GoTo Err_go_Last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_Last_Click:
    Exit Sub

Err_go_Last_Click:
    MsgBox Err.Description
    Resume Exit_go_Last_Click
End Sub


Private Sub go_Next_Click()
On Error GoTo Err_go_Next_Click

    DoCmd.GoToRecord , , acNext

Exit_go_Next_Click:
    Exit Sub

Err_go_Next_Click:
    MsgBox Err.Description
    Resume Exit_go_Next_Click
End Sub

Private Sub go_previous_Click()
On Error GoTo Err_go_previous_Click

    DoCmd.GoToRecord , , acPrevious

Exit_go_previous_Click:
    Exit Sub

Err_go_previous_Click:
    MsgBox Err.Description
    Resume Exit_go_previous_Click
End Sub


Private Sub Go_to_button_Click()

    Screen.PreviousControl.SetFocus
    Unit.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Go_to_button_Click:
    Exit Sub

Err_Go_to_button_Click:
    MsgBox Err.Description
    Resume Exit_Go_to_button_Click

End Sub

Private Sub go_to_new_Click()
On Error GoTo Err_go_to_new_Click


    DoCmd.GoToRecord , , acNewRec

Exit_go_to_new_Click:
    Exit Sub

Err_go_to_new_Click:
    MsgBox Err.Description
    Resume Exit_go_to_new_Click
End Sub


