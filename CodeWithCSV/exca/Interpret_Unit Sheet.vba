Option Compare Database
Option Explicit

Sub Go_first_Click()
On Error GoTo Err_Go_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_Go_first_Click:
    Exit Sub

Err_Go_first_Click:
    MsgBox Err.Description
    Resume Exit_Go_first_Click
    
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
Sub find_unit_Click()
On Error GoTo Err_find_unit_Click

    Screen.PreviousControl.SetFocus
    Unit_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_unit_Click:
    Exit Sub

Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
    
End Sub
Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
    
End Sub
Sub go_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
    
End Sub
Sub Close_Click()
On Error GoTo err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
