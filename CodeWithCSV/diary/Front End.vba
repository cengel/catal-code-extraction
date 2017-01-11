Option Compare Database   'Use database order for string comparisons

Private Sub Command26_Click()

End Sub


Private Sub Excavation_Click()
On Error GoTo Err_Excavation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Diary Form"
    
Exit_Excavation_Click:
    Exit Sub

Err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub

Private Sub Master_Control_Click()
On Error GoTo Err_Master_Control_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Diary Form"
    
Exit_Master_Control_Click:
    Exit Sub

Err_Master_Control_Click:
    MsgBox Err.Description
    Resume Exit_Master_Control_Click
End Sub

Sub New_Diary_Entry_Click()
On Error GoTo Err_New_Diary_Entry_Click

    DoCmd.GoToRecord , , acNewRec

Exit_New_Diary_Entry_Click:
    Exit Sub

Err_New_Diary_Entry_Click:
    MsgBox Err.Description
    Resume Exit_New_Diary_Entry_Click
    
End Sub
Sub Diary_Go_to_New_Click()
On Error GoTo Err_Diary_Go_to_New_Click


    DoCmd.GoToRecord , , acNewRec

Exit_Diary_Go_to_New_Click:
    Exit Sub

Err_Diary_Go_to_New_Click:
    MsgBox Err.Description
    Resume Exit_Diary_Go_to_New_Click
    
End Sub
Sub New_Diary_Entry2_Click()
On Error GoTo Err_New_Diary_Entry2_Click


    New_Diary_Entry_Click

Exit_New_Diary_Entry2_Click:
    Exit Sub

Err_New_Diary_Entry2_Click:
    MsgBox Err.Description
    Resume Exit_New_Diary_Entry2_Click
    
End Sub
Sub find_Click()
On Error GoTo Err_find_Click


    Screen.PreviousControl.SetFocus
    Me![Diary].SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_Click:
    Exit Sub

Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub
Sub close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
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
Sub prev_Click()
On Error GoTo Err_prev_Click


    DoCmd.GoToRecord , , acPrevious

Exit_prev_Click:
    Exit Sub

Err_prev_Click:
    MsgBox Err.Description
    Resume Exit_prev_Click
    
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
Private Sub Diary_Entry_Form_Click()
On Error GoTo Err_Diary_Entry_Form_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Diary Form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Diary_Entry_Form_Click:
    Exit Sub

Err_Diary_Entry_Form_Click:
    MsgBox Err.Description
    Resume Exit_Diary_Entry_Form_Click
    
End Sub
Private Sub Quit_Diary_System_Click()
On Error GoTo Err_Quit_Diary_System_Click


    DoCmd.Quit

Exit_Quit_Diary_System_Click:
    Exit Sub

Err_Quit_Diary_System_Click:
    MsgBox Err.Description
    Resume Exit_Quit_Diary_System_Click
    
End Sub
Private Sub Diary_Reports_Click()
On Error GoTo Err_Diary_Reports_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Report Selection Form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Diary_Reports_Click:
    Exit Sub

Err_Diary_Reports_Click:
    MsgBox Err.Description
    Resume Exit_Diary_Reports_Click
    
End Sub
Private Sub Command55_Click()
On Error GoTo Err_Command55_Click


    DoCmd.GoToRecord , , acLast

Exit_Command55_Click:
    Exit Sub

Err_Command55_Click:
    MsgBox Err.Description
    Resume Exit_Command55_Click
    
End Sub
