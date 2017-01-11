Option Compare Database



Private Sub Find_Number_Change()

End Sub

Private Sub Unit_AfterUpdate()

End Sub


Private Sub Unit_Change()

End Sub


Private Sub Unit_Enter()

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
On Error GoTo Err_find_Click


    Screen.PreviousControl.SetFocus
    'GID.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_Click:
    Exit Sub

Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub


Private Sub search_Click()
On Error GoTo Err_search_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.Close
    stDocName = "Log: Query Functions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_search_Click:
    Exit Sub

Err_search_Click:
    MsgBox Err.Description
    Resume Exit_search_Click
    
End Sub
Private Sub find_date_Click()
On Error GoTo Err_find_date_Click

    Me![Date].SetFocus
    'Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_date_Click:
    Exit Sub

Err_find_date_Click:
    MsgBox Err.Description
    Resume Exit_find_date_Click
    
End Sub
Private Sub find_area_Click()
On Error GoTo Err_find_area_Click

    Me![Area].SetFocus
    'Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_area_Click:
    Exit Sub

Err_find_area_Click:
    MsgBox Err.Description
    Resume Exit_find_area_Click
    
End Sub
Private Sub Command62_Click()
On Error GoTo Err_Command62_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Log: Search Functions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command62_Click:
    Exit Sub

Err_Command62_Click:
    MsgBox Err.Description
    Resume Exit_Command62_Click
    
End Sub
