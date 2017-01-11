Option Compare Database
Option Explicit

Sub first_Click()
On Error GoTo Err_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_first_Click:
    Exit Sub

Err_first_Click:
    MsgBox Err.Description
    Resume Exit_first_Click
    
End Sub

Private Sub Form_Current()

If Me![Finds: Basic Data.GID] <> "" Then
Me![finds].Enabled = True
Else
Me![finds].Enabled = False
End If

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
Sub Close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
Sub finds_Click()
On Error GoTo Err_finds_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Finds: Basic Data"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_finds_Click:
    Exit Sub

Err_finds_Click:
    MsgBox Err.Description
    Resume Exit_finds_Click
    
End Sub
