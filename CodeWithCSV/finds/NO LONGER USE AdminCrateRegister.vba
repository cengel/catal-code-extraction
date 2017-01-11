Option Compare Database
Option Explicit

Private Sub find_unit_Click()
On Error GoTo Err_find_unit_Click


    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim message As String, title As String, Unit As String, default As String

message = "Enter a unit number"   ' Set prompt.
title = "Searching Crate Register" ' Set title.
default = "1000"   ' Set default.
' Display message, title, and default value.
Unit = InputBox(message, title, default)

    stDocName = "Store: Find Unit in Crate"
    stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_find_unit_Click:
    Exit Sub

Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
    
End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

Me![Date Changed] = Now()

End Sub

Sub GoFirst_Click()
On Error GoTo Err_GoFirst_Click


    DoCmd.GoToRecord , , acFirst

Exit_GoFirst_Click:
    Exit Sub

Err_GoFirst_Click:
    MsgBox Err.Description
    Resume Exit_GoFirst_Click
    
End Sub
Sub Previous_Click()
On Error GoTo Err_Previous_Click


    DoCmd.GoToRecord , , acPrevious

Exit_Previous_Click:
    Exit Sub

Err_Previous_Click:
    MsgBox Err.Description
    Resume Exit_Previous_Click
    
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
