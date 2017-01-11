Option Compare Database
Option Explicit

Private Sub Ctl_2_AfterUpdate()
'update total
If IsNumeric(Me![Ctl_2]) Then
    Me.Total = Me.Ctl_2 + Me.Ctl2_4 + Me.Ctl_4
Else
    MsgBox "Please enter a numeric value", vbExclamation, "Invalid Entry"
    Me![Ctl_2] = Me![Ctl_2].OldValue
End If
End Sub

Private Sub Ctl_2_GotFocus()
'Me.Total = Me.Ctl_2 + Me.Ctl2_4 + Me.Ctl_4
End Sub

Private Sub Ctl_4_AfterUpdate()
'update total
If IsNumeric(Me![Ctl_4]) Then
    Me.Total = Me.Ctl_2 + Me.Ctl2_4 + Me.Ctl_4
Else
    MsgBox "Please enter a numeric value", vbExclamation, "Invalid Entry"
    Me![Ctl_4] = Me![Ctl_4].OldValue
End If
End Sub

Private Sub Ctl2_4_AfterUpdate()
'update total
If IsNumeric(Me![Ctl2_4]) Then
    Me.Total = Me.Ctl_2 + Me.Ctl2_4 + Me.Ctl_4
Else
    MsgBox "Please enter a numeric value", vbExclamation, "Invalid Entry"
    Me![Ctl2_4] = Me![Ctl2_4].OldValue
End If
End Sub



