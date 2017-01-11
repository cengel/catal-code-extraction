Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
'close this new 2009 form
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open

Me.Caption = "Further Units Details for Unit:" & Me![Unit Number]

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
