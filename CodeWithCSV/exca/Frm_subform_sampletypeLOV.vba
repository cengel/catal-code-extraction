Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub
