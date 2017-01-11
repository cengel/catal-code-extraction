Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub
