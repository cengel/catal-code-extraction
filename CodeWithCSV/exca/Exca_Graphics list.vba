Option Compare Database
Option Explicit

Private Sub Close_Click()
On Error GoTo err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
