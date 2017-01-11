Option Compare Database

Private Sub cmdPrint_Click()
On Error GoTo Err_cmdPrint_Click


    DoCmd.PrintOut

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub
