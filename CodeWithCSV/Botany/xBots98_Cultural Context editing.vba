Option Compare Database

Sub Command5_Click()
On Error GoTo Err_Command5_Click

Dim ComboBox As Control
Set ComboBox = Forms![Bots98: Flot Sheet]![Cultural Context Code]
ComboBox.Requery

DoCmd.Close

Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click
    
End Sub
Sub Command6_Click()
On Error GoTo Err_Command6_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_Command6_Click:
    Exit Sub

Err_Command6_Click:
    MsgBox Err.Description
    Resume Exit_Command6_Click
    
End Sub
Sub Command7_Click()
On Error GoTo Err_Command7_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_Command7_Click:
    Exit Sub

Err_Command7_Click:
    MsgBox Err.Description
    Resume Exit_Command7_Click
    
End Sub
