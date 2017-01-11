Option Compare Database
Option Explicit

Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_find

If Me!cboFindUnit <> "" Then
    DoCmd.GoToControl "UnitNumber"
    DoCmd.FindRecord Me!cboFindUnit
    Me!cboFindUnit = ""
End If

Exit Sub

err_find:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdNew_Click()
On Error GoTo err_new

    DoCmd.RunCommand acCmdRecordsGoToNew
    DoCmd.GoToControl "UnitNumber"

Exit Sub

err_new:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdPrintPrioritySheet_Click()
On Error GoTo err_print

        DoCmd.OpenReport "PrintPriorityComments", acViewPreview, , "[UnitNumber] = " & Me![UnitNumber]

Exit Sub

err_print:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    Exit Sub

End Sub
