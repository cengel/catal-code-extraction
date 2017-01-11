Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date Changed] = Now()
End Sub


Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub

Sub Command5_Click()
On Error GoTo Err_Command5_Click


    DoCmd.GoToRecord , , acLast

Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click
    
End Sub
Sub go_to_unit_Click()
On Error GoTo Err_go_to_unit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_go_to_unit_Click:
    Exit Sub

Err_go_to_unit_Click:
    MsgBox Err.Description
    Resume Exit_go_to_unit_Click
    
End Sub
