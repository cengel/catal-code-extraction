Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date Changed] = Now()
End Sub


Private Sub go_to_feature_Click()
On Error GoTo Err_go_to_feature_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    
    stLinkCriteria = "[Feature Number]=" & Me![In_feature]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_go_to_feature_Click:
    Exit Sub

Err_go_to_feature_Click:
    MsgBox Err.Description
    Resume Exit_go_to_feature_Click
    
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
