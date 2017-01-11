Option Compare Database

Private Sub cmdgotofeature_Click()

On Error GoTo Err_FeatureSheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    stLinkCriteria = "[Feature Number] = " & Me.[latestfeature]

    DoCmd.OpenForm stDocName, , , stLinkCriteria


Exit_FeatureSheet_Click:
    Exit Sub

Err_FeatureSheet_Click:
    Call General_Error_Trap
    Resume Exit_FeatureSheet_Click
End Sub

Private Sub Form_Activate()
Me.Requery
End Sub
