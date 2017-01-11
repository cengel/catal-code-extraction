Option Compare Database

Private Sub cmdgotounit_Click()
On Error GoTo Err_UnitSheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    stLinkCriteria = "[Unit Number] = " & Me.[Unit_Number]

    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
   '' REMOVED SAJ AS WILL FAIL FOR RO USERS DoCmd.GoToRecord acForm, stDocName, acNewRec

Exit_UnitSheet_Click:
    Exit Sub

Err_UnitSheet_Click:
    Call General_Error_Trap
    Resume Exit_UnitSheet_Click
End Sub
