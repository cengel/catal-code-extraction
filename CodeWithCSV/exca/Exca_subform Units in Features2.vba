Option Compare Database
Option Explicit
'********************************************************
' This form which is used on Exca: Feature Sheet as a
' subform is now read-only there - so no need to have
' code processing - its commented out apart from button
'
' SAJ v9.1
'********************************************************

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Me![Date changed] = Now()
End Sub


Private Sub Unit_AfterUpdate()
'Me.Requery
'DoCmd.GoToRecord , , acLast
End Sub

Sub Command5_Click()
'On Error GoTo Err_Command5_Click
'
'
'    DoCmd.GoToRecord , , acLast
'
'Exit_Command5_Click:
'    Exit Sub
'
'Err_Command5_Click:
'    MsgBox Err.Description
'    Resume Exit_Command5_Click
'
End Sub
Sub go_to_unit_Click()
'********************************************
'Existing code for go to unit button, added
'general error trap, now open readonly
'
'SAJ v9.1
'********************************************
On Error GoTo Err_go_to_unit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormReadOnly

    Exit Sub

Err_go_to_unit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
