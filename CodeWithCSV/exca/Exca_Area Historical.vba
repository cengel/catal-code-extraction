Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub cmdClose_Click()
'second close button to be obvious as opened as dialog
    Excavation_Click
End Sub

'********************************************************************
' This whole form is new in v9.1
'********************************************************************



Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    'Dim stDocName As String
    'Dim stLinkCriteria As String

    'stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Area Historical"
Exit Sub

err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub



