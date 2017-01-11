Option Compare Database
Option Explicit
'*********************************************************
' This form is new in v9.1 to give read only view of features
'*********************************************************




Private Sub cmdGoToBuilding_Click()
'***********************************************************
' Open building form with a filter on the number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToBuilding_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Building Sheet"
    
    stLinkCriteria = "[Number]= " & Me![Number]
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
    Exit Sub

Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
