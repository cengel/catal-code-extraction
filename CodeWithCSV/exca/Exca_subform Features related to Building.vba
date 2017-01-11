Option Compare Database
Option Explicit
'*********************************************************
' This form is new in v9.1 to give read only view of features
'*********************************************************


Private Sub cmdgotofeature_Click()
'***********************************************************
' Open feature form with a filter on the feature number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToFeature_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    
    stLinkCriteria = "[Feature Number]= " & Me![Feature Number]
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly ', acDialog
    'decided against dialog as there may be other windows that can be opened from this form
    'and if this is dialog they appear beneath it
    Exit Sub

Err_cmdGoToFeature_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
