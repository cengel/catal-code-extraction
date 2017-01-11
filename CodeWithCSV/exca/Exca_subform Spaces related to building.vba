Option Compare Database
Option Explicit
'*********************************************************
' This form is new in v9.1 to give read only view of spaces
'*********************************************************

Private Sub cmdGoToSpace_Click()
'***********************************************************
' Open space form with a filter on the space number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToSpace_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Space Sheet"
    
    'if int data type
    stLinkCriteria = "[Space Number]= " & Me![Space number]
    'char datatype
    'stLinkCriteria = "[Space Number]= '" & Me![Space Number] & "'"
    
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly ', acDialog
    'decided against dialog as there may be other windows that can be opened from this form
    'and if this is dialog they appear beneath it
    Exit Sub

Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
