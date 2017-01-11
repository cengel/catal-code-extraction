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

    stDocName = "frm_pop_phases_in_SpaceBuilding"
    
    stLinkCriteria = "[PhaseInBuilding]= '" & Me![SpacePhase] & "'"
    
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormPropertySettings
    
    Exit Sub

Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Delete(Cancel As Integer)
'must check that no units are associated with this phase before allow delete
On Error GoTo err_delete

    Dim checkit
    checkit = DCount("[Unit Number]", "[Exca: Unit Sheet]", "[PhaseInBuilding] = '" & Me!SpacePhase & "'")
    If checkit > 0 Then
        MsgBox "Units are associated with this Phase. It cannot be deleted as it is in use", vbInformation, "Action Cancelled"
        Cancel = True
    End If
Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub
