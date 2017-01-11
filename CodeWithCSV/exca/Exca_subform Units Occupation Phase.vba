Option Compare Database
Option Explicit
'**********************************************************
' This subform is new in version 9.2 - as a feature can be
' in many spaces the space field has been removed from the
' Feature tables and normalised out int Exca: Features in Spaces
' SAJ v9.2
'**********************************************************



Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click

'2010 make it easier to delete rather than using pop up
Dim resp
resp = MsgBox("Are you sure you want to delete the phasing " & Me![txtOccupationPhase] & " from this unit?", vbQuestion + vbYesNo, "Confirm Deletion")
If resp = vbYes Then
    Me![txtOccupationPhase].Locked = False
    DoCmd.RunCommand acCmdDeleteRecord
    Me![txtOccupationPhase].Locked = True
End If

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub
