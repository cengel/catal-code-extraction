Option Compare Database
Option Explicit

Private Sub cmdGoToBuilding_Click()
'new 2009 as excavation db too slow when 2 copies of access 2007 open on their machines
On Error GoTo err_cmd

If Not IsNull(Me![Unit Number]) Then
    DoCmd.OpenForm "Frm_pop_UnitDetails_Expanded", acNormal, , "[unit number] = " & Me![Unit Number], acFormReadOnly, acDialog
Else
    MsgBox "No unit number to show", vbInformation, "No unit"
End If



Exit Sub

err_cmd:
    Call General_Error_Trap
    Exit Sub

End Sub
