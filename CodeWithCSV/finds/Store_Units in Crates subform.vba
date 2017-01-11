Option Compare Database
Option Explicit

Private Sub Crate_Number_DblClick(Cancel As Integer)
'new 2008, allow double click to got to crates form to view for this crate
On Error GoTo err_crate

    If Me![Crate Number] <> "" Then
        DoCmd.OpenForm "Store: Crate Register", acNormal, , "[Crate Number] = '" & Me![Crate Number] & "'"
    End If
Exit Sub

err_crate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
On Error GoTo err_Form_Current

If GetGeneralPermissions = "Admin" Then
    '[Store: Units in Crates subform]
    Me.AllowEdits = True
    Me![Crate Letter].Enabled = True
    Me![Crate Letter].Locked = False
    Me![Crate Number].Enabled = True
    Me![Crate Number].Locked = False
    Me![Material].Locked = False
    Me![Material].Enabled = True
    Me![Description].Locked = False
    Me![Description].Enabled = True
Else
    'no permissions to modify crate material and description
End If

Exit Sub

err_Form_Current:
    Call General_Error_Trap
    Exit Sub
End Sub
