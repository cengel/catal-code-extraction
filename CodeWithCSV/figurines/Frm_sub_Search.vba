Option Compare Database
Option Explicit

Private Sub ID_number_Click()
On Error GoTo err_iddblclick

    DoCmd.OpenForm "Frm_MainData", , , "[ID Number] = '" & Me![ID number] & "'"

Exit Sub

err_iddblclick:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_Number_DblClick(Cancel As Integer)
On Error GoTo err_unitdblclick

    DoCmd.OpenForm "Frm_MainData", , , "[unitnumber] = " & Me![UnitNumber]

Exit Sub

err_unitdblclick:
    Call General_Error_Trap
    Exit Sub
End Sub
