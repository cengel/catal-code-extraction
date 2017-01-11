Option Compare Database
Option Explicit

Private Sub Unit_Number_DblClick(Cancel As Integer)
On Error GoTo err_unitdblclick

    DoCmd.OpenForm "Exca: Unit Sheet", , , "[unit number] = " & Me![Unit Number]

Exit Sub

err_unitdblclick:
    Call General_Error_Trap
    Exit Sub
End Sub
