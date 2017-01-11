Option Compare Database
Option Explicit
Private Sub GID_DblClick(Cancel As Integer)
'new season 2008. Allow record to be selected in Finds: Basic Data form
On Error GoTo err_GID

    'docmd.RunCommand acCmdSelectForm "Finds: Basic Data"
    DoCmd.OpenForm "Finds: Basic Data"
    If Forms![Finds: Basic Data]![GID].Enabled = False Then Forms![Finds: Basic Data]![txtUnit].Enabled = True
    DoCmd.GoToControl "GID"
    DoCmd.FindRecord Me![GID]
    'Me![cboFindUnit] = ""

Exit Sub

err_GID:
    Call General_Error_Trap
    Exit Sub

End Sub
