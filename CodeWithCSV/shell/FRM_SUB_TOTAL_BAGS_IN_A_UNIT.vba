Option Compare Database
Option Explicit

Private Sub cmdRefreshCount_Click()
'requery subform to refresh count of bags
On Error GoTo err_refreshcount
    Me.Requery
    
Exit Sub

err_refreshcount:
    Call General_Error_Trap
    Exit Sub
End Sub
