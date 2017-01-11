Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date Changed] = Now()
End Sub


Sub open_skell_Click()
On Error GoTo Err_open_skell_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Skeleton Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![To_Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_skell_Click:
    Exit Sub

Err_open_skell_Click:
    MsgBox Err.Description
    Resume Exit_open_skell_Click
    
End Sub
