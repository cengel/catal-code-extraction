Option Compare Database
Option Explicit

Private Sub go_to_conservation_Click()

On Error GoTo Err_go_to_conservation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[Conservation Ref]=" & Me![Conservation Ref]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_go_to_conservation_Click:
    Exit Sub

Err_go_to_conservation_Click:
    MsgBox Err.Description
    Resume Exit_go_to_conservation_Click
    
End Sub


Sub Command4_Click()
On Error GoTo Err_Command4_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command4_Click:
    Exit Sub

Err_Command4_Click:
    MsgBox Err.Description
    Resume Exit_Command4_Click
    
End Sub

