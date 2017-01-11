Option Compare Database
Option Explicit



Private Sub open_dendro_Click()
On Error GoTo Err_open_dendro_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Beads: Dendro"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_dendro_Click:
    Exit Sub

Err_open_dendro_Click:
    MsgBox Err.Description
    Resume Exit_open_dendro_Click
    
End Sub

Private Sub open_dendroHandpicked_Click()
On Error GoTo Err_open_dendroHandpicked_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Beads: Dendro_handpicked"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_dendroHandpicked_Click:
    Exit Sub

Err_open_dendroHandpicked_Click:
    MsgBox Err.Description
End Sub

Sub open_HR_Click()
On Error GoTo Err_open_HR_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Beads: Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_HR_Click:
    Exit Sub

Err_open_HR_Click:
    MsgBox Err.Description
    Resume Exit_open_HR_Click
    
End Sub
