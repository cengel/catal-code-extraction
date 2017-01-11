Option Compare Database
Option Explicit



Private Sub open_dendro_Click()
On Error GoTo Err_open_dendro_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Anthracology: Dendro"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_dendro_Click:
    Exit Sub

Err_open_dendro_Click:
    MsgBox Err.description
    Resume Exit_open_dendro_Click
    
End Sub

Private Sub open_dendroHandpicked_Click()
On Error GoTo Err_open_dendroHandpicked_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Anthracology: Dendro_handpicked"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_dendroHandpicked_Click:
    Exit Sub

Err_open_dendroHandpicked_Click:
    MsgBox Err.description
End Sub

Sub open_HR_Click()
On Error GoTo Err_open_HR_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClayObjects: Basic"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_HR_Click:
    Exit Sub

Err_open_HR_Click:
    MsgBox Err.description
    Resume Exit_open_HR_Click
    
End Sub

Private Sub openLevelOne_Click()
On Error GoTo Err_openLevelOne_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClayObjects: LevelOne"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_openLevelOne_Click:
    Exit Sub

Err_openLevelOne_Click:
    MsgBox Err.description
    Resume Exit_openLevelOne_Click
    
End Sub
