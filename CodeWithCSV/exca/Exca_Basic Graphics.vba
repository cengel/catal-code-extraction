Option Compare Database
Option Explicit

Private Sub open_details_Click()
On Error GoTo Err_open_details_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    Me.refresh
    stDocName = "Exca: Graphics list"
    
    
    stLinkCriteria = "[Graphic Number]=" & "'" & Me![Graphic Number] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_details_Click:
    Exit Sub

Err_open_details_Click:
    MsgBox Err.Description
    Resume Exit_open_details_Click
    
End Sub
Private Sub find_graph_Click()
On Error GoTo Err_find_graph_Click


    Forms![Exca: Basic Graphics].[Graphic Number].SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_graph_Click:
    Exit Sub

Err_find_graph_Click:
    MsgBox Err.Description
    Resume Exit_find_graph_Click
    
End Sub
