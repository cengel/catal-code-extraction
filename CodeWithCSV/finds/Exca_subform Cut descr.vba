Option Compare Database
Option Explicit

Sub copy_cut_Click()
On Error GoTo Err_copy_cut_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Copy Cut description"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_copy_cut_Click:
    Exit Sub

Err_copy_cut_Click:
    MsgBox Err.Description
    Resume Exit_copy_cut_Click
    
End Sub
