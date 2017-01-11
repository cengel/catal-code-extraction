Option Compare Database
Option Explicit

Sub copy_layer_Click()
On Error GoTo Err_copy_layer_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Copy Layer description"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_copy_layer_Click:
    Exit Sub

Err_copy_layer_Click:
    MsgBox Err.Description
    Resume Exit_copy_layer_Click
    
End Sub
