Option Compare Database
Option Explicit

Sub cmd_gotodebitage_Click()
On Error GoTo Err_cmd_gotodebitage_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:Debitage"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
    DoCmd.Close
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmd_gotodebitage_Click:
    Exit Sub

Err_cmd_gotodebitage_Click:
    MsgBox Err.Description
    Resume Exit_cmd_gotodebitage_Click
    
End Sub
