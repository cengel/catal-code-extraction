Option Compare Database
Option Explicit

Sub Button_OpenForm_ME_Click()
On Error GoTo Err_Button_OpenForm_ME_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:ModifiedEdges"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Button_OpenForm_ME_Click:
    Exit Sub

Err_Button_OpenForm_ME_Click:
    MsgBox Err.Description
    Resume Exit_Button_OpenForm_ME_Click
    
End Sub
Sub Button_OpenForm_Bifaces_Click()
On Error GoTo Err_Button_OpenForm_Bifaces_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:Bifaces"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Button_OpenForm_Bifaces_Click:
    Exit Sub

Err_Button_OpenForm_Bifaces_Click:
    MsgBox Err.Description
    Resume Exit_Button_OpenForm_Bifaces_Click
    
End Sub
