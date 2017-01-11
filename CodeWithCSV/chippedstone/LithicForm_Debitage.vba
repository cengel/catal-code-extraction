Option Compare Database
Option Explicit

Sub Button_CloseForm_Click()
On Error GoTo Err_Button_CloseForm_Click


    DoCmd.Close

Exit_Button_CloseForm_Click:
    Exit Sub

Err_Button_CloseForm_Click:
    MsgBox Err.Description
    Resume Exit_Button_CloseForm_Click
    
End Sub
Sub Button_OpenFormProxend_Click()
On Error GoTo Err_Button_OpenFormProxend_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:ProximalEnds"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Button_OpenFormProxend_Click:
    Exit Sub

Err_Button_OpenFormProxend_Click:
    MsgBox Err.Description
    Resume Exit_Button_OpenFormProxend_Click
    
End Sub
Sub Button_OpenForm_Blades_Click()
On Error GoTo Err_Button_OpenForm_Blades_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:Blades"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Button_OpenForm_Blades_Click:
    Exit Sub

Err_Button_OpenForm_Blades_Click:
    MsgBox Err.Description
    Resume Exit_Button_OpenForm_Blades_Click
    
End Sub
Sub Buttons_OpenForm_Cores_Click()
On Error GoTo Err_Buttons_OpenForm_Cores_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:Cores"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Buttons_OpenForm_Cores_Click:
    Exit Sub

Err_Buttons_OpenForm_Cores_Click:
    MsgBox Err.Description
    Resume Exit_Buttons_OpenForm_Cores_Click
    
End Sub
Sub Button_OpenForm_MB_Click()
On Error GoTo Err_Button_OpenForm_MB_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:ModifiedBlanks"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
        DoCmd.Minimize
        DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Button_OpenForm_MB_Click:
    Exit Sub

Err_Button_OpenForm_MB_Click:
    MsgBox Err.Description
    Resume Exit_Button_OpenForm_MB_Click
    
End Sub
Sub button_goto_bagandunit_Click()
On Error GoTo Err_button_goto_bagandunit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:BagAndUnitDescription"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]
    DoCmd.Close
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_goto_bagandunit_Click:
    Exit Sub

Err_button_goto_bagandunit_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_bagandunit_Click
    
End Sub
