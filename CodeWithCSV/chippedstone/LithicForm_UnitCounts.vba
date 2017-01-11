Option Compare Database
Option Explicit

Sub OpenForm_BagAndUnitDescription_Click()
On Error GoTo Err_OpenForm_BagAndUnitDescription_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:BagAndUnitDescription"
    
    stLinkCriteria = "[Unit]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OpenForm_BagAndUnitDescription_Clic:
    Exit Sub

Err_OpenForm_BagAndUnitDescription_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_BagAndUnitDescription_Clic
    
End Sub
