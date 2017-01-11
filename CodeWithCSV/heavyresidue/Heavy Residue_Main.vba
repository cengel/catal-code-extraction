Option Compare Database
Option Explicit

Sub open_HR_Click()
On Error GoTo Err_open_HR_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Heavy Residue: Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_HR_Click:
    Exit Sub

Err_open_HR_Click:
    MsgBox Err.Description
    Resume Exit_open_HR_Click
    
End Sub
