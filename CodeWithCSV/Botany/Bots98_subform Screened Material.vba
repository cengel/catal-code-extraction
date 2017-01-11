Option Compare Database
Option Explicit

Sub Phase_3_Click()
On Error GoTo Err_Phase_3_Click
    
    Dim stDocName As String
    Dim stLinkCriteria As String
' Refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    stDocName = "Bots98: Heavy Residue Phase 3"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "' AND [Material]=" & "'" & Me![Material] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Phase_3_Click:
    Exit Sub

Err_Phase_3_Click:
    MsgBox Err.Description
    Resume Exit_Phase_3_Click
    
End Sub
