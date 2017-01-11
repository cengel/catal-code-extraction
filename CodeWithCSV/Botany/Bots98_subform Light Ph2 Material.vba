Option Compare Database
Option Explicit

Sub Phase_3_Click()
On Error GoTo Err_Phase_3_Click

'
'    Dim stDocName As String
'    Dim stLinkCriteria As String
'' refresh
'    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
'    stDocName = "Bots98: Light Residue Phase 2"
''go to record
'    'stLinkCriteria = "([Unit]=" & Me![Unit] & " And [Sample]=""" & Me![Sample] & """ And [Flot Number]=" & Me![Flot Number] & ")"
'    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
''close this form
'    DoCmd.Close acForm, "Bots98: Flot Sheet"
    
    Dim stDocName As String
    Dim stLinkCriteria As String
' Refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    stDocName = "Bots98: Light Residue Phase 3"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "' AND [Material]=" & "'" & Me![Material] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Phase_3_Click:
    Exit Sub

Err_Phase_3_Click:
    MsgBox Err.Description
    Resume Exit_Phase_3_Click
    
End Sub
