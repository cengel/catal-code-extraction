Option Compare Database
Option Explicit

Private Sub Command6_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Heavy Residue Phase 2"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command6_Click:
    Exit Sub

Err_Command6_Click:
    MsgBox Err.Description
    Resume Exit_Command6_Click
    
End Sub

Private Sub Command7_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Heavy Residue Phase 3"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command7_Click:
    Exit Sub

Err_Command7_Click:
    MsgBox Err.Description
    Resume Exit_Command7_Click
End Sub


Private Sub Command8_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Heavy Summary"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command8_Click:
    Exit Sub

Err_Command8_Click:
    MsgBox Err.Description
    Resume Exit_Command8_Click
End Sub

Sub flot_Click()
On Error GoTo Err_flot_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Flot Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_flot_Click:
    Exit Sub

Err_flot_Click:
    MsgBox Err.Description
    Resume Exit_flot_Click
    
End Sub
Sub light_Ph2_Click()
On Error GoTo Err_light_Ph2_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Residue Phase 2"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_light_Ph2_Click:
    Exit Sub

Err_light_Ph2_Click:
    MsgBox Err.Description
    Resume Exit_light_Ph2_Click
    
End Sub
Sub light_Ph3_Click()
On Error GoTo Err_light_Ph3_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Residue Phase 3"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_light_Ph3_Click:
    Exit Sub

Err_light_Ph3_Click:
    MsgBox Err.Description
    Resume Exit_light_Ph3_Click
    
End Sub
Sub Command5_Click()
On Error GoTo Err_Command5_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Summary"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click
    
End Sub

Private Sub light_summary_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Summary"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_light_summary_Click:
    Exit Sub

Err_light_summary_Click:
    MsgBox Err.Description
    Resume Exit_light_summary_Click
End Sub


