Option Compare Database

Private Sub goto_BasicData_Click()
On Error GoTo Err_goto_BasicData_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim relationexists
    
    stDocName = "Beads: Sheet"
    
    relationexists = DLookup("[GID]", "Beads: Basic Details", "[GID] = '" & Me![GID] & "'")
    If IsNull(relationexists) Then
        'record does not exist
    Else
        'record exists - open it
        stLinkCriteria = "[GID]='" & Me![GID] & "'"
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
    End If
    

Exit_goto_BasicData_Click:
    Exit Sub

Err_goto_BasicData_Click:
    MsgBox Err.Description
    Resume Exit_goto_BasicData_Click
    
End Sub
