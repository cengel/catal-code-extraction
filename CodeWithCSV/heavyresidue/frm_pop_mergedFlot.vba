Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub ContainsFlot_AfterUpdate()
'new 2011
On Error GoTo err_contains

    Me![FlotRecordedInHR] = Me!txtFlot
Exit Sub

err_contains:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'new 2011
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    Me![txtFlot] = Me.OpenArgs
    Me!lblFlot.Caption = "Flot number " & Me![txtFlot] & " consists of Flot numbers:"
    
    Me.RecordSource = "SELECT * FROM [Heavy Residue: Flot Merge Log] WHERE [FlotRecordedInHR] = " & Me![txtFlot] & ";"
    
End If
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub
Private Sub Command26_Click()
On Error GoTo Err_Command26_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command26_Click:
    Exit Sub

Err_Command26_Click:
    MsgBox Err.Description
    Resume Exit_Command26_Click
    
End Sub
