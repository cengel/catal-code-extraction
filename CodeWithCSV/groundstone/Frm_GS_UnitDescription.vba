Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
'find the unit number
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
    
        If Me.Filter <> "" Then
            If Me.Filter <> "[Unit] = '" & Me![cboFind] & "'" Then
                MsgBox "This form was opened to only show a particular Unit. This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
                Me.FilterOn = False
            End If
        End If
        DoCmd.GoToControl Me![txtUnit].Name
        DoCmd.FindRecord Me![cboFind]
        DoCmd.GoToControl Me![txtComment].Name
   
    End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
On Error GoTo Err_gofirst_Click


    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_gofirst_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoLast_Click()
On Error GoTo Err_goLast_Click


    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_goLast_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoNext_Click()
On Error GoTo Err_goNext_Click


    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_goNext_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoPrev_Click()
On Error GoTo Err_goPrev_Click


    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_goPrev_Click:
    Call General_Error_Trap
    
End Sub

