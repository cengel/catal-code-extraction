Option Compare Database
Option Explicit
Const RecSource = "SELECT * FROM [Store: Crate Movement by Teams]"

Private Sub cboCrate_AfterUpdate()
'new 2010
On Error GoTo err_cboCrate

    If Me![cboCrate] <> "" Then
        If Me![cboFindTeam] <> "" And Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate] & ", " & Me![cboCrate]
        ElseIf Me![cboFindTeam] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboCrate]
        ElseIf Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate] & ", " & Me![cboCrate]
        Else
            Me.RecordSource = RecSource & " WHERE ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboCrate]
        End If
        'Me.Requery
    Else
        'no crate info but keep other details if there
        If Me![cboFindTeam] <> "" And Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate]
        ElseIf Me![cboFindTeam] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "'"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam]
        ElseIf Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate]
        Else
            Me.RecordSource = RecSource
            Me!lblFilter.caption = "Current Filter: none"
        End If
    End If

Exit Sub

err_cboCrate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDate_AfterUpdate()
'new 2010
On Error GoTo err_cboDate

    If Me![cboDate] <> "" Then
        If Me![cboFindTeam] <> "" And Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate] & ", " & Me![cboCrate]
        ElseIf Me![cboFindTeam] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate]
        ElseIf Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate] & ", " & Me![cboCrate]
        Else
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate]
        End If
        'Me.Requery
    Else
        If Me![cboFindTeam] <> "" And Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboCrate]
        ElseIf Me![cboFindTeam] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "'"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam]
        ElseIf Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboCrate]
        Else
            Me.RecordSource = RecSource
            Me!lblFilter.caption = "Current Filter: none"
        End If
    End If

Exit Sub

err_cboDate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindTeam_AfterUpdate()
'new 2010
On Error GoTo err_cboTeam

    If Me![cboFindTeam] <> "" Then
        If Me![cboDate] <> "" And Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate] & ", " & Me![cboCrate]
        ElseIf Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboCrate]
        ElseIf Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "' AND Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam] & ", " & Me![cboDate]
        Else
            Me.RecordSource = RecSource & " WHERE [MovedBy] = '" & Me![cboFindTeam] & "'"
            Me!lblFilter.caption = "Current Filter: " & Me![cboFindTeam]
        End If
        'Me.Requery
    Else
        If Me![cboDate] <> "" And Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "# AND ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate] & ", " & Me![cboCrate]
        ElseIf Me![cboCrate] <> "" Then
            Me.RecordSource = RecSource & " WHERE ([MovedFromCrate] = '" & Me![cboCrate] & "' OR [MovedToCrate] = '" & Me![cboCrate] & "')"
            Me!lblFilter.caption = "Current Filter: " & Me![cboCrate]
        ElseIf Me![cboDate] <> "" Then
            Me.RecordSource = RecSource & " WHERE Format([MovedOn], 'dd/mm/yyyy') = #" & Me![cboDate] & "#"
            Me!lblFilter.caption = "Current Filter: " & Me![cboDate]
        Else
            Me.RecordSource = RecSource
            Me!lblFilter.caption = "Current Filter: none"
        End If
    End If

Exit Sub

err_cboTeam:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command21_Click()
On Error GoTo err_cmdClose_Click

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdClose_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    Call General_Error_Trap
    Resume Exit_go_next_Click
End Sub

Private Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    Call General_Error_Trap
    Resume Exit_go_previous2_Click
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    Call General_Error_Trap
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    Call General_Error_Trap
    Resume Exit_go_last_Click
End Sub
Private Sub cmdAll_Click()
On Error GoTo Err_cmdAll_Click

    Me.RecordSource = RecSource
    Me.Requery
    Me![cboFindTeam] = ""
    Me![cboDate] = ""
    Me![cboCrate] = ""

Exit_cmdAll_Click:
    Exit Sub

Err_cmdAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdAll_Click
    
End Sub
