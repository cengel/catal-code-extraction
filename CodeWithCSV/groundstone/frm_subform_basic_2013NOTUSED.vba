Option Compare Database

Private Sub cboLettercode_AfterUpdate()
'update the GID
On Error GoTo err_lc

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_lc:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub FindNumber_AfterUpdate()
'update the GID
On Error GoTo err_fn

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_fn:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()

'update the GID
On Error GoTo err_unit

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_DblClick(Cancel As Integer)

On Error GoTo Err_cmdUnitDesc_Click

If Me![Unit] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[dbo_Groundstone: Unit Description_2013]", "[Unit] = " & Me![Unit])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [dbo_Groundstone: Unit Description_2013] ([Unit]) VALUES (" & Me![Unit] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "Frm_GS_UnitDescription_2013", acNormal, , "[Unit] = " & Me![Unit], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub

End Sub
