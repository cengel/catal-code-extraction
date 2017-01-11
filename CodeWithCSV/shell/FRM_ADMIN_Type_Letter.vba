Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
'close this form and return to main menu
On Error GoTo err_close
    
    DoCmd.OpenForm "FRM_menu"
    DoCmd.Restore
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:

    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Delete(Cancel As Integer)
'must check this entry has not been used before deletion allowed
On Error GoTo err_delete

Dim checknum

    
    checknum = DLookup("[Type]", "[Shell_Level_One_Data]", "[Type] = " & Me![type number])
    If IsNull(checknum) Then
        'number not exist can allow delete
        Cancel = False
    Else
        'number exists do not allow
        MsgBox "This type number has been used in data entry, please edit the relevant records first and then return to delete it.", vbCritical, "Invalid Action"
        Cancel = True
    End If
    
Exit Sub
err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub genus_AfterUpdate()
'check genus not already used
On Error GoTo err_num

    Dim oldval, checknum

    oldval = Me![genus].OldValue
    checknum = DLookup("[species]", "[Shell_Level_One_Data]", "[species] = '" & oldval & "'")
    If Not IsNull(checknum) Then
        'number exists do not allow
        MsgBox "This genus has been used in data entry, please edit the relevant records first and then return to change it.", vbCritical, "Invalid Action"
        Me![genus] = oldval
    End If


Exit Sub

err_num:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub type_number_AfterUpdate()
'check number not already used
On Error GoTo err_num

    Dim oldval, checknum

    oldval = Me![type number].OldValue
    checknum = DLookup("[Type]", "[Shell_Level_One_Data]", "[Type] = " & oldval)
    If Not IsNull(checknum) Then
        'number exists do not allow
        MsgBox "This type number has been used in data entry, please edit the relevant records first and then return to change it.", vbCritical, "Invalid Action"
        Me![type number] = oldval
    End If


Exit Sub

err_num:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'maximise
On Error GoTo err_open

    DoCmd.Maximize

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
