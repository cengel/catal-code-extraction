Option Compare Database

Private Sub cmdCloseTagFrequencies_Click()
On Error GoTo Err_Close_Form_Click

    DoCmd.Close

Exit_Close_Form_Click:
    Exit Sub

Err_Close_Form_Click:
    MsgBox Err.Description
    Resume Exit_Close_Form_Click
    
End Sub

Private Sub txtTags_DblClick(Cancel As Integer)

On Error GoTo err_txtTags
'MsgBox Me![txtTags]
    DoCmd.OpenForm "Diary_EntriesForOneTag_subform", acNormal, , "Tags = '" & Me![txtTags] & "' ", acFormReadOnly
Exit Sub

err_txtTags:
    Call General_Error_Trap
    Exit Sub
    
End Sub
