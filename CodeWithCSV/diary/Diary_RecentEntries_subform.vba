Option Compare Database


Private Sub cmdOpenDiaryEntry_Click()
'new season 2012 - open single diary entry

On Error GoTo err_opendiary
    'DoCmd.OpenForm "Diary_SingleEntry_popup", acNormal, , ,acFormReadOnly, , Me![txt_diaryID]
    DoCmd.OpenForm "Diary_SingleEntry_popup", acNormal, , "ID=" & Me![txt_diaryID], acFormReadOnly, , Me![txt_diaryID]

Exit Sub

err_opendiary:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdCloseDiaryRecentEntries_Click()
On Error GoTo Err_Close_Form_Click

    DoCmd.Close

Exit_Close_Form_Click:
    Exit Sub

Err_Close_Form_Click:
    MsgBox Err.Description
    Resume Exit_Close_Form_Click
    
End Sub

Private Sub Post_a_new_entry_Click()
On Error GoTo Err_Post_a_new_entry_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Diary Form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Post_a_new_entry_Click:
    Exit Sub

Err_Post_a_new_entry_Click:
    MsgBox Err.Description
    Resume Exit_Post_a_new_entry_Click
    
End Sub
