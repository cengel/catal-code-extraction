Option Compare Database

Private Sub cboFilterArea_AfterUpdate()
'filter - new 2011
On Error GoTo err_filter

    If Me![cboFilterArea] <> "" Then
    
        Me.[dbo_view_exca_unit_lateststatus_2bchecked].Form.Filter = "[Area] = '" & Me![cboFilterArea] & "'"
        Me.[dbo_view_exca_unit_lateststatus_checked].Form.Filter = "[Area] = '" & Me![cboFilterArea] & "'"
        Me.[dbo_view_exca_unit_lateststatus_inprogress].Form.Filter = "[Area] = '" & Me![cboFilterArea] & "'"
        Me.[dbo_view_exca_unit_lateststatus_2bchecked].Form.FilterOn = True
        Me.[dbo_view_exca_unit_lateststatus_checked].Form.FilterOn = True
        Me.[dbo_view_exca_unit_lateststatus_inprogress].Form.FilterOn = True
        Me![cboFilterArea] = ""
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFilterArea_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - new 2011
On Error GoTo err_cbofilterNot

    MsgBox "Sorry this Area does not exist in this database yet", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFilterArea].Undo
Exit Sub

err_cbofilterNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRemoveFilter_Click()
'remove unit filter - new 2011
On Error GoTo err_Removefilter

    Me![cboFilterArea] = ""
    Me.Filter = ""
    Me.[dbo_view_exca_unit_lateststatus_2bchecked].Form.FilterOn = False
    Me.[dbo_view_exca_unit_lateststatus_checked].Form.FilterOn = False
    Me.[dbo_view_exca_unit_lateststatus_inprogress].Form.FilterOn = False
    
    'DoCmd.GoToControl "cboFind"
    Me![cboFilterArea].SetFocus
    Me![cmdRemoveFilter].Visible = False
   

Exit Sub

err_Removefilter:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub refresh_Click()

Me![dbo_view_exca_unit_lateststatus_2bchecked].Requery
Me![dbo_view_exca_unit_lateststatus_checked].Requery
Me![dbo_view_exca_unit_lateststatus_inprogress].Requery

End Sub
