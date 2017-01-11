Option Compare Database


Private Sub addNew_Click()
On Error GoTo err_addNew_Click

DoCmd.GoToRecord , , acNewRec
Forms![Phytolith_NewSampleRegister].Form![UnitNumber].SetFocus

Exit Sub

err_addNew_Click:
    MsgBox "An error has occured setting up a new record, the error is as follows: " & Err.Description
    Exit Sub
End Sub

Private Sub closeform_Click()
On Error GoTo err_closeform_Click

Dim lastunit, lastsample
Dim checknum

If Not IsNull(Me![UnitNumber]) And Not IsNull(Me![SampleNumber]) And _
Me![UnitNumber] <> "" And Me![SampleNumber] <> "" Then
    lastunit = Me![UnitNumber]
    lastsample = Me![SampleNumber]
    Forms![Phytolith_NewSampleRegister].Form.Requery
    DoCmd.OpenForm "frm_Phyto_SampleRegister", acNormal
    Forms![frm_Phyto_SampleRegister].Form.Requery
    checknum = DLookup("[UniqueID]", "[Phytolith_SampleRegister]", "[UnitNumber] = " & lastunit & " And [SampleNumber] = " & lastsample)
    Debug.Print checknum
    If Not IsNull(checknum) Then
        Forms![frm_Phyto_SampleRegister].Form![UniqueID].SetFocus
        DoCmd.FindRecord checknum
    Else
        DoCmd.GoToRecord , "frm_Phyto_SampleRegister", acLast
    End If
Else
    Forms![Phytolith_NewSampleRegister].Form.Requery
    DoCmd.OpenForm "frm_Phyto_SampleRegister", acNormal
    Forms![frm_Phyto_SampleRegister].Form.Requery
    DoCmd.GoToRecord , "frm_Phyto_SampleRegister", acLast
End If


Forms![frm_Phyto_SampleRegister].Form![UnitNumber].SetFocus
DoCmd.Close , "Phytolith_NewSampleRegister"

Exit Sub

err_closeform_Click:
    MsgBox "An error has occured setting up a new record, the error is as follows: " & Err.Description
    Exit Sub

End Sub
