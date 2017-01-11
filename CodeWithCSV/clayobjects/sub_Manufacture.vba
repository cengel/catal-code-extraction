Option Compare Database

Private Sub Form_Current()
If Me.Parent![cbo_year_studied].Value = 6 Then
    Me!fingerprints_comment.Enabled = False
    Me!yn_plant.Enabled = False
    Me.sub_Manufacture_craft.Enabled = False
    Me.sub_Manufacture_applied.Enabled = False
Else
    Me!fingerprints_comment.Enabled = True
    Me!yn_plant.Enabled = True
    Me.sub_Manufacture_craft.Enabled = True
    Me.sub_Manufacture_applied.Enabled = True
End If
End Sub
