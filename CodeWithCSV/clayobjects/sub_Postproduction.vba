Option Compare Database

Private Sub Form_Current()
If Me.Parent![cbo_year_studied].Value = 6 Then
    Me!cbo_breakage.Enabled = False
    Me!breakage_detail.Enabled = False
    Me!cbo_heavywear.Enabled = False
    Me!heavy_wear_detail.Enabled = False
    Me!adhering_material.Enabled = False
Else
    Me!cbo_breakage.Enabled = True
    Me!breakage_detail.Enabled = True
    Me!cbo_heavywear.Enabled = True
    Me!heavy_wear_detail.Enabled = True
    Me!adhering_material.Enabled = True
End If
End Sub
