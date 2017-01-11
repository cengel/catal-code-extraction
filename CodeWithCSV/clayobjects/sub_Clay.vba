Option Compare Database

Private Sub Form_Current()
If Me.Parent![cbo_year_studied].Value = 6 Then
    Me!cbo_heat.Enabled = False
    Me!yn_coated.Enabled = False
    Me.sub_Clay_texture.Enabled = False

Else
    Me!cbo_heat.Enabled = True
    Me!yn_coated.Enabled = True
    Me.sub_Clay_texture.Enabled = True
End If
End Sub
