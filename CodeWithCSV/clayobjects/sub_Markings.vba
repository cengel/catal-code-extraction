Option Compare Database

Private Sub Form_Current()
If Me.Parent![cbo_year_studied].Value = 6 Then
    Me!location_detail.Enabled = False
    Me!add_clarity.Enabled = False
    Me.sub_Markings_location.Enabled = False
    Me.sub_Markings_clarity.Enabled = False
    Me.sub_Markings_application.Enabled = False
    Me!markings_depth.Enabled = False
    Me!cbo_fingernails.Enabled = False
Else
    Me!location_detail.Enabled = True
    Me!add_clarity.Enabled = True
    Me.sub_Markings_location.Enabled = True
    Me.sub_Markings_clarity.Enabled = True
    Me.sub_Markings_application.Enabled = True
    Me!markings_depth.Enabled = True
    Me!cbo_fingernails.Enabled = True


End If
End Sub
