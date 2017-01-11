Option Compare Database

Private Sub Form_Current()

If Me.Parent![cbo_year_studied].Value = 6 Then
    Me!cbo_plan2d_complete.Enabled = False
    Me!cbo_plan2d_symmetry.Enabled = False
    Me.sub_Shape_plan_2d_sidescorners.Enabled = False
    Me!plan_2d_comments.Enabled = False
    Me.sub_Shape_detail_pinched.Enabled = False
    Me!pinched_detail.Enabled = False
    Me!cbo_sect2d_complete.Enabled = False
    Me!cbo_sect2d_symmetry.Enabled = False
    Me.sub_Shape_section_2d_sidescorners.Enabled = False
    Me!section_2d_comments.Enabled = False
    Me.sub_Shape_detail_depressions.Enabled = False
    Me!depressions_detail.Enabled = False
    Me!long_comments.Enabled = False
    Me.sub_Shape_long_sidescorners.Enabled = False
    Me!cbo_long_complete.Enabled = False
    Me!cbo_long_symmetry.Enabled = False
    Me.sub_Shape_section_2d_basetop.Enabled = False
    Me.sub_Shape_long_basetop.Enabled = False
    Me.sub_Shape_long_sidescorners.Enabled = False
Else
    Me!cbo_plan2d_complete.Enabled = True
    Me!cbo_plan2d_symmetry.Enabled = True
    Me.sub_Shape_plan_2d_sidescorners.Enabled = True
    Me!plan_2d_comments.Enabled = True
    Me.sub_Shape_detail_pinched.Enabled = True
    Me!pinched_detail.Enabled = True
    Me!cbo_sect2d_complete.Enabled = True
    Me!cbo_sect2d_symmetry.Enabled = True
    Me.sub_Shape_section_2d_sidescorners.Enabled = True
    Me!section_2d_comments.Enabled = True
    Me.sub_Shape_detail_depressions.Enabled = True
    Me!depressions_detail.Enabled = True
    Me!long_comments.Enabled = True
    Me.sub_Shape_long_sidescorners.Enabled = True
    Me!cbo_long_complete.Enabled = True
    Me!cbo_long_symmetry.Enabled = True
    Me.sub_Shape_section_2d_basetop.Enabled = True
    Me.sub_Shape_long_basetop.Enabled = True
    Me.sub_Shape_long_sidescorners.Enabled = True

End If

End Sub
