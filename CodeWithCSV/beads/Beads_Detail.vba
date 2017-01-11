Option Compare Database


Private Sub add_description_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub amount_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub cbo_geocategory_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub cbo_material_Change()
On Error GoTo err_filter

If Me!cbo_material.Value = 2 Then
    Me!faunal_element.Visible = True
    Me!faunal_element.Enabled = True
Else
    Me!faunal_element.Visible = False
    Me!faunal_element.Enabled = False
End If

    Forms![Beads: Sheet]![timestamp] = Now()

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cbo_rawmaterial_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub colour_detail_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub count_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub diameter_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub faunal_element_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Form_Current()
On Error GoTo err_filter

If Me!cbo_material.Value = 2 Then
    Me!faunal_element.Visible = True
    Me!faunal_element.Enabled = True
Else
    Me!faunal_element.Visible = False
    Me!faunal_element.Enabled = False
End If

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub height_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld322_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld324_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld326_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld328_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld330_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld332_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld334_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub Kombinationsfeld336_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub length_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub manufacture_body_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub manufacture_margins_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub manufacture_useface_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub micromacro_body_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub micromacro_margins_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub micromacro_useface_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub munsell_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub ornament_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub repair_recycle_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub subtype_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub weardegree_body_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub weardegree_margins_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub weardegree_useface_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub

Private Sub width_Change()
    Forms![Beads: Sheet]![timestamp] = Now()
End Sub
