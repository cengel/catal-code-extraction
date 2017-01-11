Option Compare Database


Private Sub cmdPrintUnitSheet_Click()
On Error GoTo err_print

    If LCase(Me![Category]) = "layer" Or LCase(Me![Category]) = "cluster" Then
        DoCmd.OpenReport "R_Unit_Sheet_layercluster", acViewPreview, , "[unit number] = " & Me![related_unit]
    ElseIf LCase(Me![Category]) = "cut" Then
        DoCmd.OpenReport "R_Unit_Sheet_cut", acViewPreview, , "[unit number] = " & Me![related_unit]
    ElseIf LCase(Me![Category]) = "skeleton" Then
        DoCmd.OpenReport "R_Unit_Sheet_skeleton", acViewPreview, , "[unit number] = " & Me![related_unit]
    End If
Exit Sub

err_print:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub gotorelatedunit_Click()
On Error GoTo Err_gotorelatedunit_Click
    
    'Me.Parent.SetFocus
    'DoCmd.GoToRecord , , acGoTo, "[Unit Number] = " & Me![related_unit]
    DoCmd.OpenForm "Exca: Unit Sheet", , , "[Unit Number] = " & Me![related_unit]
    
Exit_gotorelatedunit_Click:
    Exit Sub

Err_gotorelatedunit_Click:
    MsgBox Err.Description
    Resume Exit_gotorelatedunit_Click
    
End Sub
