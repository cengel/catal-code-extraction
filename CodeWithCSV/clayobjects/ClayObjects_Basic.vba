Option Compare Database
Option Explicit
Private Sub Update_GID()
If Not IsNull(Me![Unit]) And Not IsNull(Me![Sample]) And Not IsNull(Me![findletter]) Then
    Me![GID] = Me![Unit] & "." & Me![findletter] & Me![Sample]
Else
    Debug.Print "GID is not valid"
End If
End Sub



Private Sub button_Clay_Click()
On Error GoTo err_button_clay_Click
Dim checknum
Dim sql
        
If Me![clay_object] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[clay_object]", "[ClayObjects_Clay]", "[clay_object] = " & Me![clay_object])
    If Not IsNull(checknum) Then
        'the number does exist so the subform will behave fine
    Else
        sql = "INSERT INTO [ClayObjects_Clay] ([clay_object]) VALUES (" & Me![clay_object] & ");"
        DoCmd.RunSQL sql
        
        'ToggleFormReadOnly Me, False
    End If
End If
Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Clay.Visible = True
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Shape.Visible = False
Me.Form!sub_Markings.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = True
Me!button_markings.Enabled = True
Me!button_Clay.Enabled = False
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = True
Me.Form!sub_Clay.Requery
Exit Sub

err_button_clay_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub button_dim_Click()
Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Measure.Visible = True
Me.Form!sub_Shape.Visible = False
Me.Form!sub_Markings.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = False
Me!button_shape.Enabled = True
Me!button_markings.Enabled = True
Me!button_Clay.Enabled = True
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = True
Me.Form!sub_Measure.Requery
End Sub

Private Sub button_manufacture_Click()
On Error GoTo err_button_manufacture_Click
Dim checknum
Dim sql
        
If Me![clay_object] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[clay_object]", "[ClayObjects_Manufacture]", "[clay_object] = " & Me![clay_object])
    If Not IsNull(checknum) Then
        'the number does exist so the subform will behave fine
    Else
        sql = "INSERT INTO [ClayObjects_Manufacture] ([clay_object]) VALUES (" & Me![clay_object] & ");"
        DoCmd.RunSQL sql
        
        'ToggleFormReadOnly Me, False
    End If
End If
Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Manufacture.Visible = True
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Shape.Visible = False
Me.Form!sub_Markings.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = True
Me!button_markings.Enabled = True
Me!button_manufacture.Enabled = False
Me!button_Clay.Enabled = True
Me!button_post.Enabled = True
Me.Form!sub_Manufacture.Requery
Exit Sub

err_button_manufacture_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub button_markings_Click()
On Error GoTo err_button_markings_Click
Dim checknum
Dim sql
        
If Me![clay_object] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[clay_object]", "[ClayObjects_Markings]", "[clay_object] = " & Me![clay_object])
    If Not IsNull(checknum) Then
        'the number does exist so the subform will behave fine
    Else
        sql = "INSERT INTO [ClayObjects_Markings] ([clay_object]) VALUES (" & Me![clay_object] & ");"
        DoCmd.RunSQL sql
        
        'ToggleFormReadOnly Me, False
    End If
End If
Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Markings.Visible = True
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Shape.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = True
Me!button_markings.Enabled = False
Me!button_Clay.Enabled = True
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = True
Me.Form!sub_Markings.Requery
Exit Sub

err_button_markings_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub button_post_Click()
Me.Form!sub_Postproduction.Visible = True
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Markings.Visible = False
Me.Form!sub_Shape.Visible = False
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = True
Me!button_markings.Enabled = True
Me!button_Clay.Enabled = True
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = False
Me.Form!sub_Postproduction.Requery
End Sub

Private Sub button_shape_Click()
On Error GoTo err_button_shape_Click
Dim checknum
Dim sql
        
If Me![clay_object] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[clay_object]", "[ClayObjects_Shape]", "[clay_object] = " & Me![clay_object])
    If Not IsNull(checknum) Then
        'the number does exist so the subform will behave fine
    Else
        sql = "INSERT INTO [ClayObjects_Shape] ([clay_object]) VALUES (" & Me![clay_object] & ");"
        DoCmd.RunSQL sql
        
        'ToggleFormReadOnly Me, False
    End If
End If

Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Markings.Visible = False
Me.Form!sub_Shape.Visible = True
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = False
Me!button_markings.Enabled = True
Me!button_Clay.Enabled = True
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = True
Me.Form!sub_Shape.Requery
Exit Sub

err_button_shape_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboFilterUnit_AfterUpdate()
'filter - new 2011
On Error GoTo err_filter

    If Me![cboFilterUnit] <> "" Then
        Me.Filter = "[Unit] = " & Me![cboFilterUnit]
        Me.FilterOn = True
        Me![cboFilterUnit] = ""
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFilterUnit_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - new 2011
On Error GoTo err_cbofilterNot

    MsgBox "Sorry this Unit does not exist in this database yet", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFilterUnit].Undo
Exit Sub

err_cbofilterNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
'new 2011
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    DoCmd.GoToControl "GID"
    DoCmd.FindRecord Me![cboFind]
    ''DoCmd.GoToControl "Analyst"
    Me![cboFind] = ""
End If


Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub




Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - new 2011
On Error GoTo err_cbofindNot

    MsgBox "Sorry this GID does not exist in the database", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub clay_object_AfterUpdate()
On Error GoTo err_clay_object_AfterUpdate
Dim checknum
Dim sql

If Me![clay_object] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[clay_object]", "[ClayObjects_Basic]", "[clay_object] = " & Me![clay_object])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the clay object " & Me![clay_object] & " already exists, please enter another number.", vbInformation, "Duplicate Unit Number"
        
        If Not IsNull(Me![clay_object].OldValue) Then
            'return field to old value if there was one
            Me![clay_object] = Me![clay_object].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "Unit"
            DoCmd.GoToControl "clay_object"
            Me![clay_object].SetFocus
            
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry
        'ToggleFormReadOnly Me, False
        checknum = DLookup("[clay_object]", "[ClayObjects_Shape]", "[clay_object] = " & Me![clay_object])
         If Not IsNull(checknum) Then
                'the number does exist so the subform will behave fine
        Else
            sql = "INSERT INTO [ClayObjects_Shape] ([clay_object]) VALUES (" & Me![clay_object] & ");"
            DoCmd.RunSQL sql
        
        'ToggleFormReadOnly Me, False
    End If
    End If
End If

Exit Sub

err_clay_object_AfterUpdate:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cmdAddNew_Click()
'********************************************************************
' Create new record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgonew_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acNewRec
    ''DoCmd.GoToControl Me![frm_subform_basic].Form![Analyst].Name
    DoCmd.GoToControl Me![Unit].Name
    Exit Sub

Err_cmdgonew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'new 2011 - control the delete of a record to ensure both tables are clear
On Error GoTo err_del

Dim response
    response = MsgBox("Do you really want to remove GID " & Me!GID & " (DB Id " & Me![clay_object] & ") and all its related material records from your database?", vbYesNo + vbQuestion, "Remove Record")
    If response = vbYes Then
        Dim sql
        
    sql = "Delete FROM [ClayObjects_Basic] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_plan_2d] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_section_2d] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_long] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_plan_2d_sidescorners] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_section_2d_sidescorners] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_long_sidescorners] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_long_basetop] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_section_2d_basetop] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_detail_pinched] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape_detail_depressions] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Shape] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Measure] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Markings] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Manufacture] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Postproduction] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Markings_location] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Markings_clarity] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Markings_application] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Markings_type] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay_primary_colour] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay_texture] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay_surface] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay_inclusions] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Clay_inclusions_specified] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Manufacture_craft] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Manufacture_applied] WHERE [clay_object] = " & Me![clay_object] & ";"
        DoCmd.RunSQL sql
    sql = "Delete FROM [ClayObjects_Post_burning] WHERE [clay_object] = " & Me![clay_object] & ";"
        
        Me.Requery
        MsgBox "Deletion completed", vbInformation, "Done"
        
        Me![cboFind].Requery
        Me![cboFilterUnit].Requery
        
    End If
Exit Sub

err_del:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_cmdgofirst_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoLast_Click()
'********************************************************************
' Go to last record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoLast_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_cmdgoLast_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoNext_Click()
'********************************************************************
' Go to next record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoNext_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_cmdgoNext_Click:
    If Err.Number = 2105 Then
        MsgBox "No more records to show", vbInformation, "End of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdGoPrev_Click()
'********************************************************************
' Go to previous record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoprevious_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_cmdgoprevious_Click:
    If Err.Number = 2105 Then
        MsgBox "Already at the begining of the recordset", vbInformation, "Begining of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub



Private Sub cmdRemoveFilter_Click()
'remove unit filter - new 2011
On Error GoTo err_Removefilter

    Me![cboFilterUnit] = ""
    Me.Filter = ""
    Me.FilterOn = False
    
    DoCmd.GoToControl "cboFind"
    Me![cmdRemoveFilter].Visible = False
   

Exit Sub

err_Removefilter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Flot_Sample_number_AfterUpdate()
'new season 2006 - get soil vol from flot log
On Error GoTo err_sample

Update_GID


Exit Sub

err_sample:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Form_Load()
Me.Form!sub_Postproduction.Visible = False
Me.Form!sub_Clay.Visible = False
Me.Form!sub_Markings.Visible = False
Me.Form!sub_Shape.Visible = True
Me.Form!sub_Measure.Visible = False
Me.Form!sub_Manufacture.Visible = False
Me!button_dim.Enabled = True
Me!button_shape.Enabled = False
Me!button_markings.Enabled = True
Me!button_Clay.Enabled = True
Me!button_manufacture.Enabled = True
Me!button_post.Enabled = True
End Sub

Private Sub Sample_AfterUpdate()
On Error GoTo err_Sample_AfterUpdate
Dim checknum

If Me![Sample] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[findnumber]", "[ClayObjects_Basic]", "[findnumber] = " & Me![Sample] & " AND [unit] = " & Me![Unit] & " AND [findsampleletter] = '" & Me![findletter] & "'")
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the find " & Me![Unit] & "." & Me![findletter] & "" & Me![Sample] & " already exists, please enter another number.", vbInformation, "Duplicate gid Number"
        
        If Not IsNull(Me![Sample].OldValue) Then
            'return field to old value if there was one
            Me![Sample] = Me![Sample].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry
        'ToggleFormReadOnly Me, False
    End If
End If

Update_GID

Exit Sub

err_Sample_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Sample_Change()
'comment out saj
'Update_GID
End Sub


Private Sub Sample_Enter()
'SAJ before versioning - this causes sql update error to be returned to user even
'they have not tried to edit anything, most confusing and unnecessary so removed
' 11/01/05
'Update_GID
End Sub


Private Sub Unit_AfterUpdate()
Update_GID
End Sub

Private Sub Unit_Change()
'Update_GID
End Sub



