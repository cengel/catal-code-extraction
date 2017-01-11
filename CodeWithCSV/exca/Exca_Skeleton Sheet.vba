Option Compare Database   'Use database order for string comparisons

Private Sub Category_AfterUpdate()

If Me![Category] = "Cut" Then
    Me![LayerLabel].Visible = False
    Me![CutLAbel].Visible = True
Else
    Me![LayerLabel].Visible = True
    Me![CutLAbel].Visible = False
End If

End Sub

Private Sub Category_Change()

If Me![Category] = "Cut" Then
    Me![LayerLabel].Visible = False
    Me![CutLAbel].Visible = True
Else
    Me![LayerLabel].Visible = True
    Me![CutLAbel].Visible = False
End If

End Sub


Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Unit Sheet"
    
Exit_Excavation_Click:
    Exit Sub

err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub

Sub find_unit_Click()
On Error GoTo Err_find_unit_Click


    Screen.PreviousControl.SetFocus
    Unit_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_unit_Click:
    Exit Sub

Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
    
End Sub


Private Sub Form_AfterInsert()
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub

Private Sub Form_AfterUpdate()
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub

Private Sub Form_Current()

If Me![Category] <> "Skeleton" Then
    'return to unit sheet
    
    DoCmd.Close
Else
    'do nothing
End If

End Sub

Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
    
End Sub


Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
    
End Sub

Sub go_to_last_Click()

On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
    
End Sub





Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    MsgBox Err.Description
    Resume Exit_go_previous2_Click
    
End Sub

Private Sub Master_Control_Click()
On Error GoTo Err_Master_Control_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Unit Sheet"
    
Exit_Master_Control_Click:
    Exit Sub

Err_Master_Control_Click:
    MsgBox Err.Description
    Resume Exit_Master_Control_Click
End Sub

Sub New_entry_Click()
On Error GoTo Err_New_entry_Click


    DoCmd.GoToRecord , , acNewRec
    Mound.SetFocus
    
Exit_New_entry_Click:
    Exit Sub

Err_New_entry_Click:
    MsgBox Err.Description
    Resume Exit_New_entry_Click
    
End Sub
Sub interpretation_Click()
On Error GoTo Err_interpretation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    'refresh
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    
    'go to form
    stDocName = "Interpret: Unit Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_interpretation_Click:
    Exit Sub

Err_interpretation_Click:
    MsgBox Err.Description
    Resume Exit_interpretation_Click
    
End Sub
Sub Command466_Click()
On Error GoTo Err_Command466_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_Command466_Click:
    Exit Sub

Err_Command466_Click:
    MsgBox Err.Description
    Resume Exit_Command466_Click
    
End Sub
Sub Open_priority_Click()
On Error GoTo Err_Open_priority_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Priority Detail"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Open_priority_Click:
    Exit Sub

Err_Open_priority_Click:
    MsgBox Err.Description
    Resume Exit_Open_priority_Click
    
End Sub
Sub go_feature_Click()
On Error GoTo Err_go_feature_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    
    stLinkCriteria = "[Feature Number]=" & Forms![Exca: Unit Sheet]![Exca: subform Features for Units]![In_feature]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_go_feature_Click:
    Exit Sub

Err_go_feature_Click:
    MsgBox Err.Description
    Resume Exit_go_feature_Click
    
End Sub
Sub Close_Click()
On Error GoTo err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
Sub open_copy_details_Click()
On Error GoTo Err_open_copy_details_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Copy unit details form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_copy_details_Click:
    Exit Sub

Err_open_copy_details_Click:
    MsgBox Err.Description
    Resume Exit_open_copy_details_Click
    
End Sub

Private Sub Unit_number_Exit(Cancel As Integer)
On Error GoTo Err_Unit_number_Exit

    Me.refresh
    'DoCmd.Save acTable, "Exca: Unit Sheet"
    
Exit_Unit_number_Exit:
    Exit Sub

Err_Unit_number_Exit:
   
    'MsgBox Err.Description
    
    'MsgBox "This unit already exists in the database. Use the 'Find' button to go to it.", vbOKOnly, "duplicate"
    DoCmd.DoMenuItem acFormBar, acEditMenu, 0, , acMenuVer70
    
    Cancel = True
        
    Resume Exit_Unit_number_Exit
End Sub


