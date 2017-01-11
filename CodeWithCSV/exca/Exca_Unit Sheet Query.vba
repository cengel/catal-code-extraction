Option Explicit
Option Compare Database   'Use database order for string comparisons

Private Sub Category_AfterUpdate()

Select Case Me.Category

Case "cut"
    'descr
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = True
    'data
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    Me![Exca: Unit Data Categories CUT subform]![Data Category] = "cut"
        'the rest need to be blank
    Me![Exca: Unit Data Categories CUT subform]![In Situ] = ""
    Me![Exca: Unit Data Categories CUT subform]![Location] = ""
    Me![Exca: Unit Data Categories CUT subform]![Description] = ""
    Me![Exca: Unit Data Categories CUT subform]![Material] = ""
    Me![Exca: Unit Data Categories CUT subform]![Deposition] = ""
    Me![Exca: Unit Data Categories CUT subform]![basal spit] = ""
    Me.refresh
    
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

Case "layer"
    'descr
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: Unit Data Categories LAYER subform]![Data Category] = ""
        'the rest need to be blank
    Me![Exca: Unit Data Categories LAYER subform]![In Situ] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Location] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Description] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Material] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Deposition] = ""
    Me![Exca: Unit Data Categories LAYER subform]![basal spit] = ""
    Me.refresh
    
    
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    
Case "cluster"
    'descr
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform]![Data Category] = "cluster"
        'the rest need to be blank
    Me![Exca: Unit Data Categories CLUSTER subform]![In Situ] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Location] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Description] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Material] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Deposition] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![basal spit] = ""
    Me.refresh
        
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

Case "skeleton"
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
    'the rest need to be blank
    Me![Exca: Unit Data Categories SKELL subform]![In Situ] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Location] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Description] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Material] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Deposition] = ""
    Me![Exca: Unit Data Categories SKELL subform]![basal spit] = ""
        
    Me.refresh
    
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = True
    Me![subform Unit: stratigraphy  same as].Visible = False
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: subform Skeletons same as].Visible = True

End Select

End Sub

Private Sub copy_method_Click()
On Error GoTo Err_copy_method_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Copy unit methodology"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_copy_method_Click:
    Exit Sub

Err_copy_method_Click:
    MsgBox Err.Description
    Resume Exit_copy_method_Click
    

End Sub

Private Sub cut_desc_Click()
On Error GoTo Err_cut_desc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Cut Display"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cut_desc_Click:
    Exit Sub

Err_cut_desc_Click:
    MsgBox Err.Description
    Resume Exit_cut_desc_Click
    
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
Dim stDocName As String
Dim stLinkCriteria As String
    
'restore all category forms
Me![Exca: Unit Data Categories CUT subform].Visible = True
Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
Me![Exca: Unit Data Categories LAYER subform].Visible = True
    
'define which form to show
Select Case Me.Category

Case "layer"
    'descr
    Me![layer desc].Visible = True
    Me![cut desc].Visible = False
    Me![skell desc].Visible = False
    
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
   
    
    'skelli
    'Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

Case "cut"
    'descr
    'Me![Exca: Subform Layer descr].Visible = False
    'Me![Exca: Subform Cut descr].Visible = True
    Me![layer desc].Visible = False
    Me![cut desc].Visible = True
    Me![skell desc].Visible = False
    
    'data
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    Me![Exca: Unit Data Categories CUT subform]![Data Category] = "cut"
    Me.refresh
    
    'skelli
    'Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

    
Case "cluster"
    'descr
    'Me![Exca: Subform Layer descr].Visible = True
    'Me![Exca: Subform Cut descr].Visible = False
    Me![layer desc].Visible = True
    Me![cut desc].Visible = False
    Me![skell desc].Visible = False
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform]![Data Category] = "cluster"
    Me.refresh
    
    'skelli
    'Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

Case "skeleton"
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
    Me.refresh
    
    'desc
    Me![layer desc].Visible = False
    Me![cut desc].Visible = False
    Me![skell desc].Visible = True
    
    'skelli
    'Me![Exca: subform Skeleton Sheet].Visible = True
    Me![subform Unit: stratigraphy  same as].Visible = False
    'Me![Exca: Subform Layer descr].Visible = False
    'Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: subform Skeletons same as].Visible = True

Case Else
'descr
    'Me![Exca: Subform Layer descr].Visible = True
    'Me![Exca: Subform Cut descr].Visible = False
    Me![layer desc].Visible = True
    Me![cut desc].Visible = False
    Me![skell desc].Visible = False
    
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    'skelli
    'Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

End Select

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

Private Sub skell_desc_Click()
On Error GoTo Err_skell_desc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Skeleton Display"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_skell_desc_Click:
    Exit Sub

Err_skell_desc_Click:
    MsgBox Err.Description
    Resume Exit_skell_desc_Click
    
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


Sub Command497_Click()
On Error GoTo Err_Command497_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Skeleton Sheet"
    
    stLinkCriteria = "[Exca: Unit Sheet.Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command497_Click:
    Exit Sub

Err_Command497_Click:
    MsgBox Err.Description
    Resume Exit_Command497_Click
    
End Sub
Sub go_skell_Click()
On Error GoTo Err_go_skell_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Skeleton Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_go_skell_Click:
    Exit Sub

Err_go_skell_Click:
    MsgBox Err.Description
    Resume Exit_go_skell_Click
    
End Sub
Private Sub layer_desc_Click()
On Error GoTo Err_layer_desc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Layer Display"
    
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_layer_desc_Click:
    Exit Sub

Err_layer_desc_Click:
    MsgBox Err.Description
    Resume Exit_layer_desc_Click
    
End Sub

