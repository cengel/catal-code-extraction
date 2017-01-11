Option Explicit
Option Compare Database   'Use database order for string comparisons

Private Sub Area_AfterUpdate()
'********************************************
'Update the mound field to reflect the mound
'associated with the area, mound is now a read
'only field and users do not have to enter it
'
'SAJ v9.1
'********************************************
On Error GoTo err_Area_AfterUpdate

If Me![Area].Column(1) <> "" Then
    Me![Mound] = Me![Area].Column(1)
End If

Exit Sub
err_Area_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

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

Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
        DoCmd.GoToControl "Unit Number"
        DoCmd.FindRecord Me![cboFindUnit]
        Me![cboFindUnit] = ""
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'********************************************
'Add a new record
'
'SAJ v9.1
'********************************************
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Unit Number"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
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
'existing code to catch updates - its all over the place! Just added error trap
'think only before update is all thats required but shall keep in just in case
' SAJ v9.1
On Error GoTo err_Form_AfterInsert
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()

Exit Sub

err_Form_AfterInsert:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_AfterUpdate()
'existing code to catch updates - its all over the place! Just added error trap
'think only before update is all thats required but shall keep in just in case
' SAJ v9.1
On Error GoTo err_Form_AfterUpdate
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
Exit Sub

err_Form_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'existing code to catch updates - its all over the place! Just added error trap
'think only before update is all thats required but shall keep in just in case
' SAJ v9.1
On Error GoTo err_Form_BeforeUpdate
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
Exit Sub

err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'*************************************************************************************
' Updates since Nov 2005
'
'-Read/Only users getting update permission error as here in On_current code
' attempting to update the [Unit Data Categories <category> subform![data categories]
' field to the value of the Category combo. This only needs to be done at Category_AfterUpdate
' where it was present (looks like code has just been pasted but not amended here).
'-intro error trap
' SAJ v.9  23/11/05 (on)
'
' Also check state of record, if no unit number ie: new record make all fields read
' only so user must enter valid feature num before proceeding.
'
' This will also be useful when intro more adv security checking
' SAJ v9.1
'*************************************************************************************
Dim stDocName As String
Dim stLinkCriteria As String
    
On Error GoTo err_Form_Current
    
'if no unit number set all fields readonly
If IsNull(Me![Unit Number]) Or Me![Unit Number] = "" Then 'make rest of fields read only
    ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
    Me![lblMsg].Visible = True
Else
    ToggleFormReadOnly Me, False
    Me![lblMsg].Visible = False
End If
'current unit field always needs to be locked
Me![Text407].Locked = True


'priority button
If Me![Priority Unit] = True Then
    Me![Open Priority].Enabled = True
Else
    Me![Open Priority].Enabled = False
End If

'restore all category forms
Me![Exca: Unit Data Categories CUT subform].Visible = True
Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
Me![Exca: Unit Data Categories LAYER subform].Visible = True
    
'define which form to show
Select Case Me.Category

Case "layer"
    'descr
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
   
    
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

Case "cut"
    'descr
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = True
    'data
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    'SAJ v9 update of field restricted to category_afterupdate
    'Me![Exca: Unit Data Categories CUT subform]![Data Category] = "cut"
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
    'SAJ v9 update of field restricted to category_afterupdate
    'Me![Exca: Unit Data Categories CLUSTER subform]![Data Category] = "cluster"
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
    'SAJ v9 update of field restricted to category_afterupdate
    'Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
    Me.refresh
    
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = True
    Me![subform Unit: stratigraphy  same as].Visible = False
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: subform Skeletons same as].Visible = True

Case Else
'descr
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    'data
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    'skelli
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

End Select
Exit Sub

err_Form_Current: 'SAJ
    General_Error_Trap 'sub in generalprocedures module
    Exit Sub
End Sub


Private Sub Form_Open(Cancel As Integer)
'*************************************************************************************
' SAJ v.9.1
' form is so big maximise it so can see as much as poss - this is now required as to keep
' the main menu looking compact the system no longer maximises on startup
'*************************************************************************************
'DoCmd.Maximize
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
'replaced by cmdAddNew
'On Error GoTo Err_New_entry_Click
'
'
'    DoCmd.GoToRecord , , acNewRec
'    Mound.SetFocus
'
'Exit_New_entry_Click:
'    Exit Sub
'
'Err_New_entry_Click:
'    MsgBox Err.Description
'    Resume Exit_New_entry_Click
'
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
'***************************************************
' Existing close button revamped - image changed from
' default close (shut door) to trowel as in rest of
' season. Also made to specifically name form not just .close
'
' SAJ v9.1
'***************************************************
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



Private Sub Unit_Number_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure duplicate unit numbers not entered
' which would result in nasty key violation msg back from sql server if not
' trapped. Duplicates were previously dealt with by an undo at unit_number_exit,
' but this undo would blank the whole record with no explaination so trying
' to explain problem to user here.
'
' FOr further info concerning this functionality see the comment in
' Form - Exca: Feature Sheet, control: Feature Number, After Update
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Unit_Number_AfterUpdate
Dim checknum

If Me![Unit Number] <> "" Then
    'check that unit num not exist
    checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![Unit Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the Unit Number " & Me![Unit Number] & " already exists, please enter another number.", vbInformation, "Duplicate Unit Number"
        
        If Not IsNull(Me![Unit Number].OldValue) Then
            'return field to old value if there was one
            Me![Unit Number] = Me![Unit Number].OldValue
        Else
            'oh the joys, to keep the focus on feature have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "Year"
            DoCmd.GoToControl "Unit Number"
            Me![Unit Number].SetFocus
            
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry
        ToggleFormReadOnly Me, False
    End If
End If

Exit Sub

err_Unit_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_number_Exit(Cancel As Integer)
'*****************************************************
' This existing code is commented out and replaced by
' a handling procedure after update - the reason being
' this blanks all edits to this record done so far with
' no explaination to the user why, it also use legacy
' domenuitem.
' SAJ v9.1
'*****************************************************
'On Error GoTo Err_Unit_number_Exit
'
'    Me.Refresh
'    'DoCmd.Save acTable, "Exca: Unit Sheet"
'
'Exit_Unit_number_Exit:
'    Exit Sub
'
'Err_Unit_number_Exit:
'
'    'MsgBox Err.Description
'
'    'MsgBox "This unit already exists in the database. Use the 'Find' button to go to it.", vbOKOnly, "duplicate"
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 0, , acMenuVer70
'
'    Cancel = True
'
'    Resume Exit_Unit_number_Exit
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
