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

Private Sub Building_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure building num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter building num when they can
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Building_AfterUpdate

Dim checknum, msg, retval, sql

If Me![Building] <> "" Then
    'first check its valid
    If IsNumeric(Me![Building]) Then
    
        'check that space num does exist
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
        If IsNull(checknum) Then
            msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                sql = "INSERT INTO [Exca: Building Details] ([Number]) VALUES (" & Me![Building] & ");"
                DoCmd.RunSQL sql
                
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![Building], acFormEdit, acDialog
            End If
        Else
            'valid number, enable view button
            Me![cmdGoToBuilding].Enabled = True
        End If
    
    Else
        'not a vaild building number
        MsgBox "This Building number is not numeric, it cannot be checked for validity", vbInformation, "Non numeric Entry"
    End If
End If

Exit Sub

err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Category_AfterUpdate()
'***********************************************************************************
' This is existing code. It determines which sub forms to show on the basis of the
' general category picked. The categories are hardcoded as a value list and then the
' values used in the code here. SF says these will not change and is ok for this to
' remain the case.
'
' Categories in the values list are: "Layer";"Cut";"Cluster";"Skeleton"
' SAJ v9.1 - did intro error trap that not there before, added data categories SKELL
' subform vis/invis inline with other subforms of this nature
'***********************************************************************************
On Error GoTo Err_Category_AfterUpdate

Select Case Me.Category
'action is based on value selected
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
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
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
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
    
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

    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False

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

    'SAJ v9.1 make this visible to make consistent with other forms of this nature
    Me![Exca: Unit Data Categories SKELL subform].Visible = True
End Select
Exit Sub

Err_Category_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'for existing number the field with be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![Unit Number].Enabled = False Then Me![Unit Number].Enabled = True
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

Private Sub cmdGoToBuilding_Click()
'***********************************************************
' Open space form with a filter on the number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToBuilding_Click
Dim checknum, msg, retval, sql, permiss

If Not IsNull(Me![Building]) Or Me![Building] <> "" Then
    'check that building num does exist
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
    If IsNull(checknum) Then
        'number not exist - now see what permissions user has
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            msg = "This Building Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
            If retval = vbNo Then
                MsgBox "No Building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                sql = "INSERT INTO [Exca: Building Details] ([Number]) VALUES (" & Me![Building] & ");"
                DoCmd.RunSQL sql
                
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![Building], acFormEdit, acDialog
            End If
        Else
            'user is readonly so just tell them record not exist
            MsgBox "Sorry but this building record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Building Record"
        End If
    Else
        'record exists - open it
        Dim stDocName As String
        Dim stLinkCriteria As String

        stDocName = "Exca: Building Sheet"
    
        stLinkCriteria = "[Number]= " & Me![Building]
        'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog, "FILTER"
        'decided against dialog as can open other forms on the building form and they would appear underneath it
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, , "FILTER"
    End If
    
End If

Exit Sub

Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoToSpace_Click()
'***********************************************************
' Open space form with a filter on the number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToSpace_Click
Dim checknum, msg, retval, sql, permiss

If Not IsNull(Me![Space]) Or Me![Space] <> "" Then
    'check that space num does exist
    checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
    If IsNull(checknum) Then
        'number not exist - now see what permissions user has
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            msg = "This Space Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
        
            If retval = vbNo Then
                MsgBox "No Space record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number]) VALUES ('" & Me![Space] & "');"
                DoCmd.RunSQL sql
                
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = '" & Me![Space] & "'", acFormEdit, acDialog
            End If
        Else
            'user is readonly so just tell them record not exist
            MsgBox "Sorry but this space record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Space Record"
        End If
    Else
        'record exists - open it
        Dim stDocName As String
        Dim stLinkCriteria As String

        stDocName = "Exca: Space Sheet"
    
        stLinkCriteria = "[Space Number]= '" & Me![Space] & "'"
        'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog, "FILTER"
        'decided against dialog as can open other forms on the space form and they would appear underneath it
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, , "FILTER"
    End If
    
End If

Exit Sub

Err_cmdGoToSpace_Click:
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

Private Sub FastTrack_Click()
'*********************************************************************
' Introduce logic to fast track option - if this is selected then
' Not excavated must be false
' SAJ v9.1
'*********************************************************************
On Error GoTo err_FastTrack_Click

    If Me![FastTrack] = True Then
        Me![NotExcavated] = False
    End If
Exit Sub

err_FastTrack_Click:
    Call General_Error_Trap
    Exit Sub
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
'*******************************************************************
'existing code to catch updates - its all over the place! Just added error trap
'think only before update is all thats required but shall keep in just in case
'
'Also new requirement - if user edits record but no plan number exists
'then prompt them
' SAJv9.1
'*******************************************************************
On Error GoTo err_Form_BeforeUpdate

If IsNull(Me![Exca: subform Unit Plans].Form![Graphic Number]) Then
    'this event will trigger when move to subform, which is really hard
    'on this form as there are so many. So really can only catch this
    'when they are editing at the bottom of the form
    If Me.ActiveControl.Name = "Discussion" Or Me.ActiveControl.Name = "Checked By" Or Me.ActiveControl.Name = "Date Checked" Or Me.ActiveControl.Name = "Phase" Then
        MsgBox "There is no Plan number entered for this Unit. Please can you enter one soon", vbInformation, "What is the Plan Number?"
    End If
End If

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
'
' Make DataCategories SKELL subform vis/invis in line with other subforms of this nature
' other changes marked ** SAJ v9.1
' New requirement that unit number cannot be edited after entry. This
' can only be done by an administrator so lock field here
' SAJ v9.1
'*************************************************************************************
Dim stDocName As String
Dim stLinkCriteria As String
    
On Error GoTo err_Form_Current
    
Dim permiss
permiss = GetGeneralPermissions
If permiss = "ADMIN" Or permiss = "RW" Then
    'if no unit number set all fields readonly
    If IsNull(Me![Unit Number]) Or Me![Unit Number] = "" Then 'make rest of fields read only
        ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
        Me![lblMsg].Visible = True
        
        'no Unit number - new record allow entry
        Me![Unit Number].Locked = False
        Me![Unit Number].Enabled = True
        Me![Unit Number].BackColor = 16777215
        Me![Unit Number].SetFocus
        
    Else
        ''ToggleFormReadOnly Me, False
        ''Me![lblMsg].Visible = False
        
        'if coming in as a filter thats readonly then send in extra arg
        If Me.FilterOn = True And Me.AllowEdits = False Then
            'when popped up from the feature form this was allowing new records to be added, altered to fix
            'ToggleFormReadOnly Me, False, "NoAdditions"
            ToggleFormReadOnly Me, True, "NoAdditions"
        Else
            'if a filter is on remember no additions
            If Me.FilterOn Then
                ToggleFormReadOnly Me, False, "NoAdditions"
            Else
                ToggleFormReadOnly Me, False
            End If
            'unit number exists, lock field
            Me![Year].SetFocus
            Me![Unit Number].Locked = True
            Me![Unit Number].Enabled = False
            Me![Unit Number].BackColor = Me.Section(0).BackColor
        End If
        Me![lblMsg].Visible = False
    End If
End If

'current unit field always needs to be locked
Me![Text407].Locked = True


'priority button
If Me![Priority Unit] = True Then
    Me![Open Priority].Enabled = True
Else
    Me![Open Priority].Enabled = False
End If

'go to space button - new
If IsNull(Me![Space]) Or Me![Space] = "" Then
    Me![cmdGoToSpace].Enabled = False
Else
    Me![cmdGoToSpace].Enabled = True
End If

'go to building button - new
If IsNull(Me![Building]) Or Me![Building] = "" Then
    Me![cmdGoToBuilding].Enabled = False
Else
    Me![cmdGoToBuilding].Enabled = True
End If

' ** SAJ v9.1 -  reverse this logic - hide all subs to be reinstated by Case below - will stop flashing on/off
'restore all category forms
'Me![Exca: Unit Data Categories CUT subform].Visible = True
'Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
'Me![Exca: Unit Data Categories LAYER subform].Visible = True
Me![Exca: Unit Data Categories CUT subform].Visible = False
Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
Me![Exca: Unit Data Categories LAYER subform].Visible = False
Me![Exca: Unit Data Categories SKELL subform].Visible = False

'when skelli subform is shown these fields are hidden by it but still there
'so expecting tabbing - when skelli form shown the tab stop is set to false, set it back here
Me![Description].TabStop = True
Me![Recognition].TabStop = True
Me![Definition].TabStop = True
Me![Execution].TabStop = True
Me![Condition].TabStop = True

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
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False

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
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
    
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
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
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
    'SAJ v9.1 make this visible to make consistent with other forms of this nature
    Me![Exca: Unit Data Categories SKELL subform].Visible = True
    
    'when skelli subform is shown these fields are hidden by it but still there
    'so expecting tabbing - when skelli form shown the tab stop is set to false, set it back here
    Me![Description].TabStop = False
    Me![Recognition].TabStop = False
    Me![Definition].TabStop = False
    Me![Execution].TabStop = False
    Me![Condition].TabStop = False
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
    'SAJ v9.1 make this invisible - see case skeleton above
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
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
On Error GoTo err_Form_Open:

Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Then
        ' ToggleFormReadOnly Me, False ' on current will set it up for these users
    Else
        'set read only form here, just once
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
        Me![Unit Number].BackColor = Me.Section(0).BackColor
        Me![Unit Number].Locked = True
        Me![copy_method].Enabled = False
    End If
    
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        'disable find and add new in this instance
        Me![cboFindUnit].Enabled = False
        Me![cmdAddNew].Enabled = False
    End If

Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
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

Private Sub NotExcavated_Click()
'*********************************************************************
' Introduce logic to Not excavated option - if this is selected then
' Priority Unit and Fast track must be false
' SAJ v9.1
'*********************************************************************
On Error GoTo err_NotExcavated_Click
Dim checknum, checknum1, sql1

    If Me![NotExcavated] = True Then
    
        'this action means the Priority unit field cannot be checked - however
        'if there has been info entered about the priority then don't allow this action
        If Me![Priority Unit] = True Then
            'must first check unit number is in priority table as if its not (and due to no linking of this info
            'prior to this version this could have happened) the check for data will fail anyway
            checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number])
            If Not IsNull(checknum) Then
                'unit number there, do second check to see if data entered
                checknum1 = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number] & " AND [Exca: Priority Detail].Priority =1 AND [Exca: Priority Detail].Comment Is Null AND [Exca: Priority Detail].Discussion Is Null")
                If IsNull(checknum1) Then
                    'the unit has priority data so this change from priorty to not exca can go ahead
                    MsgBox "Sorry there is information relating to this Unit as a Priority, you cannot change this Unit to Not Excavated", vbExclamation, "Priority Information"
                    Me![NotExcavated] = False
                Else
                    'the unit has no priority specific data so can be removed from there
                    sql1 = "DELETE * FROM [Exca: Priority Detail] WHERE [Unit number] = " & Me![Unit Number] & ";"
                    DoCmd.RunSQL sql1
                    MsgBox "This Unit is no longer marked a Priority. This action has been allowed because it had no Priority specific information entered.", vbExclamation, "Priority change"
                    GoTo allow_check
                End If
            Else
                'unit number not in priorty table anyway so ok to uncheck
                GoTo allow_check
            End If
        Else
            'priority unit not set anyway so ok to uncheck this
            GoTo allow_check
        End If
        
    End If
Exit Sub

allow_check:
    Me![FastTrack] = False
    Me![Priority Unit] = False
Exit Sub

err_NotExcavated_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub Open_priority_Click()
'******************************************************************
' existing code, entered generic error trap v9.1
'
' add save record as otherwise newly entered records not appear in
' priority form
'
' pre this version this form was not linked to the checking of the
' priority on here - so must intro check to ensure the record does
' exist in Priority Details table, if not add it as it should. This
' problem will not happen in future, this is for old records
' SAJ v9.1
'******************************************************************
On Error GoTo Err_Open_priority_Click

    DoCmd.RunCommand acCmdSaveRecord
    
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, sql, permiss
    
    'check unit number already in Priority detail table
    checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Unit Number] = " & Me![Unit Number])
    If IsNull(checknum) Then
        'number not exist - now see what permissions user has
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            'no can't be found - add it as this user has permission to
            sql = "INSERT INTO [Exca: Priority Detail] ([Unit Number], [DateSet]) VALUES (" & Me![Unit Number] & ", #" & Date & "#);"
            DoCmd.RunSQL sql
        Else
            'user is readonly so just tell them record not exist
            MsgBox "Sorry but this unit record has not been added to the priority detail table yet, there is no record to view.", vbInformation, "Missing Priority Record"
        End If
    End If
    
    'now carry on and open form
    stDocName = "Exca: Priority Detail"
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Open_priority_Click:
    Exit Sub

Err_Open_priority_Click:
    Call General_Error_Trap
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
    'removed open excavation (menu) as now this form can be opened from other places
    'stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
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



Private Sub Priority_Unit_Click()
'*********************************************************************
' Introduce logic to Priority Unit option - if this is selected then
' Not Excavated must be false.
' Previous to this version there was no link between checking this box and
' the unit being recorded in the Priority Detail table - this link has been
' introduced.
' SAJ v9.1
'*********************************************************************
On Error GoTo err_Priority_Unit_Click
Dim checknum, checknum1, sql, sql1

    If Me![Priority Unit] = True Then
        Me![NotExcavated] = False
        'check unit num not already in priority details table - if not add it (this will catch any mismatches that are already there)
        checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Unit Number] = " & Me![Unit Number])
        If IsNull(checknum) Then
            sql = "INSERT INTO [Exca: Priority Detail] ([Unit Number], [DateSet]) VALUES (" & Me![Unit Number] & ", #" & Date & "#);"
            DoCmd.RunSQL sql
        End If
        Me![Open Priority].Enabled = True
    Else
        'priority units can only be unchecked if no data has been entered into record, having the priority
        'of 1 has to be taken as no data as this is the default assigned when record entered
        'so this will check if the unit record is there with a priority of 1 but no comment and no discussion
        
        'however first must check if number there are all otherwise check below will come back null
        'ie can't uncheck when you should be able to as record not exist in priority detail table
        checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number])
        If Not IsNull(checknum) Then
           'unit number there, do second check to see if data entered
            checknum1 = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number] & " AND [Exca: Priority Detail].Priority =1 AND [Exca: Priority Detail].Comment Is Null AND [Exca: Priority Detail].Discussion Is Null")
            If IsNull(checknum1) Then
                'the unit has prioirty data, it can be unallocated
                MsgBox "Sorry there is information relating to this Unit as a Priority, you cannot uncheck it", vbExclamation, "Priority Information"
                Me![Priority Unit] = True
            Else
                'no data so allow it to be removed from Priority detail table
                sql1 = "DELETE * FROM [Exca: Priority Detail] WHERE [Unit number] = " & Me![Unit Number] & ";"
                DoCmd.RunSQL sql1
                Me![Open Priority].Enabled = False
            End If
        Else
            'unit number not there in priority detail- so can uncheck
            Me![Open Priority].Enabled = False
        End If
        
    End If
Exit Sub

err_Priority_Unit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Space_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure space num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter space num when they can
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Space_AfterUpdate

Dim checknum, msg, retval, sql

If Me![Space] <> "" Then
    'first check its valid
    If IsNumeric(Me![Space]) Then
    
        'check that space num does exist
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
        If IsNull(checknum) Then
            msg = "This Space Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number]) VALUES ('" & Me![Space] & "');"
                DoCmd.RunSQL sql
                
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = '" & Me![Space] & "'", acFormEdit, acDialog
            End If
        Else
            'valid number, enable view button
            Me![cmdGoToSpace].Enabled = True
        End If
    
    Else
        'not a vaild space building number
        MsgBox "This Space number is not numeric, it cannot be checked for validity", vbInformation, "Non numeric Entry"
    End If
End If

Exit Sub

err_Space_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
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
            'oh the joys, to keep the focus on unit have to flip to year then back
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
