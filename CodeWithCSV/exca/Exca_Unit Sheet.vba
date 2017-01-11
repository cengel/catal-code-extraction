Option Explicit
Option Compare Database   'Use database order for string comparisons

Sub UpdateDataCategory()
'new 2008 this local sub deals with updating the data category table if the excavation status is set to
'void, natural or unstratified
On Error GoTo err_updatedatacategory
    'Me![Exca: Unit Data Categories CUT subform]![Data Category] = "cut"
    'could check if data category exists then update, if not insert but might be quicker just delete if there and reinsert
    Dim sql1
    'sql1 = "DELETE from [Exca: Unit Data Categories] where [Unit Number] = " & Me![Unit Number] & ";"
    'DoCmd.RunSQL sql1
    If spString <> "" Then
        Dim mydb As DAO.Database
        Dim myq1 As QueryDef
    
        Set mydb = CurrentDb
        Set myq1 = mydb.CreateQueryDef("")
        ''myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
        myq1.Connect = spString
    
            myq1.ReturnsRecords = False
            myq1.sql = "sp_Excavation_Delete_DataCategory_Entry " & Me![Unit Number]
            myq1.Execute
            
        myq1.Close
        Set myq1 = Nothing
        mydb.Close
        Set mydb = Nothing
    Else
        'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
        MsgBox "The data category record has not been deleted, please update it manually.", vbCritical, "Error"
    End If
    'MsgBox "data category removed"
    
    If Me![cboExcavationStatus] = "void" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'arbitrary', 'void (unused unit no)', '','','','');"
        DoCmd.RunSQL sql1
        'MsgBox "data cat inserted"
    ElseIf Me![cboExcavationStatus] = "natural" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'natural', '', '','','','');"
        DoCmd.RunSQL sql1
        'MsgBox "data cat inserted"
    ElseIf Me![cboExcavationStatus] = "unstratified" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'arbitrary', 'unstratified', '','','','');"
        DoCmd.RunSQL sql1
        'MsgBox "data cat inserted"
    End If
   
   'when undertook this action: chgn status to unstrat, change a data cat value, chg status to exca, chgn category from layer to cluster
   'was getting error 7887 data deleted by another user and #deleted on screen. This was happening when
   'category after update was trying to set datacategory to 'cluster' so I add requery of all here
   'and seems to work, bit over kill but I can't track how to define more (its 40 degrees and its 3.15pm on thursday end of season 2008! SAJ)
    Me![Exca: Unit Data Categories LAYER subform].Requery
    Me![Exca: Unit Data Categories CLUSTER subform].Requery
    Me![Exca: Unit Data Categories CUT subform].Requery
    Me![Exca: Unit Data Categories SKELL subform].Requery
    
    Me![Category] = Me![cboExcavationStatus]
    
    Call Form_Current 'update screen correctly
Exit Sub

err_updatedatacategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub Delete_Category_SubTable_Entry(deleteFrom, Unit)
'new 2008, when category is changed a record in either the cut or skeleton table might have to be
'deleted by RW users do not have delete permissions - get around this with a store proc
'pass in table to delete from and unit number
On Error GoTo err_delete_cat

If spString <> "" Then
    Dim mydb As DAO.Database
    Dim myq1 As QueryDef
    
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    ''myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
    myq1.Connect = spString
    
        myq1.ReturnsRecords = False
        myq1.sql = "sp_Excavation_Delete_Category_SubTable_Entry " & Unit & ", '" & deleteFrom & "'"
        myq1.Execute
            
    myq1.Close
    Set myq1 = Nothing
    mydb.Close
    Set mydb = Nothing
Else
    'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
    MsgBox "The " & deleteFrom & " record cannot be deleted, please restart the database, set this unit back to " & deleteFrom & " and try this change again", vbCritical, "Error"
End If

Exit Sub

err_delete_cat:
    Call General_Error_Trap
    Exit Sub
End Sub
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
    'MAY 2009 - in follow up to the new timeperiod field added make sure a default value is entered there
    'have to cater for possibility this could be an update
    If IsNull(Me![TimePeriod]) Then
        If Me![Mound] = "West" Then
            Me![TimePeriod] = "Chalcolithic"
        ElseIf Me![Mound] = "Off-Site" Then
            Me![TimePeriod] = "Unknown"
        Else
            Me![TimePeriod] = "Neolithic"
        End If
    Else
        'timeperiod was not empty so check with user
        Dim response
        If Me![Mound] = "West" And Me![TimePeriod] <> "Chalcolithic" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Chalcolithic, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Chalcolithic"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        ElseIf Me![Mound] = "East" And Me![TimePeriod] <> "Neolithic" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Neolithic, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Neolithic"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        ElseIf Me![Mound] = "Off-Site" And Me![TimePeriod] <> "Unknown" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Unknown, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Unknown"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        End If
    End If
End If

'new 22nd Aug 2008 SAJ - new FT field with a pull down list dependant on area
If Me![Area] <> "" Then
    Me![cboFT].RowSource = "SELECT [Exca: Foundation Trench Description].FTName, [Exca: Foundation Trench Description].Description, [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder FROM [Exca: Foundation Trench Description] WHERE [Area] = '" & Me![Area] & "' ORDER BY [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder;"
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

'v11.2 2008 - this still allows a user to change category and doesn't clean up the sub table.
'have spent time cleaning this up so now code to stop it happening. I planned to implement this
'so that it would only flag a message if there was actual data in the subtable but as this requires
'RS check it would be slower than simple does a record exist dlookup - so stuck with that.
If Me![Category].OldValue <> "" Or Not IsNull(Me![Category]) Then
    'what is the old category? check its not a change cluster - layer, layer - cluster as this doesn't matter is same sub table
    If Not ((Me![Category].OldValue = "cluster" Or Me![Category].OldValue = "layer") And (Me![Category] = "cluster" Or Me![Category] = "layer")) Then
    
        Dim checkit
        checkit = Null
        If Me![Category].OldValue = "cut" Then 'check for cut info
            checkit = DLookup("[Unit Number]", "[Exca: Descriptions Cut]", "[Unit Number] = " & Me![Unit Number])
        ElseIf Me![Category].OldValue = "layer" Or Me![Category].OldValue = "cluster" Then
            checkit = DLookup("[Unit Number]", "[Exca: Descriptions Layer]", "[Unit Number] = " & Me![Unit Number])
        ElseIf Me![Category].OldValue = "skeleton" Then
            checkit = DLookup("[Unit Number]", "[Exca: Skeleton Data]", "[Unit Number] = " & Me![Unit Number])
        End If
    
        If Not IsNull(checkit) Then
            'there was a record
            Dim resp, sql
            resp = MsgBox("By changing the category of this Unit you will lose the " & Me![Category].OldValue & " specific data (if any). Do you still want to change the category?", vbQuestion + vbYesNo, "Confirm Action")
            If resp = vbNo Then
                Me![Category] = Me![Category].OldValue
            ElseIf resp = vbYes Then
                'must delete sub table info
                If Me![Category].OldValue = "layer" Or Me![Category].OldValue = "cluster" Then
                    'sql = "DELETE FROM [Exca: Descriptions Layer] WHERE [Unit Number] = " & Me![Unit Number] & ";"
                    'DoCmd.RunSQL sql
                    Call Delete_Category_SubTable_Entry("layer", Me![Unit Number])
                ElseIf Me![Category].OldValue = "cut" Then
                    'sql = "DELETE FROM [Exca: Descriptions Cut] WHERE [Unit Number] = " & Me![Unit Number] & ";"
                    'DoCmd.RunSQL sql
                    Call Delete_Category_SubTable_Entry("cut", Me![Unit Number])
                ElseIf Me![Category].OldValue = "skeleton" Then
                    'sql = "DELETE FROM [Exca: Skeleton Data] WHERE [Unit Number] = " & Me![Unit Number] & ";"
                    'DoCmd.RunSQL sql
                    Call Delete_Category_SubTable_Entry("skeleton", Me![Unit Number])
                End If
            End If
    
        End If
    End If
End If

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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False

    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False

Case "skeleton"
    'added the if statement to accommodate for 2013 new skeleton sheet
    If Forms![Exca: Unit Sheet]!Year < 2013 Then
        'MsgBox "I will give you the old form for entry"

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
        Me![Exca: subform Skeleton Sheet 2013].Visible = False
        Me![subform Unit: stratigraphy  same as].Visible = False
        Me![Exca: Subform Layer descr].Visible = False
        Me![Exca: Subform Cut descr].Visible = False
        Me![Exca: subform Skeletons same as].Visible = True

        'SAJ v9.1 make this visible to make consistent with other forms of this nature
        Me![Exca: Unit Data Categories SKELL subform].Visible = True
    Else
        'MsgBox "I will give you the new form for entry"
        
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
        Me![Exca: subform Skeleton Sheet].Visible = False
        Me![Exca: subform Skeleton Sheet 2013].Visible = True
        Me![subform Unit: stratigraphy  same as].Visible = False
        Me![Exca: Subform Layer descr].Visible = False
        Me![Exca: Subform Cut descr].Visible = False
        Me![Exca: subform Skeletons same as].Visible = True

        'SAJ v9.1 make this visible to make consistent with other forms of this nature
        Me![Exca: Unit Data Categories SKELL subform].Visible = True

    End If
End Select
Exit Sub

Err_Category_AfterUpdate:
    Call General_Error_Trap
    
    Exit Sub
End Sub





Private Sub cboExcavationStatus_AfterUpdate()
'new 2008 - if the user sets the status to anything other than excavated and not excavated the value will
'be written into the category field as well. This ensures the unit is clearly labelled as void, natural, unstrat etc
'but have to check no info was added for layer/cut/skel etc
On Error GoTo err_cboExcaStatus

    If Me![cboExcavationStatus] <> "excavated" And Me![cboExcavationStatus] <> "not excavated" And Me![cboExcavationStatus] <> "partially excavated" Then
        'take action with the category field
        If Me![Category] = "" Or IsNull(Me![Category]) Then
            'no existing category so simply update to exca status
            Me![Category] = Me![cboExcavationStatus]
            Me![Category].Locked = True
            Me![Category].Enabled = False
        Else
            If Me![Category] = "cut" Or Me![Category] = "skeleton" Then
                'anything other than layer table data must be removed as this is the default view for unknown categories
                Dim checkit
                checkit = Null
                If Me![Category] = "cut" Then 'check for cut info
                    checkit = DLookup("[Unit Number]", "[Exca: Descriptions Cut]", "[Unit Number] = " & Me![Unit Number])
                ElseIf Me![Category].OldValue = "skeleton" Then
                    checkit = DLookup("[Unit Number]", "[Exca: Skeleton Data]", "[Unit Number] = " & Me![Unit Number])
                End If
    
                If Not IsNull(checkit) Then
                    'there was a record
                    Dim resp, sql
                    resp = MsgBox("By changing the status of this Unit you will lose the " & Me![Category].OldValue & " specific data (if any). Do you still want to change the status?", vbQuestion + vbYesNo, "Confirm Action")
                    If resp = vbNo Then
                        Me![cboExcavationStatus] = Me![cboExcavationStatus].OldValue
                    ElseIf resp = vbYes Then
                        'must delete sub table info
                        If Me![Category] = "cut" Then
                            'sql = "DELETE FROM [Exca: Descriptions Cut] WHERE [Unit Number] = " & Me![Unit Number] & ";"
                            'DoCmd.RunSQL sql
                            Call Delete_Category_SubTable_Entry("cut", Me![Unit Number])
                            'MsgBox "record was deleted from cut"
                            
                            Call UpdateDataCategory 'local sub
                                                        
                        ElseIf Me![Category] = "skeleton" Then
                            'sql = "DELETE FROM [Exca: Skeleton Data] WHERE [Unit Number] = " & Me![Unit Number] & ";"
                            'DoCmd.RunSQL sql
                            Call Delete_Category_SubTable_Entry("skeleton", Me![Unit Number])
                            'MsgBox "record was deleted from skeleton"
                            
                            Call UpdateDataCategory 'local sub
                        End If
                    End If
                Else
                    'its ok to change category
                    Call UpdateDataCategory 'local sub
                End If
                Me![Category].Locked = True
                Me![Category].Enabled = False
            Else 'If Me![Category] = "void" Or Me![Category] = "natural" Or Me![Category] = "unstratified" Then
                'must simply sort out data category fields
                Call UpdateDataCategory 'local sub
            End If
        End If
    Else
        'season 2009 - last year I wrote this card and it was REALLY hard to capture all the change scenarios so I'm wary of altering it
        'but Lisa pointed out that if you change not excavated to excavated you loose the data categories which would be via
        'this bit of code. So I've changed the OR to AND so it only runs when the status is not "not excavated" and not "excavated"
        'to leave the data categories alone - the excavators are going to test it to check it works ok in diff scenarios
        'If Me![cboExcavationStatus].OldValue <> "Excavated" Or Me![cboExcavationStatus].OldValue <> "Not Excavated" Then
        If Me![cboExcavationStatus].OldValue <> "Excavated" And Me![cboExcavationStatus].OldValue <> "Not Excavated" Then
            Call UpdateDataCategory
            Me![Category] = "Layer" 'set default to layer
            Me![Category].Locked = False
            Me![Category].Enabled = True
        Else
            'MsgBox "here"
        End If
    End If

Exit Sub

err_cboExcaStatus:
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
        '2009 move the blank down to after the gotocontrol as code in Year_lostfocus checks
        'for a value when the year looses it - complicated passing of focus nightmare stuff
        'Me![cboFindUnit] = ""
        '2009 focus will bounce on Year fld and can be easily over written to make sure stays here
        DoCmd.GoToControl "cboFindUnit"
        Me![cboFindUnit] = ""
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - 2009
On Error GoTo err_cbofindNot

    MsgBox "Sorry this Unit cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFindUnit].Undo
    '2009 if not esc the list will stay pulled down making it hard to go direct to Add new or where ever as
    'have to escape the pull down list first
    SendKeys "{ESC}"
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFT_AfterUpdate()
If Me![cboFT] <> "" Then
    Me![cmdGoToFT].Enabled = True
End If


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

Private Sub cmdEditPhase_Click()
'new 2009 - phasing can be for a building or a space so locked field on unit sheet and
'will let this button do the hard work of checking out which (don't want more going on with OnCurrent
'of unit sheet). SAJ
On Error GoTo err_edit

If Not IsNull(Me![Unit Number]) Then
    'check if this unit has a building number
    Dim checkB, checkSp, getB, getSp, counter, response
    checkB = DCount("[In_Building]", "[Exca: Units in Buildings]", "[Unit] = " & Me![Unit Number])
    If checkB = 0 Then
        'this unit is not associated with a building, is it associated with a space?
        checkSp = DCount("[In_Space]", "[Exca: Units in Spaces]", "[Unit] = " & Me![Unit Number])
        If checkSp = 0 Then
            MsgBox "This unit is not associated with a Building or a Space so it cannot be phased in this way", vbInformation, "Nothing to Phase"
            Exit Sub
        Else
            If checkSp > 1 Then
                'unit associated with > 1 space - which space does the user want to phase at this time?
                counter = 1
                Me![Exca: subform  Features in Spaces].Form.RecordsetClone.MoveFirst
                Do Until counter > checkSp
                    
                    'get the space numbers associate with this unit from the RS clone, the form name for the Units in Space subform is actually Features in Spaces on
                    'this form which is messy but the consequnces of a change are not worth investigating, in RS clone field 1 is the unit, field 2 is the space and 3 is the date changed
                    'so loop the RS clone getting the space numbers and asking user do they want to phase each space
                    response = MsgBox("Do you want to phase this unit to Space " & Me![Exca: subform  Features in Spaces].Form.RecordsetClone(1).Value & "?" & _
                                Chr(13) & Chr(13) & "Clicking No will prompt the question for the next Space in the list if there are more.", vbQuestion + vbYesNoCancel, "Which Space to phase now?")
                    If response = vbYes Then
                        DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: SpacePhases].SpacePhase FROM [Exca: SpacePhases] WHERE [Exca: SpacePhases].SpaceNumber=" & Me![Exca: subform  Features in Spaces].Form.RecordsetClone(1).Value & ";" 'open form with space number
                        Exit Do
                    ElseIf response = vbCancel Then
                        Exit Do
                    End If
                    'MsgBox Me![Exca: subform  Features in Spaces].Form.RecordsetClone(1).Value
                    'move on
                    Me![Exca: subform  Features in Spaces].Form.RecordsetClone.MoveNext
                    counter = counter + 1
                Loop
                'MsgBox "This unit is associated with more than one Space. Currently the system does not support phasing a unit to more than one space. Please discuss this with Shahina.", vbInformation, "Multiple Space Numbers"
                
                'Exit Sub
            Else
                'phase to a space
                getSp = DLookup("[In_Space]", "[Exca: Units in Spaces]", "[Unit] = " & Me![Unit Number])
                DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: SpacePhases].SpacePhase FROM [Exca: SpacePhases] WHERE [Exca: SpacePhases].SpaceNumber=" & getSp & ";" 'open form with space number
            End If
        End If
    Else
        If checkB > 1 Then
            MsgBox "This unit is associated with more than one Building. Currently the system does not support phasing a unit to more than one building. Please discuss this with Shahina.", vbInformation, "Multiple Building Numbers"
            Exit Sub
        Else
            'phase to a building
            getB = DLookup("[In_Building]", "[Exca: Units in Buildings]", "[Unit] = " & Me![Unit Number])
            DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: BuildingPhases].BuildingPhase FROM [Exca: BuildingPhases] WHERE [Exca: BuildingPhases].BuildingNumber=" & getB & ";" 'open form with building number
        End If
    End If
    Me![Exca: subform Units Occupation Phase].Requery
End If

Exit Sub

err_edit:
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
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
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

Private Sub cmdGoToFT_Click()
'***********************************************************************
' Open FT form read only from here.

' SAJ v11.1
'***********************************************************************
On Error GoTo Err_cmdGoToFT_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    
    stDocName = "Exca: Admin_Foundation_Trenches"
    
    If Not IsNull(Me![cboFT]) Or Me![cboFT] <> "" Then
        'record exists - open it
        stLinkCriteria = "[FTName]='" & Me![cboFT] & "' AND [Area] = '" & Me![Area] & "'"
           
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        
    Else
        MsgBox "No FT record to view", vbInformation, "No FT Name"
    End If
    
Exit_cmdGoToFT_Click:
    Exit Sub


Err_cmdGoToFT_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToFT_Click
    
End Sub

Private Sub cmdGoToImage_Click()
'********************************************************************
' New button for version 9.1 which allows any available images to be
' displayed - links to the Image_Metadata table that has been exported
' from Portfolio
' SAJ v9.1
'********************************************************************
On Error GoTo err_cmdGoToImage_Click

Dim mydb As DAO.Database
Dim tmptable As TableDef, tblConn, I, msg, fldid
Set mydb = CurrentDb

    'get the field id for unit in the catalog that matches this year
    'NEW 2007 method where by portfolio now uses its own sql database
    Dim myq1 As QueryDef, connStr
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = mydb.TableDefs("view_Portfolio_Previews_2008").Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    'myq1.sql = "sp_Portfolio_GetUnitFieldID " & Me![Year]
    myq1.sql = "sp_Portfolio_GetUnitFieldID_2008 " & Me![Year]
    
    Dim myrs As Recordset
    Set myrs = myq1.OpenRecordset
    ''MsgBox myrs.Fields(0).Value
    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
        fldid = 0
    Else
        fldid = myrs.Fields(0).Value
    End If
        
    myrs.Close
    Set myrs = Nothing
    myq1.Close
    Set myq1 = Nothing
    
    
    
    For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
             
    If tmptable.Connect <> "" Then
        tblConn = tmptable.Connect
        Exit For
    End If
    Next I
    
    If tblConn <> "" Then
        'If InStr(tblConn, "catalsql") = 0 Then
        If InStr(tblConn, "catalsql") = 0 Then
            'if on site the image can be loaded from the server directly into Access
            'DoCmd.OpenForm "Image_Display", acNormal, , "[Unit] = '" & Me![Unit Number] & "'", acFormReadOnly, acDialog
            '2007 - need year to define catalog
            'DoCmd.OpenForm "Image_Display", acNormal, , "[Unit] = '" & Me![Unit Number] & "'", acFormReadOnly, acDialog, Me![Year]
            'DoCmd.OpenForm "Image_Display", acNormal, , "[StringValue] = '" & Me![Unit Number] & "' AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
            'DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Unit Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Unit Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
            
        Else
            'database is running remotely must access images via internet
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=unit&id=" & Me![Unit Number])
        End If

    Else
        
    End If
    
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
    
Exit Sub

err_cmdGoToImage_Click:
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
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
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

Private Sub cmdPrintUnitSheet_Click()
On Error GoTo err_print

    If LCase(Me![Category]) = "layer" Or LCase(Me![Category]) = "cluster" Then
        DoCmd.OpenReport "R_Unit_Sheet_layercluster", acViewPreview, , "[unit number] = " & Me![Unit Number]
    ElseIf LCase(Me![Category]) = "cut" Then
        DoCmd.OpenReport "R_Unit_Sheet_cut", acViewPreview, , "[unit number] = " & Me![Unit Number]
    ElseIf LCase(Me![Category]) = "skeleton" Then
        DoCmd.OpenReport "R_Unit_Sheet_skeleton", acViewPreview, , "[unit number] = " & Me![Unit Number]
    End If
Exit Sub

err_print:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReportProblem_Click()
'bring up a popup to allow user to report a problem
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "unit number;" & Me![Unit Number]
    
Exit Sub

err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypeLOV_Click()
'2009 bring up type list so people can see it
On Error GoTo err_typeLOV

    DoCmd.OpenForm "Frm_subform_sampletypeLOV", acNormal, , , acFormReadOnly, acDialog
    

Exit Sub

err_typeLOV:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewSketch_Click()
'new season 2009 - open the diary sketch
On Error GoTo err_opensketch
    
    ''LATE AUGUST 2009 SEASON
    ''Due to overwork of the forms OnCurrent event the check for the existence of a sketch plan
    ''now happens when the user clicks this button
    Dim Path
    Dim fname

    'check if can find sketch image
    'using global constanst Declared in globals-shared
    'path = "\\catal\Site_Sketches\"
    
    If Me![Year] < 2015 Then
    Path = sketchpath
    Path = Path & Me![Unit Number] & ".jpg"
    Else
    Path = sketchpath2015
    Path = Path & "units\sketches\" & "U" & Me![Unit Number] & "*" & ".jpg"
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        Debug.Print fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "units\sketches\" & fname
    End If
    
    If Dir(Path) = "" Then
        'directory not exist
        MsgBox "The sketch plan of this unit has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_graphic", acNormal, , , acFormReadOnly, , Me![Unit Number]
    End If
    'DoCmd.OpenForm "frm_pop_graphic", acNormal, , , acFormReadOnly, , Me![Unit Number]

Exit Sub

err_opensketch:
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



Private Sub Exca__subform_Skeleton_Sheet_Enter()

End Sub

Private Sub FastTrack_Click()
'*********************************************************************
' Introduce logic to fast track option - if this is selected then
' Not excavated must be false
' SAJ v9.1
'*********************************************************************
On Error GoTo err_FastTrack_Click
    '2009 this was still here even the field had gone
    'If Me![FastTrack] = True Then
    '    Me![NotExcavated] = False
    'End If
''7July2010 - change rquest Cord - always enable remaining vol
''all below commented out
''    Dim permiss
''    permiss = GetGeneralPermissions
''    'season 2010 to protect Cords cleaning in 2010 keep all volume fields lockd for units 2003 - 2008
''    If (Me![Area] <> "4040" And Me![Area] <> "South") Or (permiss = "ADMIN" Or (Me![Year] < 2003)) Then
''        'season 2009 tie this in with Unsieved Volume field (Remaining Volume in database)
''
''       If Me![FastTrack] = True Then
''            Me![RemainingVolume].Enabled = True
''       Else
''           'must check if there is anything in unsieved volume, if there is don't allow fast track to be turned off until sorted
''           If Me![RemainingVolume] <> 0 And Not IsNull(Me![RemainingVolume]) And Me![RemainingVolume] <> "" Then
''                MsgBox "This unit records an Unsieved volume figure. There is no unsieved volume for a non fast track unit so please sort this out first (if it is invalid please remove the unsieved volume figure)", vbExclamation, "Volume Problem"
''              Me![FastTrack] = True
''           Else
''               Me![RemainingVolume].Enabled = False
''            End If
''       End If
''
''    End If
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
'
'22Aug2008
'changes to include FT, excavationstatus and phase field
'29June2009
'exsuper permissions group introduced
'6July2009
'check site sketch exists - does this slow down movement between records
'6July2010
'! Cord has mass cleaned volumes so to ensure they stay clean for general edits the volume fields
' 2003 - 2008 will be locked. Only admins can edit
'*************************************************************************************
Dim stDocName As String
Dim stLinkCriteria As String
    
On Error GoTo err_Form_Current
    
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And ([Unit Number] <> 0 Or IsNull([Unit Number])) Then
    'check that unit not 0 added 26/3/09 bu SAJ in attempt to keep unit 0 locked to prevent overwrite
    'see also code addition in onopen.
    'if no unit number set all fields readonly
    If IsNull(Me![Unit Number]) Or Me![Unit Number] = "" Then
        'make rest of fields read only
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
                'SAJ season 2006 - don't allow deletions from this screen
                ToggleFormReadOnly Me, False, "NoDeletions"
            End If
            'unit number exists, lock field
            'Me![Year].SetFocus
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

'go to space button - new - replaced by subform v9.2
'If IsNull(Me![Space]) Or Me![Space] = "" Then
'    Me![cmdGoToSpace].Enabled = False
'Else
'    Me![cmdGoToSpace].Enabled = True
'End If

'go to building button - new - replaced by subform v9.2
'If IsNull(Me![Building]) Or Me![Building] = "" Then
'    Me![cmdGoToBuilding].Enabled = False
'Else
'    Me![cmdGoToBuilding].Enabled = True
'End If

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

'OLD IMAGE LINK METHOD WHERE DATA IMPORTED INTO SQL
'find out is any images available
Dim imageCount, Imgcaption
'imageCount = DCount("[Unit]", "Image_Metadata", "[Unit] = '" & Me![Unit Number] & "'")
'image metadata now normalised
'imageCount = DCount("[Unit]", "Image_Metadata_Units", "[Unit] = '" & Me![Unit Number] & "'")
'If imageCount > 0 Then
'    Imgcaption = imageCount
'    If imageCount = 1 Then
'        Imgcaption = Imgcaption & " Image to Display"
'    Else
'        Imgcaption = Imgcaption & " Images to Display"
'    End If
'    Me![cmdGoToImage].Caption = Imgcaption
'    Me![cmdGoToImage].Enabled = True
'Else
'    Me![cmdGoToImage].Caption = "No Image to Display"
'    Me![cmdGoToImage].Enabled = False
'End If
    
''LATE AUGUST 2009 SEASON
''We have recurring Error 52 Bad File name messages popping up until user UpdateDatabases, it will work a while
''and then reappear - is this related to this network call = timeout/corruption? Taking it out for now
''to see, when user presses button they will take pot luck on there being images
''NEW 2007 method where by portfolio now uses its own sql database
''commented out until portfolio pics up on web - sept 2007
''Dim mydb As DAO.Database
''Dim myq1 As QueryDef, connStr
''    Set mydb = CurrentDb
''    Set myq1 = mydb.CreateQueryDef("")
''    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
''    myq1.ReturnsRecords = True
    'myq1.sql = "sp_Portfolio_CountImagesForUnit '" & Me![Unit Number] & "', " & Me![Year]
    'bug - if no year odbc call failed
''    If IsNull(Me![Year]) Then
''        myq1.sql = "sp_Portfolio_CountImagesForUnit_2008 '" & Me![Unit Number] & "', ''"
''    Else
''        myq1.sql = "sp_Portfolio_CountImagesForUnit_2008 '" & Me![Unit Number] & "', " & Me![Year]
''    End If
''    Dim myrs As Recordset
''    Set myrs = myq1.OpenRecordset
    ''MsgBox myrs.Fields(0).Value
''    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
''           imageCount = 0
''    Else
''        imageCount = myrs.Fields(0).Value
''   End If

''myrs.close
''Set myrs = Nothing
''myq1.close
''Set myq1 = Nothing
''mydb.close
''Set mydb = Nothing
    
''If imageCount > 0 Then
''    Imgcaption = imageCount
''    If imageCount = 1 Then
''        Imgcaption = Imgcaption & " Image to Display"
''    Else
''        Imgcaption = Imgcaption & " Images to Display"
''    End If
    'new 2009 to avoid call to portfolio tables
    Imgcaption = "Images of Unit"
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
''Else
''    Me![cmdGoToImage].Caption = "No Image to Display"
''    Me![cmdGoToImage].Enabled = False
''End If
    
Dim Path

'check if can find sketch image
'using global constanst Declared in globals-shared
'path = "\\catal\Site_Sketches\"
Path = sketchpath

Path = Path & Me![Unit Number] & ".jpg"
    
    ''LATE AUGUST 2009 SEASON - see above portfolio link for reason
    ''If Dir(path) = "" Then
        'directory not exist
    ''    Me!cmdViewSketch.Enabled = False
    ''Else
        Me!cmdViewSketch.Enabled = True
    ''End If
   
   
'''OFFSITE 2009 - ignore photos and sketches offsite
'''JUST TAKE THESE TWO LINES OUT ON SITE TO RETRIEVE FUNCTIONALITY
''Me![cmdGoToImage].Enabled = False
''Me!cmdViewSketch.Enabled = False

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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    'SAJ v9.1 make this invisible - see case skeleton below
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "skeleton"
    ' added if statement 2013 to accommodate for new skeleton sheet
    If Me![Year] < 2013 Then
    'MsgBox "I will give you the old form for display"
        'data
        Me![Exca: Unit Data Categories CUT subform].Visible = False
        Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
        Me![Exca: Unit Data Categories LAYER subform].Visible = False
        'SAJ v9 update of field restricted to category_afterupdate
        'Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
        Me.refresh
    
        'skelli
        Me![Exca: subform Skeleton Sheet].Visible = True
        Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Else
     '    MsgBox "I will give you the new form for display"
        'data
        Me![Exca: Unit Data Categories CUT subform].Visible = False
        Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
        Me![Exca: Unit Data Categories LAYER subform].Visible = False
        'SAJ v9 update of field restricted to category_afterupdate
        'Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
        Me.refresh
    
        'skelli
        Me![Exca: subform Skeleton Sheet].Visible = False
        Me![Exca: subform Skeleton Sheet 2013].Visible = True
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
    End If
    
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
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    'SAJ v9.1 make this invisible - see case skeleton above
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
End Select

'new 22nd Aug 2008 SAJ - new FT field with a pull down list dependant on area
If Me![Area] <> "" Then
    Me![cboFT].RowSource = "SELECT [Exca: Foundation Trench Description].FTName, [Exca: Foundation Trench Description].Description, [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder FROM [Exca: Foundation Trench Description] WHERE [Area] = '" & Me![Area] & "' ORDER BY [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder;"
End If
'FT button
If Me![cboFT] <> "" Then
    Me![cmdGoToFT].Enabled = True
Else
    Me![cmdGoToFT].Enabled = False
End If

'lock this field for everyone bar admin
If (permiss = "ADMIN" Or permiss = "exsuper") Then
    Me![cboFT].Locked = False
    Me![cboFT].Enabled = True
    'Me![Phase].Locked = False 'now controlled by button
    'Me![Phase].Enabled = True
    Me![cmdEditPhase].Enabled = True
    Me![cboTimePeriod].Locked = False
    Me![cboTimePeriod].Enabled = True
Else
    Me![cboFT].Locked = True
    Me![cboFT].Enabled = False
    'Me![Phase].Locked = True 'now controlled by button
    'Me![Phase].Enabled = False
    Me![cmdEditPhase].Enabled = False
    Me![cboTimePeriod].Locked = True
    Me![cboTimePeriod].Enabled = False
End If

'new exca status field effects category availabilty
If Me![cboExcavationStatus] <> "excavated" And Me![cboExcavationStatus] <> "not excavated" Then
    'Me![Category].Locked = True
    Me![Category].Enabled = False
Else
    'Me![Category].Locked = False
    Me![Category].Enabled = True
End If

If permiss <> "ADMIN" And (Me![Year] >= 2003 And Me![Year] <= 2008 And (Me![Area] = "4040" Or Me![Area] = "South")) Then
    'new 6July2010 - locked volume fields
    Me![TotalSampleAmount].Enabled = False
    Me![Dry sieve volume].Enabled = False
    Me![RemainingVolume].Enabled = False
    Me![TotalDepositVolume].Enabled = False
    Me![HowVolumeCalc].Enabled = False
Else
    Me![TotalSampleAmount].Enabled = True
    Me![Dry sieve volume].Enabled = True
    
    Me![TotalDepositVolume].Enabled = True
    Me![HowVolumeCalc].Enabled = True

    ''7July2010 - change rquest Cord - always enable remaining vol
    'original volumes enabling
    ''If Me![FastTrack] = True Then
        Me![RemainingVolume].Enabled = True
    ''Else
    ''    Me![RemainingVolume].Enabled = False
    ''End If
End If
'new 2010 - show level and hodder phase if is one
'Dim getHP
'getHP = DLookup("[HodderPhase]", "[Exca: unit sheet with relationships]", "[unit number] = " & Me![Unit Number])
'Me![txtHodderPhase] = getHP

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
    If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And ([Unit Number] <> 0) Then
        'check that unit not 0 added 26/3/09 bu SAJ in attempt to keep unit 0 locked to prevent overwrite
        'see also code addition in oncurrent and in else below.
        ' ToggleFormReadOnly Me, False ' on current will set it up for these users
    Else
        'set read only form here, just once
        ToggleFormReadOnly Me, True
        '26/3/09 code might be in here if unit is 0 but user still has RW, in this case
        'must have addnew button available so only disable if it not one of these
        If permiss <> "ADMIN" And permiss <> "RW" And permiss <> "exsuper" Then
            Me![cmdAddNew].Enabled = False
        ElseIf permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            Me.AllowAdditions = True 'this ensures rw can add record right from start
        End If
        Me![Unit Number].BackColor = Me.Section(0).BackColor
        Me![Unit Number].Locked = True
        Me![copy_method].Enabled = False
    End If
    
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        'disable find and add new in this instance
        Me![cboFindUnit].Enabled = False
        Me![cmdAddNew].Enabled = False
    Else
        'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
        DoCmd.GoToControl "cboFindUnit"

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
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
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



Private Sub print_bulk_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_print_bulk_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "print_bulk_units"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
   '' REMOVED SAJ AS WILL FAIL FOR RO USERS DoCmd.GoToRecord acForm, stDocName, acNewRec

Exit_print_bulk_Click:
    Exit Sub

Err_print_bulk_Click:
    Call General_Error_Trap
    Resume Exit_print_bulk_Click
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
        'picked up in 2009 as field is no longer here
        'Me![NotExcavated] = False
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
        checknum = DLookup("[Category]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
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

'if after checks the field has a value hide the enter number msg
If Me![Unit Number] <> "" Then Me![lblMsg].Visible = False
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

Private Sub Year_AfterUpdate()
'TRYING THIS CODE IN LOST FOCUS TO TRY MAKE IT FOOL PROOF
'Season 2009 - fed up of cleaning up year entries - putting a message to indicate an invalid year
'not fool proof as after this can simply tab off the control with no consequence, putting the code in on lost focus
'If IsNull(Me![Year]) Or Me![Year] = "" Then
'    MsgBox "You must enter the year this unit number was excavated, or allocated if not excavated yet", vbInformation, "Invalid Year"
'    DoCmd.GoToControl "Area"
'    DoCmd.GoToControl "Year"
'    Me![Year].SetFocus
'ElseIf Me![Year] < 1993 Or Me![Year] > ThisYear Then
'    MsgBox Me![Year] & "is not a valid Year please try again", vbInformation, "Invalid Year"
'    'Me![Year] = ""
'    DoCmd.GoToControl "Area"
'    DoCmd.GoToControl "Year"
'    Me![Year].SetFocus
'End If
End Sub


Private Sub Year_LostFocus()
'Season 2009 - fed up of cleaning up year entries - putting a message to indicate an invalid year
'to stop it popping up and causing an error after a search to an existing entry to a unit with any invalid
'year eg: 999999 that has no year put a check in to see if there is value in cboFindUnit
'focus moving - bit of a tangle hence the error trap simply ignores it at present
On Error GoTo err_Year

    If IsNull(Me![cboFindUnit]) Then
        'not a search

        If IsNull(Me![Year]) Or Me![Year] = "" Then
            MsgBox "You must enter the year this unit number was excavated, or allocated if not excavated yet", vbInformation, "Invalid Year"
            DoCmd.GoToControl "Area"
            DoCmd.GoToControl "Year"
            Me![Year].SetFocus
        ElseIf Me![Year] < 1993 Or Me![Year] > ThisYear Then
            MsgBox Me![Year] & " is not a valid Year please try again", vbInformation, "Invalid Year"
            'Me![Year] = ""
            DoCmd.GoToControl "Area"
            DoCmd.GoToControl "Year"
            Me![Year].SetFocus
        End If
    End If
Exit Sub

err_Year:
    'ignore it
    Exit Sub

End Sub
