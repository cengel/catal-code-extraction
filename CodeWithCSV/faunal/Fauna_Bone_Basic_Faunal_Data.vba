Option Compare Database   'Use database order for string comparisons
Option Explicit 'added by saj

Dim WhereGo 'SAJ this var is used to track where the user goes to when they close form, used by form deactivate

Private Sub UpdateGID()
'this is called by this form to update the GID field if either the
'Unit, find letter or number fields are altered
'replaces calls to macro bone.Update GID
' SAJ
On Error GoTo err_UpdateGID

 Me![GID] = [Forms]![Fauna_Bone_Basic_Faunal_Data]![Unit number] & "." & [Forms]![Fauna_Bone_Basic_Faunal_Data]![letter code] & [Forms]![Fauna_Bone_Basic_Faunal_Data]![find number]
    

Exit Sub

err_UpdateGID:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub button_goto_unitdescription_Click()
On Error GoTo Err_button_goto_unitdescription_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
If Me![Unit number] <> "" Then
    'new check for unit number entered by saj
    stDocName = "Fauna_Bone_Faunal_Unit_Description"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    
        
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter or select a Unit Number first", vbInformation, "No Unit Number"
End If

Exit_button_goto_unitdescription_Click:
    Exit Sub

Err_button_goto_unitdescription_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_unitdescription_Click

End Sub
Sub button_goto_cran_postcran_Click()
On Error GoTo Err_button_goto_cran_postcran_Click

'new season 2006 - track movement
WhereGo = "Post/Cran"

    Dim stDocCranial As String
    Dim stDocPostCranial As String
    Dim stLinkCriteria As String
    Dim stElementType As String 'was not declared before option explicit SAJ
    Dim checknum, sql
    
    stDocCranial = "Fauna_Bone_Cranial"
    stDocPostCranial = "Fauna_Bone_Postcranial"
    stElementType = "Fauna_Bone_Basic_Faunal_Data.Field40"
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
    ' MR July 18, 2005
    'DoCmd.Save 'commented out by saj placed below
    
    If Me![GID] <> "" Then
        'new check for GID entered by saj
        'the form is only minimised so must save data manually here - saj
        DoCmd.RunCommand acCmdSaveRecord
        'new 2006 saj - dont let go anywhere until picked element
        If Field40 <> "" Then
            If Field40 < 24 Then
                'new for season 2006 - see if the cranial record exists if not create it - SAJ
                checknum = DLookup("[GID]", "[Fauna_Bone_Cranial]", "[GID] = '" & Me![GID] & "'")
                If IsNull(checknum) Then
                    'gid not exist there yet so create it ready for data entry
                    sql = "INSERT INTO [Fauna_Bone_Cranial] ([GID], [Unit number], [Letter code], [Find number]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ");"
                    DoCmd.RunSQL sql
                End If
            
                DoCmd.Minimize
                DoCmd.OpenForm stDocCranial, , , stLinkCriteria
            Else
                'new for season 2006 - see if the cranial record exists if not create it - SAJ
                checknum = DLookup("[GID]", "[Fauna_Bone_PostCranial]", "[GID] = '" & Me![GID] & "'")
                If IsNull(checknum) Then
                    'gid not exist there yet so create it ready for data entry
                    sql = "INSERT INTO [Fauna_Bone_PostCranial] ([GID], [Unit number], [Letter code], [Find number]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ");"
                    DoCmd.RunSQL sql
                End If
                DoCmd.Minimize
                DoCmd.OpenForm stDocPostCranial, , , stLinkCriteria
            End If
        Else
            MsgBox "Please fill in the Element field for this record first", vbInformation, "No Element"
        End If
    Else
        MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
    End If

Exit_button_goto_cran_postcran_Click:
    Exit Sub

Err_button_goto_cran_postcran_Click:
    If Err.Number = 2046 And Me.Dirty = False Then
        'save record has failed  - DB might be Read only, added dirty check to make sure nothing added
        Resume Next
    Else
        Call General_Error_Trap
        Resume Exit_button_goto_cran_postcran_Click
    End If
End Sub

Private Sub Button23_Click()
' This event used to call the macro Bone.new basic record, translated to code
' sets the field [Forms]![Fauna_Bone_Basic_Faunal_Data]![Unit number] to
' [Forms]![Fauna_Bone_Faunal_Unit_Description]![Unit number] - I've extended this
' to trap poss the Unit desc form not open and to cope with it
' SAJ
On Error GoTo err_but23
Dim oldnum

oldnum = Me![Unit number]

DoCmd.RunCommand acCmdRecordsGoToNew
Me![Unit number] = Forms![Fauna_Bone_Faunal_Unit_Description]![Unit number]
DoCmd.GoToControl "Find number"
Exit Sub

getUnitNo:
    Dim retVal, checknum, sql, retVal2
    If oldnum <> "" Then
        retVal = MsgBox("Does the new record apply to Unit " & oldnum & "?", vbQuestion + vbYesNo, "New Record for Unit")
        If retVal = vbYes Then
            Me![Unit number] = oldnum
            Me![Unit number].Locked = False
            Me![Unit number].BackColor = 16777215
            Me![letter code].Locked = False
            Me![letter code].BackColor = 16777215
            Me![find number].Locked = False
            Me![find number].BackColor = 16777215
        Else
            retVal = InputBox("Please enter the Unit number below:", "Unit number")
            If retVal = "" Then
                MsgBox "New record entry cancelled", vbCritical, "No Unit Number Entered"
                DoCmd.RunCommand acCmdRecordsGoToLast
                Exit Sub
            Else
                'ok unit number entered by user, check if its in Unit descrip table
                checknum = DLookup("[Unit Number]", "[Fauna_Bone_Faunal_Unit_Description]", "[Unit number] = " & retVal)
                If IsNull(checknum) Then
                    retVal2 = MsgBox("The Unit Number " & retVal & " does not exist in the FUD, if you wish to continue with this entry you will be passed back to the FUD now." & Chr(13) & Chr(13) & "Do you want to continue with this entry?", vbExclamation + vbYesNo, "No Matching FUD")
                    If retVal2 = vbYes Then
                        'insert unit into FUD
                        sql = "INSERT INTO [Fauna_Bone_Faunal_Unit_Description] ([Unit Number]) VALUES (" & retVal & ");"
                        DoCmd.RunSQL sql
                        DoCmd.OpenForm "Fauna_Bone_Faunal_Unit_Description", acNormal, , "[Unit Number] = " & retVal
                        Exit Sub
                    Else
                        'cancel op
                        DoCmd.RunCommand acCmdRecordsGoToLast
                        Exit Sub
                    End If
                Else
                    'new 2009 - never did anything with entered value before!
                    Me![Unit number] = retVal
                    Me![Unit number].Locked = False
                    Me![Unit number].BackColor = 16777215
                    Me![letter code].Locked = False
                    Me![letter code].BackColor = 16777215
                    Me![find number].Locked = False
                    Me![find number].BackColor = 16777215
                End If
                'Me![Unit number] = retVal
                'DoCmd.GoToControl "Find number"
            End If
        End If
    End If
Exit Sub
err_but23:
    If Err.Number = 2450 Then
        GoTo getUnitNo
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ - filter remove request NR 5/7/06
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    
        If Me.Filter <> "" Then
            If Me.Filter <> "[GID] = '" & Me![cboFind].Column(1) & "'" Then
            '    MsgBox "This form was opened to only show the record " & Me.Filter & ". This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
                Me.FilterOn = False
                Me.Filter = ""
            End If
        End If

    DoCmd.GoToControl "GID"
    DoCmd.FindRecord Me![cboFind]
    Me!cboFind = ""

End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMenu_Click()
'new menu button
On Error GoTo err_cmdMenu

    DoCmd.OpenForm "Bone", acNormal
    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdMenu:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Comments_LostFocus()
'new 2010 - take focus up to post/cran button - focus movement still not quite wrking
'to teams satisfaction, this might help but can be taken out if not
On Error GoTo err_Comments

    DoCmd.GoToControl "button.goto.cran/postcran"
Exit Sub

err_Comments:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Field34_AfterUpdate()
On Error GoTo err_field34

'new season 2009 = for post ex field
PostEx_BFD_BodyPortion False, Me![Field34], Me![Field40]
PostEx_BFD_SizeClass False, Me![Field34]

Exit Sub

err_field34:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Field40_AfterUpdate()
'New code to replace call to macro Bone.on element update. The macro used to open the cranial
' / post cranial form and delete the record that way - no need to do this in code just run sql to delete
' SAJ
On Error GoTo err_field40
Dim sql, retVal
    If [Forms]![Fauna_Bone_Basic_Faunal_Data]![Element] < 24 And DCount("*", "Fauna_Bone_Postcranial", "[GID]=[Forms]![Fauna_Bone_Basic_Faunal_Data]![GID]") > 0 Then
        'if the element val < 24 and GID exist in table post cranial remove it as its now a cranial record
        retVal = MsgBox("A post cranial record for this GID already exists and this action will remove it as the element you have chosen is Cranial. Are you sure you want to continue?", vbQuestion + vbYesNo, "Change of Element")
        If retVal = vbNo Then
            Me![Field40] = Me![Field40].OldValue
            Exit Sub
        Else
            'if local run sql here
            sql = "DELETE FROM [Fauna_Bone_Postcranial] WHERE [GID] = '" & Me![GID] & "';"
            DoCmd.RunSQL sql
        End If
    ElseIf [Forms]![Fauna_Bone_Basic_Faunal_Data]![Element] > 23 And DCount("*", "Fauna_Bone_Cranial", "[GID]=[Forms]![Fauna_Bone_Basic_Faunal_Data]![GID]") > 0 Then
        'if element > 23 and GID exist in table cranial then remove it as its not post cranial
        retVal = MsgBox("A cranial record for this GID already exists and this action will remove it as the element you have chosen is Post-Cranial. Are you sure you want to continue?", vbQuestion + vbYesNo, "Change of Element")
        If retVal = vbNo Then
            Me![Field40] = Me![Field40].OldValue
            Exit Sub
        Else
            sql = "DELETE FROM [Fauna_Bone_Cranial] WHERE [GID] = '" & Me![GID] & "';"
            DoCmd.RunSQL sql
        End If
   End If

    'new season 2009 = for post ex field
    PostEx_BFD_BodyPortion False, Me![Field34], Me![Field40]

Exit Sub

err_field40:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Find_number_AfterUpdate()
On Error GoTo err_findnum

    'check existence of GID - new 2008 wishlist - saj
    If IsNull(Me![Unit number].OldValue) Then
        Dim checknum, GID, unit
        GID = [Forms]![Fauna_Bone_Basic_Faunal_Data]![Unit number] & "." & [Forms]![Fauna_Bone_Basic_Faunal_Data]![letter code] & [Forms]![Fauna_Bone_Basic_Faunal_Data]![find number]
        checknum = DLookup("[GID]", "[Fauna_Bone_Basic_Faunal_Data]", "[GID] = '" & GID & "'")
        If Not IsNull(checknum) Then
            'exists
            MsgBox "This gid number exists already, entry cancelled", vbInformation, "Duplicate GID"
            unit = Me![Unit number]
            Me.Undo
            
            Me![Unit number] = unit
            DoCmd.GoToControl "unit number"
            'why won't going to the find number field work? it just skips onto sample number??
            'DoCmd.GoToControl "find number"
            'Me.[find number].SetFocus
            Exit Sub
        End If
    End If
    
     'replaces call to bone.Update GID (used to be called OnChange also but this unecess)
    'SAJ
    Call UpdateGID 'this is private sub above

Exit Sub

err_findnum:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Form_Activate()
'season 2006 - to track movement from this screen set a local module var here
On Error GoTo err_act

   WhereGo = ""
   
Exit Sub

err_act:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Form_Current()
'new 2009 lock unit number field if not a new entry
On Error GoTo err_current

If Me![Unit number] <> "" And Me![letter code] <> "" And Me![find number] Then
    Me![Unit number].Locked = True
    'Me![Unit number].BackColor = 10079487 'kathy not like color same as background
    Me![Unit number].BackColor = 26367
    Me![letter code].Locked = True
    'Me![letter code].BackColor = 10079487
    Me![letter code].BackColor = 26367
    Me![find number].Locked = True
    'Me![find number].BackColor = 10079487
    Me![find number].BackColor = 26367
Else
    'this code is repeated in button23_click - the addnew button
    Me![Unit number].Locked = False
    Me![Unit number].BackColor = 16777215
    Me![letter code].Locked = False
    Me![letter code].BackColor = 16777215
    Me![find number].Locked = False
    Me![find number].BackColor = 16777215
End If
Exit Sub

err_current:
    If Err.Number = 2447 Then
        'for some reason when FUD not open and new record created
        'the if statement crashes on first line
        'message about invalid use of !. - can't get why - something to do with it
        'getting the value from the FUD on new record. So simply trapping
        Resume Next
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub Form_Deactivate()
'new season 2006 - request that if user closes without entering cran or post cran record
'there be some form of warning, this event used to cal macro: saverecord
On Error GoTo err_deact

    'only force save if can
   ' If Not IsNull(Me![Find Number]) And Not IsNull(Me![Unit Number]) And Not IsNull(Me![letter code]) Then
   '     DoCmd.RunCommand acCmdSaveRecord
   ' Else
   '     'DoCmd.RunCommand acCmdUndo
   ' End If
    
    If WhereGo <> "Post/Cran" Then
        'only do this check if not heading off to cran or post cran form
        'MsgBox "must check"
        'If Me![Field40] <> 1 Or Me![Field40] <> 117 Or Me![Field40] <> 118 Or Me![Field40] <> 119 Then - whoops wrong operator
        'If Me![Field40] <> 1 And Me![Field40] <> 117 And Me![Field40] <> 118 And Me![Field40] <> 119 Then
            'applies to all elements except 1, 117, 118, 119
        'change july 2007, saj - add 116
        If Me![Field40] <> 1 And Me![Field40] <> 117 And Me![Field40] <> 118 And Me![Field40] <> 119 And Me![Field40] <> 116 Then
            'applies to all elements except 1, 117, 118, 119, 116
            
            Dim numcheck, tablename, msgname
            If Me![Field40] < 24 Then
                tablename = "Fauna_Bone_Cranial"
                msgname = "Cranial"
            Else
                tablename = "Fauna_Bone_Postcranial"
                msgname = "Post Cranial"
            End If
            numcheck = DLookup("[GID]", tablename, "[GID] = '" & Me![GID] & "'")
            If IsNull(numcheck) Then
                'no cran/post cran msg so flag up
                MsgBox "A " & msgname & " record has not been entered for this GID. Please do not forget.", vbInformation, "Data Reminder"
            End If

        End If
    Else
        'MsgBox "no check"
    End If
Exit Sub

err_deact:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
'2008 wishlist, stop error message coming up
'this is not going to stop a message but trying to control it.
    'MsgBox DataErr
    'Response = acDataErrDisplay

    Const conInvalidKey = 3146
    Const conNullValue = 3162
    Dim strMsg As String
    Dim msgresp

    If DataErr = conInvalidKey Then
        response = acDataErrContinue
        strMsg = "Missing unit number, letter code or sample number. Do wish to cancel this entry?"
        msgresp = MsgBox(strMsg, vbExclamation + vbYesNo, "Primary Key Violation")
        If msgresp = vbYes Then
            Me.Undo
        End If
    ElseIf DataErr = conNullValue Then
        'invalid use of null might be unit number blank so undo record
        response = acDataErrContinue
        strMsg = "There is a value missing preventing this record from being saved. Do wish to cancel this entry?"
        msgresp = MsgBox(strMsg, vbExclamation + vbYesNo, "Primary Key Violation")
        If msgresp = vbYes Then
            Me.Undo
        End If
        
    End If


End Sub

Private Sub Form_GotFocus()
'this was a macro call to Bone.Update BFD Unit. It traps the scenario that the form is opened
' and the unit number is 0 but it relies on Unit Desc form being open to gather the unit number
'from there, if opened from the main menu Unit Desc will not be opened so this is trapped
' SAJ
On Error GoTo err_frmfocus

    If Me![Unit number] = 0 Then
        Me![Unit number] = Forms![Fauna_Bone_Faunal_Unit_Description]![Unit number]
    
    End If

Exit Sub

err_frmfocus:
    If Err.Number = 2450 Then
        'form not open so ignore this action
        Exit Sub
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub



Private Sub Form_LostFocus()
'new season 2006 - requesst that if user closes
'saverecord
'MsgBox "here"
End Sub



Private Sub Form_Open(Cancel As Integer)
'new 2009 - show/hide post ex fields
On Error GoTo err_open

If GetGeneralPermissions = "Admin" Then 'rissa request change from visible only to admin 21july09
    Me![txtBodyPortion].Locked = False
    Me![txtSizeClass].Locked = False
    Me![txtBodyPortion].BackColor = 16777215
    Me![txtSizeClass].BackColor = 16777215
Else
   Me![txtBodyPortion].Locked = True
   Me![txtBodyPortion].BackColor = 26367
    Me![txtSizeClass].Locked = True
    Me![txtSizeClass].BackColor = 26367
End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Letter_code_AfterUpdate()
'replaces call to bone.Update GID (used to be called OnChange also but this unecess)
'SAJ
Call UpdateGID 'this is private sub above
End Sub

Private Sub OpenDZInstructions_Click()
On Error GoTo Err_OpenDZInstructions_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "DZ_instructions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OpenDZInstructions_Click:
    Exit Sub

Err_OpenDZInstructions_Click:
    MsgBox Err.Description
    Resume Exit_OpenDZInstructions_Click
    
End Sub
Private Sub Command86_Click()
On Error GoTo Err_Command86_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "DZ_instructions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command86_Click:
    Exit Sub

Err_Command86_Click:
    MsgBox Err.Description
    Resume Exit_Command86_Click
    
End Sub

Private Sub Unit_number_AfterUpdate()
'replaces call to bone.Update GID (used to be called onEnter and OnChange also but this unecess)
'SAJ
Call UpdateGID 'this is private sub above
End Sub
