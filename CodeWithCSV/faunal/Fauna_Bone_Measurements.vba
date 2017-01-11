Option Compare Database   'Use database order for string comparisons
Option Explicit 'saj

Private Sub Button23_Click()
'altered season 2006 - saj
'error trap and new way of creating new rec
Dim tempGID
Dim tempUnit
Dim tempLetter
Dim tempFind

Dim newSet, checkRec
newSet = InputBox("Please enter the new measurement set for GID " & Me![GID] & " below:", "New Measurement Set")
If newSet <> "" Then
    tempGID = GID
    tempUnit = Unit_number
    tempLetter = Letter_code
    tempFind = Find_number

    'new 2008 wishlist check this measurement set not exist alread
    checkRec = DLookup("[GID]", "Fauna_Bone_Measurements", "[GID] = '" & tempGID & "' AND [Measurement Set] = " & newSet)
    If IsNull(checkRec) Then
        'this GID and measurement set doesn't exist so allow - new 2008
        
        'from here to else is old 2006 code
        'DoCmd.DoMenuItem MenuBar:=acFormBar, MenuName:=3, Command:=0
        DoCmd.RunCommand acCmdRecordsGoToNew

        GID = tempGID
        Unit_number = tempUnit
        Letter_code = tempLetter
        Find_number = tempFind
        Me![Measurement set] = newSet

    Else
        'this measurement set exists for the GID so stop creation - new 2008
        MsgBox "This Measurement Set already exists for this GID. Please use the find list to locate it.", vbInformation, "Record Already Exists"
        DoCmd.GoToControl "cboFind"
    End If
End If

Exit Sub

err_23:
    Call General_Error_Trap
    Exit Sub
End Sub


Sub button_goto_previousform_Click()
On Error GoTo Err_button_goto_previousform_Click

    Dim stDocCranial As String
    Dim stDocPostCranial As String
    Dim stLinkCriteria As String
    Dim stElementType
    
    stDocCranial = "Fauna_Bone_Cranial"
    stDocPostCranial = "Fauna_Bone_Postcranial"
    'SAJ season 2006 - this depends on the basic form being open so now the
    'recordsource of this form is the modification table with the basic table
    'joined to get the element value
    ''stElementType = Forms![Fauna_Bone_Basic_Faunal_Data]![Field40]
    stElementType = Me![Element]
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
   If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
        If stElementType < 24 Then

            DoCmd.Minimize
            DoCmd.OpenForm stDocCranial, , , stLinkCriteria

        Else
        
            DoCmd.Minimize
            DoCmd.OpenForm stDocPostCranial, , , stLinkCriteria
        End If
    Else
        MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No GID Number"
    End If

Exit_button_goto_previousform_Click:
    Exit Sub

Err_button_goto_previousform_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_previousform_Click
    
End Sub
Sub button_open_modification_Click()
'altered season 2006 - saj
'check if modification record exists and if not ask user to create
On Error GoTo Err_button_open_modification_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, sql, retVal
    stDocName = "Fauna_Bone_Modification"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
 If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    
    'new for season 2006 - see if the modification record exists if not create it - SAJ
    checknum = DLookup("[GID]", "[Fauna_Bone_Modification]", "[GID] = '" & Me![GID] & "'")
    If IsNull(checknum) Then
        'gid not exist there yet
        retVal = MsgBox("No Modification record exists yet for GID " & Me![GID] & ", do you want to create one now?", vbQuestion + vbYesNo, "Create New Modification Record")
        If retVal = vbNo Then
            'do nothing
            Exit Sub
        Else
            'so create it ready for data entry
            sql = "INSERT INTO [Fauna_Bone_Modification] ([GID], [Unit number], [Letter code], [Find number]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ");"
            DoCmd.RunSQL sql
        End If
    End If
       
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_button_open_modification_Click:
    Exit Sub

Err_button_open_modification_Click:
    Call General_Error_Trap
    Resume Exit_button_open_modification_Click
    
End Sub
Sub button_goto_artefacts_Click()
'altered season 2006 - saj
'check if artifact record exists and if not ask user to create
On Error GoTo Err_button_goto_artefacts_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, sql, retVal
    stDocName = "Fauna_Bone_Artifacts"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
           
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    
    'new for season 2006 - see if the modification record exists if not create it - SAJ
    checknum = DLookup("[GID]", "[Fauna_Bone_Artifacts]", "[GID] = '" & Me![GID] & "'")
    If IsNull(checknum) Then
        'gid not exist there yet
        retVal = MsgBox("No Artifact record exists yet for GID " & Me![GID] & ", do you want to create one now?", vbQuestion + vbYesNo, "Create New Modification Record")
        If retVal = vbNo Then
            'do nothing
            Exit Sub
        Else
            'so create it ready for data entry
            sql = "INSERT INTO [Fauna_Bone_Artifacts] ([GID], [Unit number], [Letter code], [Find number]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ");"
            DoCmd.RunSQL sql
        End If
    End If

               
    DoCmd.Minimize
    
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
    'shouldn't these be after open? - yes moved bug found by rebecca
    [Forms]![Fauna_Bone_Artifacts].[Unit number] = [Forms]![Fauna_Bone_Measurements].[Unit number]
    [Forms]![Fauna_Bone_Artifacts].[letter code] = [Forms]![Fauna_Bone_Measurements].[letter code]
    [Forms]![Fauna_Bone_Artifacts].[find number] = [Forms]![Fauna_Bone_Measurements].[find number]
    'if we are going to do this should do GID as well - added by SAJ
    [Forms]![Fauna_Bone_Artifacts].[GID] = [Forms]![Fauna_Bone_Measurements].[GID]
    
    
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If
Exit_button_goto_artefacts_Click:
    Exit Sub

Err_button_goto_artefacts_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_artefacts_Click
    
End Sub

Sub button_goto_gid_Click()
On Error GoTo Err_button_goto_gid_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
        
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_goto_gid_Click:
    Exit Sub

Err_button_goto_gid_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_gid_Click
    
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ - slightly different here as works as a filter to go directly to
'GID and measurement set
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    'If Me.Filter <> "" Then
    '        If Me.Filter <> "[GID] = '" & Me![cboFind].Column(1) & "'" Then
    '            MsgBox "This form was opened to only show the record " & Me.Filter & ". This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
    '            Me.FilterOn = False
    '            Me.Filter = ""
    '        End If
    '    End If
    Me.Filter = "[GID] = '" & Me![cboFind] & "' AND [Measurement Set] = " & Me![cboFind].Column(1)
    Me.FilterOn = True
    'Me![lblFilter].Caption = "Showing GID " & Me![cboFind] & " set " & Me![cboFind].Column(1)
    'Me![lblFilter].Visible = True
    
    'DoCmd.GoToControl "GID"
    'DoCmd.FindRecord Me![cboFind]

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

Private Sub Form_Current()
'new go to control command
On Error GoTo err_current

    DoCmd.GoToControl "Field101"
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Deactivate()
'2008 wishlist - Rissa has lost records when moving back to the BFD and then
'the server blipping. The existing saverecord is a macro placed on lost focus
'but having save here should hopefully capture more
On Error GoTo err_formdeact
    DoCmd.RunCommand acCmdSaveRecord
    

Exit Sub

err_formdeact:
    Call General_Error_Trap
    Exit Sub
End Sub
