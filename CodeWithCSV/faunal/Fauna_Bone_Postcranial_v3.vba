Option Compare Database   'Use database order for string comparisons
Option Explicit 'added by saj

Sub Button_Goto_BFD_Click()
On Error GoTo Err_Button_Goto_BFD_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
If Me![GID] <> "" Then
    'new check for GID entered by saj
    DoCmd.Close
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_Button_Goto_BFD_Click:
    Exit Sub

Err_Button_Goto_BFD_Click:
    MsgBox Err.Description
    Resume Exit_Button_Goto_BFD_Click

End Sub
Sub button_goto_measurement_Click()
'altered season 2006 - saj
'check if any measurement records exist and if not ask user to create

On Error GoTo Err_button_goto_measurement_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim retVal, checknum, sql, getMeasurementSet
    stDocName = "Fauna_Bone_Measurements"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
        
    'new for season 2006 - see if the modification record exists if not create it - SAJ
    checknum = DLookup("[GID]", "[Fauna_Bone_Measurements]", "[GID] = '" & Me![GID] & "'")
    If IsNull(checknum) Then
        'gid not exist there yet
        retVal = MsgBox("No Measurement records exist yet for GID " & Me![GID] & ", do you want to create one now?", vbQuestion + vbYesNo, "Create New Modification Record")
        If retVal = vbNo Then
            'do nothing
            Exit Sub
        Else
            getMeasurementSet = InputBox("Please enter the measurement set number below:", "Measurement Set")
            If getMeasurementSet <> "" Then
                'so create it ready for data entry
                sql = "INSERT INTO [Fauna_Bone_Measurements] ([GID], [Unit number], [Letter code], [Find number], [Measurement Set]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ", " & getMeasurementSet & ");"
                DoCmd.RunSQL sql
            End If
        End If
    End If
        
        
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_button_goto_measurement_Click:
    Exit Sub

Err_button_goto_measurement_Click:
    If Err.Number = 2046 And Me.Dirty = False Then
        'save record has failed  - DB might be Read only, added dirty check to make sure nothing added
        Resume Next
    Else
        Call General_Error_Trap
        Resume Exit_button_goto_measurement_Click
    End If
End Sub
Sub button_goto_modification_Click()
'altered season 2006 - saj
'check if modification record exists and if not ask user to create
On Error GoTo Err_button_goto_modification_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim retVal, checknum, sql
    
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

Exit_button_goto_modification_Click:
    Exit Sub

Err_button_goto_modification_Click:
    If Err.Number = 2046 And Me.Dirty = False Then
        'save record has failed  - DB might be Read only, added dirty check to make sure nothing added
        Resume Next
    Else
        Call General_Error_Trap
        Resume Exit_button_goto_modification_Click
    End If
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
    
    [Forms]![Fauna_Bone_Artifacts].[Unit number] = [Forms]![Fauna_Bone_Postcranial].[Unit number]
    [Forms]![Fauna_Bone_Artifacts].[letter code] = [Forms]![Fauna_Bone_Postcranial].[letter code]
    [Forms]![Fauna_Bone_Artifacts].[find number] = [Forms]![Fauna_Bone_Postcranial].[find number]
    'if we are going to do this should do GID as well - added by SAJ
    [Forms]![Fauna_Bone_Artifacts].[GID] = [Forms]![Fauna_Bone_Postcranial].[GID]

Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_button_goto_artefacts_Click:
    Exit Sub

Err_button_goto_artefacts_Click:
    If Err.Number = 2046 And Me.Dirty = False Then
        'save record has failed  - DB might be Read only, added dirty check to make sure nothing added
        Resume Next
    Else
        Call General_Error_Trap
        Resume Exit_button_goto_artefacts_Click
    End If
End Sub
Sub button_goto_unitBFD_Click()
On Error GoTo Err_button_goto_unitBFD_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_goto_unitBFD_Click:
    Exit Sub

Err_button_goto_unitBFD_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_unitBFD_Click
    
End Sub
Sub button_goto_bfdgid_Click()
On Error GoTo Err_button_goto_bfdgid_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
        
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_goto_bfdgid_Click:
    Exit Sub

Err_button_goto_bfdgid_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_bfdgid_Click
    
End Sub
Sub cmd_gotobfdunit_Click()
On Error GoTo Err_cmd_gotobfdunit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmd_gotobfdunit_Click:
    Exit Sub

Err_cmd_gotobfdunit_Click:
    MsgBox Err.Description
    Resume Exit_cmd_gotobfdunit_Click
    
End Sub

Private Sub button_gotobfdunit_Click()

End Sub


Private Sub cboFind_AfterUpdate()
'new find combo by SAJ - filter msg remved request from NR 5/7/06
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
'this used to call Bone.Update PostCran Unit Letter and Find Number
'which did the following:
On Error GoTo err_curr

'If [Forms]![Fauna_Bone_PostCranial]![Unit number] = 0 Then 'this was ok in MF but not in here as there is a unit 0!
If IsNull([Forms]![Fauna_Bone_Postcranial]![Unit number]) Then
    [Forms]![Fauna_Bone_Postcranial]![Unit number] = [Forms]![Fauna_Bone_Basic_Faunal_Data]![Unit number]
    [Forms]![Fauna_Bone_Postcranial]![letter code] = [Forms]![Fauna_Bone_Basic_Faunal_Data]![letter code]
    [Forms]![Fauna_Bone_Postcranial]![find number] = [Forms]![Fauna_Bone_Basic_Faunal_Data]![find number]
    [Forms]![Fauna_Bone_Postcranial]![GID] = [Forms]![Fauna_Bone_Basic_Faunal_Data]![GID]
End If
DoCmd.GoToControl "Field23"
Exit Sub

err_curr:
    If Err.Number = 2450 Then
        'form not open so ignore this action
        Exit Sub
    Else
        Call General_Error_Trap
    End If
Exit Sub
End Sub


