Option Compare Database   'Use database order for string comparisons
Option Explicit 'SAJ

Sub button_goto_previousform_Click()
On Error GoTo Err_button_goto_previousform_Click

    Dim stDocCranial As String
    Dim stDocPostCranial As String
    Dim stLinkCriteria As String
    Dim stElementType 'new by saj as opt explicit
    
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





Sub button_goto_bfdgid_Click()
On Error GoTo Err_button_goto_bfdgid_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"

If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_button_goto_bfdgid_Click:
    Exit Sub

Err_button_goto_bfdgid_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_bfdgid_Click
    
End Sub
Sub button_goto_artefact_Click()
'altered season 2006 - saj
'check if artifact record exists and if not ask user to create
On Error GoTo Err_button_goto_artefact_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, sql, retVal
    
    stDocName = "Fauna_Bone_Artifacts"
    
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
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
    
    [Forms]![Fauna_Bone_Artifacts].[Unit number] = [Forms]![Fauna_Bone_Modification].[Unit number]
    [Forms]![Fauna_Bone_Artifacts].[letter code] = [Forms]![Fauna_Bone_Modification].[letter code]
    [Forms]![Fauna_Bone_Artifacts].[find number] = [Forms]![Fauna_Bone_Modification].[find number]
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No GID Number"
End If

Exit_button_goto_artefact_Click:
    Exit Sub

Err_button_goto_artefact_Click:
   Call General_Error_Trap
    Resume Exit_button_goto_artefact_Click
    
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ - filter msg removed request from NR 5/7/06
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
'new go to control code
On Error GoTo err_current
    'causing error, changed. Faunal Wishlist Aug 2008
    'DoCmd.GoToControl "Field101"
    'DoCmd.GoToControl "cboFind"
    'request from claire 18/07/09 please change to first entry field
    DoCmd.GoToControl "Field41"
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub
