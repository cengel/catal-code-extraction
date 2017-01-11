Option Compare Database   'Use database order for string comparisons
Option Explicit 'saj

Private Sub button_goto_contact_Click()
On Error GoTo Err_button_goto_contact_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Contact"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
        
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
    If IsNull([Forms]![Fauna_Bone_Contact].[GID]) Then
        'no existing record
        [Forms]![Fauna_Bone_Contact].[Unit number] = [Forms]![Fauna_Bone_Artifacts].[Unit number]
        [Forms]![Fauna_Bone_Contact].[letter code] = [Forms]![Fauna_Bone_Artifacts].[letter code]
        [Forms]![Fauna_Bone_Contact].[find number] = [Forms]![Fauna_Bone_Artifacts].[find number]
        [Forms]![Fauna_Bone_Contact].[GID] = [Forms]![Fauna_Bone_Artifacts].[GID]
        [Forms]![Fauna_Bone_Contact].[ContactOrder] = 1
    End If
Exit_button_goto_contact_Click:
    Exit Sub

Err_button_goto_contact_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_contact_Click
End Sub


Private Sub button_goto_previousform_Click()
On Error GoTo Err_button_goto_previousform_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    ' MR July 18, 2005
    ' Assume that application logic doesn't require that GID is carried back
    ' as it will be already in the form from the way in
    'SAJ season 2006 can get to this form from main menu so need GID carried back
    'to BFD to cover circumstance its not open already, commented out if
    ''If [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = 0 Then
        stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    ''End If
        
    If Me![GID] <> "" Then
        'new check for GID entered by saj
        'the form is only minimised so must save data manually here - saj
        DoCmd.RunCommand acCmdSaveRecord
        
        DoCmd.Minimize

        DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No GID Number"
    End If
    
Exit_button_goto_previousform_Click:
    Exit Sub

Err_button_goto_previousform_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_previousform_Click
    
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
'new goto control command here
On Error GoTo err_current
DoCmd.GoToControl "Field102"

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub
