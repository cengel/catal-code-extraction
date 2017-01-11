Option Compare Database   'Use database order for string comparisons
Option Explicit 'saj

Sub button_goto_unitdescription_Click()
On Error GoTo Err_button_goto_unitdescription_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Faunal_Unit_Description"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    
If Me![Unit number] <> "" Then
    'new check for Unit entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
        MsgBox "Please enter select a Unit first", vbInformation, "No Unit Number"
End If

Exit_button_goto_unitdescription_Click:
    Exit Sub

Err_button_goto_unitdescription_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_unitdescription_Click
    
End Sub
Sub button_goto_cran_postcran_Click()
On Error GoTo Err_button_goto_cran_postcran_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    
If Me![Unit number] <> "" Then
    'new check for Unit entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
            
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
    'If [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = 0 Then 'there is a 0 unit in this system
    If IsNull([Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number]) Then
    [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = [Forms]![Bone: Short Faunal Data].[Unit number]
    End If
Else
        MsgBox "Please enter select a Unit first", vbInformation, "No Unit Number"
End If

Exit_button_goto_cran_postcran_Click:
    Exit Sub

Err_button_goto_cran_postcran_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_cran_postcran_Click
    
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    If Me.Filter <> "" Then
            If Me.Filter <> "[Unit] = " & Me![cboFind] Then
                MsgBox "This form was opened to only show the record " & Me.Filter & ". This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
                Me.FilterOn = False
                Me.Filter = ""
            End If
        End If
    DoCmd.GoToControl "Unit Number"
    DoCmd.FindRecord Me![cboFind]

End If

Exit Sub

err_cboFind:
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
