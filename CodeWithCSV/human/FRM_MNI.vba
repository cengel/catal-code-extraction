Option Compare Database
Option Explicit
Private which

Private Sub cmdRecalc_Click()
'Wow - 1st Sept 2009 - SAJ - this is good!
On Error GoTo err_cmdRecalc

If which <> "" Then
    Dim response
    response = MsgBox("This process will re-calculate all the MNI's shown below and you will loose any comments." & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Re-Calculation")
    If response = vbYes Then
    
        Me![FRM_SUBFORM_MNI].Visible = False
        Me![txtFeature].Visible = True
        Me![txtMsg].Visible = True
        
        If which = 1 Then 'feature
            CalcMNI
        ElseIf which = 2 Then 'space
            CalcSpaceMNI
        ElseIf which = 3 Then 'building
            CalcBuildingMNI
        End If
        Me![FRM_SUBFORM_MNI].Requery
        Me![FRM_SUBFORM_MNI].Visible = True
        Me![txtMsg] = "Re-calculation complete"
    End If
Else
    MsgBox "The form does not know which MNI to calculate and cannot proceed", vbCritical, "Error"
End If
Exit Sub

err_cmdRecalc:
    If Err.Number = 3156 Then
        MsgBox "Sorry but you are logged in as HumanRemains and cannot perform this function as it includes a delete. You must be logged in as HumanRemainsLeader.", vbExclamation, "Permission Denied"
        Me![FRM_SUBFORM_MNI].Visible = True
    
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)

On Error GoTo err_open
which = ""
'cut in here in 2010 with different options for which MNI
If IsNull(Me.OpenArgs) Then
    MsgBox "Form opened without selecting which MNI, operation cancelled", vbExclamation, "Invalid Call"
    DoCmd.Close acForm, Me.Name
Else
    'openargs specify which MNI
    which = Me.OpenArgs
    
    If which = 1 Then
        Me![lblTitle].Caption = "Calculate MNI for Burial Features"
        Me![FRM_SUBFORM_MNI].SourceObject = "FRM_SUBFORM_MNI"
        'Q_MNI_LastGenerated
        Me![FRM_SUBFORM_MNI_LastGenerated].Form.RecordSource = "Q_MNI_LastGenerated"
    ElseIf which = 2 Then
        Me![lblTitle].Caption = "Calculate MNI for Spaces with skeleton units"
        Me![FRM_SUBFORM_MNI].SourceObject = "FRM_SUBFORM_MNI_SPACE"
        Me![FRM_SUBFORM_MNI_LastGenerated].Form.RecordSource = "Q_MNI_LastGenerated_Space"
    ElseIf which = 3 Then
        Me![lblTitle].Caption = "Calculate MNI for Buildings with skeleton units"
        Me![FRM_SUBFORM_MNI].SourceObject = "FRM_SUBFORM_MNI_BUILDING"
        Me![FRM_SUBFORM_MNI_LastGenerated].Form.RecordSource = "Q_MNI_LastGenerated_Building"
    End If
    
    'added jan 2010 when catalhoyuk allowed access to this bit of the database
    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Me![cmdRecalc].Enabled = True
    Else
        Me![cmdRecalc].Enabled = False
    End If

    Me![txtMsg].Visible = False
    Me![txtFeature].Visible = False
    Me![FRM_SUBFORM_MNI].Visible = True
End If
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub
