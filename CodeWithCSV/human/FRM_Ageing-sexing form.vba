Option Compare Database
Option Explicit

Private Sub cboAgeCategory_AfterUpdate()
'must warn user of consequences of change
On Error GoTo err_cboAgeCategory

Dim msg, retVal

    If Me![cboAgeCategory].OldValue <> "" Then
        msg = "When the system is fully developed this change will be checked to see what implications it might have if data has"
        msg = msg & " already been entered into the Neonate, Juvenile or Adult form." & Chr(13) & Chr(13) & "No check exists at present"
        msg = msg & " and it is up to you to tidy up any existing data" & Chr(13) & Chr(13) & "Continue with this change?"
        retVal = MsgBox(msg, vbYesNo, "Development Point")
        If retVal = vbNo Then
            Me![cboAgeCategory] = Me![cboAgeCategory].OldValue
        End If
        
    End If
    Call SortOutButtons(Me)
Exit Sub

err_cboAgeCategory:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub cboFind_AfterUpdate()
'find skeleton record - SAJ
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
        Me.Filter = "[Unit Number] = " & Me![cboFind] & " AND [Individual Number] = " & Me!cboFind.Column(1)
        Me.FilterOn = True
    End If
Exit Sub

err_cboFind:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdAll_Click()
'take off any filter - saj
On Error GoTo err_all
    '2009
    'Me.FilterOn = False
    'Me.Filter = ""
    '2009 - combo might only be showing for one unit - reset
    Me!cboFind.RowSource = "SELECT [HR_ageing and sexing].[unit number], [HR_ageing and sexing].[Individual number] FROM [HR_ageing and sexing] ORDER BY [HR_ageing and sexing].[unit number], [HR_ageing and sexing].[Individual number]; "
Exit Sub

err_all:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdGuide_Click()
'new season 2010
On Error GoTo err_cmdGuide

    DoCmd.OpenForm "frm_pop_agesexguide", acNormal, , , acFormReadOnly

Exit Sub

err_cmdGuide:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdMenu_Click()
Call ReturnToMenu(Me)

End Sub

Private Sub CmdOpenJuvenileFrm_Click()
On Error GoTo Err_CmdOpenJuvenileFrm_Click

    Call DoRecordCheck("HR_Juvenile_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_axial", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons juvenile form"
    ''season 2007, saj
    Me.Requery
    stDocName = "FRM_Juvenile"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
    
Exit_CmdOpenJuvenileFrm_Click:
    Exit Sub

Err_CmdOpenJuvenileFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenJuvenileFrm_Click
    
End Sub
Private Sub CmdOpenAdultFrm_Click()
On Error GoTo Err_CmdOpenAdultFrm_Click

    Call DoRecordCheck("HR_Adult_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Axial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons adult form"
    ''season 2007, saj
    Me.Requery
    stDocName = "FRM_Adult"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
    
Exit_CmdOpenAdultFrm_Click:
    Exit Sub

Err_CmdOpenAdultFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAdultFrm_Click
    
End Sub
Private Sub CmdOpenUnitDescFrm_Click()
On Error GoTo Err_CmdOpenUnitDescFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenUnitDescFrm_Click:
    Exit Sub

Err_CmdOpenUnitDescFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenUnitDescFrm_Click
    
End Sub
Private Sub CmdOpenMainMenuFrm_Click()

Call ReturnToMenu(Me)

    
End Sub
Private Sub CmdOpenDecidTeethFrm_Click()
On Error GoTo Err_CmdOpenDecidTeethFrm_Click

    Call DoRecordCheck("HR_Teeth development measurement", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth development score", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth wear", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons DECIDUOUS TEETH"
    ''season 2007, saj
    Me.Requery
    stDocName = "FRM_Deciduous_Teeth"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenDecidTeethFrm_Click:
    Exit Sub

Err_CmdOpenDecidTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenDecidTeethFrm_Click
    
End Sub
Private Sub CmdOpenPermTeethFrm_Click()
On Error GoTo Err_CmdOpenPermTeethFrm_Click

    Call DoRecordCheck("HR_Teeth development measurement", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth development score", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth wear", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons PERMANENT TEETH"
    ''season 2007, saj
    Me.Requery
    stDocName = "FRM_Permanent_Teeth"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenPermTeethFrm_Click:
    Exit Sub

Err_CmdOpenPermTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenPermTeethFrm_Click
    
End Sub
Private Sub CmdOpenNeonateFrm_Click()
On Error GoTo Err_CmdOpenNeonateFrm_Click

    Call DoRecordCheck("HR_Neonate_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Neonate_arm_leg_data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Neonate_Axial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String
    Me.Requery

    stDocName = "FRM_simons NEONATAL FORM"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenNeonateFrm_Click:
    Exit Sub

Err_CmdOpenNeonateFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenNeonateFrm_Click
    
End Sub

Private Sub Form_Current()
'for each skeleton check the age category to determin which form buttons to offer
On Error GoTo err_current
Call SortOutButtons(Me)
'    If Me![cboAgeCategory] <> "" Then
'
'        If Me![cboAgeCategory] = 0 Then
'            Me![CmdOpenNeonateFrm].Enabled = True
'            Me![CmdOpenJuvenileFrm].Enabled = False
'            Me![CmdOpenAdultFrm].Enabled = False
'        ElseIf Me![cboAgeCategory] = 1 Or Me![cboAgeCategory] = 2 Or Me![cboAgeCategory] = 3 Or Me![cboAgeCategory] = 4 Then
'            Me![CmdOpenNeonateFrm].Enabled = False
'            Me![CmdOpenJuvenileFrm].Enabled = True
'            Me![CmdOpenAdultFrm].Enabled = False
'        ElseIf Me![cboAgeCategory] = 5 Or Me![cboAgeCategory] = 6 Or Me![cboAgeCategory] = 7 Then
'            Me![CmdOpenNeonateFrm].Enabled = False
'            Me![CmdOpenJuvenileFrm].Enabled = False
'            Me![CmdOpenAdultFrm].Enabled = True
'        Else
'            Me![CmdOpenNeonateFrm].Enabled = True
'            Me![CmdOpenJuvenileFrm].Enabled = True
'            Me![CmdOpenAdultFrm].Enabled = True
'        End If
'
'   Else
'        Me![CmdOpenNeonateFrm].Enabled = False
'            Me![CmdOpenJuvenileFrm].Enabled = False
'            Me![CmdOpenAdultFrm].Enabled = False
'    End If
Exit Sub

err_current:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Delete(Cancel As Integer)
'new may 2009, now cascade delete setup this is a dangerous operation - warn user
On Error GoTo err_delete

    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Dim response
        response = MsgBox("Deleting this skeleton will mean permanent deletion of all data associated with this particular skeleton in this database." & Chr(13) & Chr(13) & "Do you really want to delete " & Me![txtUnit] & ".B" & Me![txtIndivid] & "?", vbCritical + vbYesNo, "Critical Delete")
        If response = vbNo Then
            Cancel = True
        Else
            Cancel = False
        End If
    Else
        MsgBox "You do not have permission to delete this record, please contact your team leader"
        Cancel = True
    End If
    

Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub
