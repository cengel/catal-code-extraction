Option Compare Database

Private Sub cboFind_AfterUpdate()
'find skeleton record - SAJ
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
        Me.Filter = "[UnitNumber] = " & Me![cboFind] & " AND [Individual Number] = " & Me!cboFind.Column(1)
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

    Me.FilterOn = False
    Me.Filter = ""
Exit Sub

err_all:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdMenu_Click()
Call ReturnToMenu(Me)
End Sub

Private Sub CmdOpenJuvenileFrm_Click()
On Error GoTo Err_CmdOpenJuvFrm_Click

    Call DoRecordCheck("HR_Juvenile_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_axial", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons juvenile form"
    ''saj season 2007
    stDocName = "FRM_Juvenile"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
    
Exit_CmdOpenJuvFrm_Click:
    Exit Sub

Err_CmdOpenJuvFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenJuvFrm_Click
    
End Sub

Private Sub CmdOpenNeonateFrm_Click()
On Error GoTo Err_CmdOpenNeonateFrm_Click
    
    Call DoRecordCheck("HR_Neonate_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Neonate_arm_leg_data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Neonate_Axial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_simons NEONATAL FORM"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenNeonateFrm_Click:
    Exit Sub

Err_CmdOpenNeonateFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenNeonateFrm_Click
    
End Sub
Private Sub CmdOpenJuvFrm_Click()
On Error GoTo Err_CmdOpenJuvFrm_Click

    Call DoRecordCheck("HR_Juvenile_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_axial", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons juvenile form"
    ''saj season 2007
    stDocName = "FRM_Juvenile"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
    
Exit_CmdOpenJuvFrm_Click:
    Exit Sub

Err_CmdOpenJuvFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenJuvFrm_Click
    
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
    ''saj 2007
    stDocName = "FRM_Adult"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

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

Private Sub Form_Current()
'new season 2007 - hide and show buttons depending on age category
On Error GoTo err_current

    'the age should be brought into this form by the invisible field cboAgeCategory
    'that uses a function (GetSkeletonAge) to obtain it.
    'the following function then uses this value to set the adult/juvenile/neonatal buttons up
    Call SortOutButtons(Me)


Exit Sub

err_current:
    General_Error_Trap
    Exit Sub
End Sub
