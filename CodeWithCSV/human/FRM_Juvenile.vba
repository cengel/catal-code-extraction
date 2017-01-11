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

Private Sub CmdOpenDecidTeethFrm_Click()
On Error GoTo Err_CmdOpenDecidTeethFrm_Click

    Call DoRecordCheck("HR_Teeth development measurement", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth development score", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth wear", Me![txtUnit], Me![txtIndivid], "UnitNumber")
  
    Dim stDocName As String
    Dim stLinkCriteria As String

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

    'stDocName = "FRM_simons PERMANENT TEETH"
    stDocName = "FRM_Permanent_Teeth"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenPermTeethFrm_Click:
    Exit Sub

Err_CmdOpenPermTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenPermTeethFrm_Click
    
End Sub
Private Sub CmdOpenMeasFrm_Click()
On Error GoTo Err_CmdOpenMeasFrm_Click

    Call DoRecordCheck("HR_Measurements version 2", Me![txtUnit], Me![txtIndivid], "UnitNumber")

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Measurement form version 2"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenMeasFrm_Click:
    Exit Sub

Err_CmdOpenMeasFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMeasFrm_Click
    
End Sub
Private Sub CmdOpenUnitDescFrm_Click()
On Error GoTo Err_CmdOpenUnitDescFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenUnitDescFrm_Click:
    Exit Sub

Err_CmdOpenUnitDescFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenUnitDescFrm_Click
    
End Sub
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click

    Call DoRecordCheck("HR_ageing and sexing", Me![txtUnit], Me![txtIndivid], "Unit Number")

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , "[Unit Number] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub

Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
    
End Sub
Private Sub CmdOpenMainMenuFrm_Click()
On Error GoTo Err_CmdOpenMainMenuFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_MAINMENU"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenMainMenuFrm_Click:
    Exit Sub

Err_CmdOpenMainMenuFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMainMenuFrm_Click
    
End Sub

Private Sub Command462_Click()
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

Private Sub Command463_Click()
Call ReturnToMenu(Me)

End Sub

Private Sub Form_Open(Cancel As Integer)
'saj new season 2007 - form jumping to middle on open - make sure it keeps the top at the top
On Error GoTo err_open

    DoCmd.GoToControl "FRM_SUBFORM_Juvenile_skull"
    'docmd.GoToControl ""

Exit Sub

err_open:
    General_Error_Trap
    Exit Sub
End Sub
