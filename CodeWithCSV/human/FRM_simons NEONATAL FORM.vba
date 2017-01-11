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

Private Sub CmdOpenDeciduousTeethFrm_Click()
On Error GoTo Err_CmdOpenDeciduousTeethFrm_Click

    Call DoRecordCheck("HR_Teeth development measurement", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth development score", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth wear", Me![txtUnit], Me![txtIndivid], "UnitNumber")
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons DECIDUOUS TEETH"
    'saj season 2007
    stDocName = "FRM_Deciduous_Teeth"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenDeciduousTeethFrm_Click:
    Exit Sub

Err_CmdOpenDeciduousTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenDeciduousTeethFrm_Click
    
End Sub
Private Sub CmdMeasFrm_Click()
On Error GoTo Err_CmdMeasFrm_Click

    Call DoRecordCheck("HR_Measurements version 2", Me![txtUnit], Me![txtIndivid], "UnitNumber")

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    stDocName = "FRM_Measurement form version 2"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name

Exit_CmdMeasFrm_Click:
    Exit Sub

Err_CmdMeasFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdMeasFrm_Click
    
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
Private Sub CmdOpenMenuFrm_Click()
Call ReturnToMenu(Me)

    
End Sub
Private Sub Deciduous_teeth_Click()
On Error GoTo Err_Deciduous_teeth_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

Exit_Deciduous_teeth_Click:
    Exit Sub

Err_Deciduous_teeth_Click:
    MsgBox Err.Description
    Resume Exit_Deciduous_teeth_Click
    
End Sub
Private Sub openfrmdecid_neonatalform_Click()
On Error GoTo Err_openfrmdecid_neonatalform_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_simons DECIDUOUS TEETH"
    
    stLinkCriteria = "[Individual number]=" & Me![List3]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_openfrmdecid_neonatalform_Click:
    Exit Sub

Err_openfrmdecid_neonatalform_Click:
    MsgBox Err.Description
    Resume Exit_openfrmdecid_neonatalform_Click
    
End Sub
Private Sub DeciduousTeeth_Click()
On Error GoTo Err_DeciduousTeeth_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_simons DECIDUOUS TEETH"
    
    stLinkCriteria = "[Individual number]=" & Me![txtIndivid]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_DeciduousTeeth_Click:
    Exit Sub

Err_DeciduousTeeth_Click:
    MsgBox Err.Description
    Resume Exit_DeciduousTeeth_Click
    
End Sub
