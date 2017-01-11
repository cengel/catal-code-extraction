Option Compare Database   'Use database order for string comparisons
Option Explicit
'******************************************************************************
' this sub menu is new with v9.2 - it offers all the admin functions
' SAJ v9.2
'******************************************************************************



Private Sub cmdBuildings_Click()
On Error GoTo Err_cmdBuildings_Click
  
    DoCmd.OpenForm "Exca: Admin_Buildings", acNormal

    Exit Sub

Err_cmdBuildings_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCleaning_Click()
On Error GoTo Err_cmdCleaning_Click
  
    DoCmd.OpenForm "Excavation:AdminCleaningMenu", acNormal

    Exit Sub

Err_cmdCleaning_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdClose_Click()
'********************************************
' This will close this form and return to main menu
' SAJ v9.2
'********************************************
On Error GoTo err_cmdClose_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdClose_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdFeature_Click()
'Features
On Error GoTo Err_cmdFeatures_Click

    
    DoCmd.OpenForm "Exca: Admin_Features", acNormal


    Exit Sub

Err_cmdFeatures_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdFeatureType_Click()
On Error GoTo Err_cmdFeatureType_Click
  
    DoCmd.OpenForm "Exca: Admin_FeatureTypeSubTypeLOV", acNormal

    Exit Sub

Err_cmdFeatureType_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdFT_Click()
On Error GoTo err_cmdFT_Click

    
    DoCmd.OpenForm "Exca: Admin_Foundation_Trenches", acNormal


    Exit Sub

err_cmdFT_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdHPhase_Click()
On Error GoTo err_cmdHPhase_Click

    DoCmd.OpenForm "Exca: Admin_HodderPhaseLOV", acNormal
    Exit Sub

err_cmdHPhase_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdLevel_Click()
On Error GoTo Err_cmdLevel_Click

    
    DoCmd.OpenForm "Exca: Admin_LevelLOV", acNormal


    Exit Sub

Err_cmdLevel_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdProblem_Click()
On Error GoTo Err_cmdLevel_Click

    
    DoCmd.OpenForm "Exca: Admin_ProblemReports", acNormal

    Exit Sub

Err_cmdLevel_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSampleTypes_Click()
On Error GoTo Err_cmdSampleTypes_Click

    
    DoCmd.OpenForm "Exca: Admin_SampleTypesLOV", acNormal


    Exit Sub

Err_cmdSampleTypes_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdSpace_Click()
On Error GoTo Err_cmdSpaces_Click

    
    DoCmd.OpenForm "Exca: Admin_Spaces", acNormal


    Exit Sub

Err_cmdSpaces_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnits_Click()
On Error GoTo Err_cmdUnits_Click

    
    DoCmd.OpenForm "Exca: Admin_Units", acNormal


    Exit Sub

Err_cmdUnits_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnits2_Click()
cmdUnits_Click
End Sub

Private Sub Command18_Click()
cmdLevel_Click
End Sub

Private Sub Command24_Click()
cmdClose_Click
End Sub


Private Sub Command25_Click()
cmdFeatureType_Click
End Sub

Private Sub Command29_Click()
cmdFeature_Click
End Sub

Private Sub Command31_Click()
cmdSpace_Click
End Sub

Private Sub Command33_Click()
cmdBuildings_Click
End Sub

Private Sub Command35_Click()
cmdFT_Click
End Sub

Private Sub Command37_Click()
cmdCleaning_Click
End Sub

Private Sub Command39_Click()
cmdProblem_Click
End Sub

Private Sub Command42_Click()
cmdHPhase_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
