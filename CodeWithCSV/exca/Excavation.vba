Option Compare Database   'Use database order for string comparisons

Private Sub Area_Sheet_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Area_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Area Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Area_Sheet_Click:
    Exit Sub

Err_Area_Sheet_Click:
    Call General_Error_Trap
    Resume Exit_Area_Sheet_Click
End Sub


Private Sub Building_Sheet_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Building_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Building Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Building_Sheet_Click:
    Exit Sub

Err_Building_Sheet_Click:
    Call General_Error_Trap
    Resume Exit_Building_Sheet_Click
End Sub

Private Sub Button10_Click()
Building_Sheet_Click
End Sub

Private Sub Button11_Click()
Space_Sheet_Button_Click
End Sub


Private Sub Button12_Click()
Feature_Sheet_Button_Click
End Sub


Private Sub Button13_Click()
Unit_Sheet_Click
End Sub


Private Sub Button17_Click()
Area_Sheet_Click
End Sub

Private Sub Button9_Click()
Return_to_Master_Con_Click
End Sub

Private Sub cmdAdmin_Click()
'************************************************************
' New with v9.1 - admininstration menu
'
' SAJ v9.1
'************************************************************
On Error GoTo err_cmdAdmin_Click

    DoCmd.OpenForm "Excavation:AdminMenu", acNormal, , , acFormReadOnly
    
Exit Sub

err_cmdAdmin_Click:
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

Private Sub cmdLists_Click()
'************************************************************
' New with v9.1 - two lists buttons orig on this form moved
' to new predefined lists submenu
'
' SAJ v9.1
'************************************************************
On Error GoTo err_cmdLists_Click

    DoCmd.OpenForm "Excavation:ListsMenu", acNormal, , , acFormReadOnly
    
Exit Sub

err_cmdLists_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command18_Click()
Open_priority_Click
End Sub

Private Sub Command25_Click()
cmdLists_Click
End Sub

Private Sub Command27_Click()
cmdAdmin_Click
End Sub

Private Sub Command29_Click()
cmdFT_Click
End Sub

Private Sub Feature_Sheet_Button_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Feature_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Feature_Sheet_Button_Click:
    Exit Sub

Err_Feature_Sheet_Button_Click:
    Call General_Error_Trap
    Resume Exit_Feature_Sheet_Button_Click
End Sub


Private Sub FeatureStatus_Click()
On Error GoTo Err_FeatureStatus_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FeaturesheetStatus"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
Exit_FeatureStatus_Click:
    Exit Sub

Err_FeatureStatus_Click:
    Call General_Error_Trap
    Resume Exit_FeatureStatus_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'DoCmd.RunCommand acCmdSizeToFit
If GetGeneralPermissions = "Admin" Then
    Me![cmdAdmin].Enabled = True
    Me![Command27].Enabled = True
Else
    Me![cmdAdmin].Enabled = False
    Me![Command27].Enabled = False
End If
End Sub

Sub Open_priority_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Open_priority_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Priority Detail"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Open_priority_Click:
    Exit Sub

Err_Open_priority_Click:
    Call General_Error_Trap
    Resume Exit_Open_priority_Click
    
End Sub



Private Sub Return_to_Master_Con_Click()
'************************************************************
' Removed link to macro: Exca: Macros. This quits out of Access
'
' SAJ v9.1
'************************************************************
'legacy code already here
'DoCmd.DoMenuItem acFormBar, acFileMenu, 14, , acMenuVer70
On Error GoTo err_Return_to_Master_Con_Click

    DoCmd.Quit acQuitSaveAll
    
Exit Sub

err_Return_to_Master_Con_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Space_Sheet_Button_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Space_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Space Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Space_Sheet_Button_Click:
    Exit Sub

Err_Space_Sheet_Button_Click:
    Call General_Error_Trap
    Resume Exit_Space_Sheet_Button_Click

End Sub


Private Sub Unit_Sheet_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Unit_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
   '' REMOVED SAJ AS WILL FAIL FOR RO USERS DoCmd.GoToRecord acForm, stDocName, acNewRec

Exit_Unit_Sheet_Click:
    Exit Sub

Err_Unit_Sheet_Click:
    Call General_Error_Trap
    Resume Exit_Unit_Sheet_Click

End Sub


Sub Feature_types_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_Feature_types_Click

    Dim stDocName As String

    stDocName = "Exca: Feature Types"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Feature_types_Click:
    Exit Sub

Err_Feature_types_Click:
    Call General_Error_Trap
    Resume Exit_Feature_types_Click
    
End Sub
Sub interpret_Click()
'orig code - just added general error trap - SAJ v9.1
On Error GoTo Err_interpret_Click

    Dim stDocName As String

    stDocName = "Exca: List Interpretive Categories"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_interpret_Click:
    Exit Sub

Err_interpret_Click:
    Call General_Error_Trap
    Resume Exit_interpret_Click
    
End Sub

Private Sub UnitStatus_Click()
On Error GoTo Err_UnitStatus_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "UnitsheetStatus"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
   '' REMOVED SAJ AS WILL FAIL FOR RO USERS DoCmd.GoToRecord acForm, stDocName, acNewRec

Exit_UnitStatus_Click:
    Exit Sub

Err_UnitStatus_Click:
    Call General_Error_Trap
    Resume Exit_UnitStatus_Click
End Sub


