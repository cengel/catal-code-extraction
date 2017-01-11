Option Compare Database   'Use database order for string comparisons

Private Sub Area_Sheet_Click()
On Error GoTo Err_Area_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Finds: Basic Data"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Area_Sheet_Click:
    Exit Sub

Err_Area_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Area_Sheet_Click
End Sub


Private Sub Building_Sheet_Click()
On Error GoTo Err_Building_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Store: Crate Register"
    stLinkCriteria = ""
    
    'Added where clause to select for subset of crates only depending on user 2013 season, amended 2014
    If CrateLetterFlag = "P" Or CrateLetterFlag = "HB" Or CrateLetterFlag = "OB" Or CrateLetterFlag = "PH" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter = '" & CrateLetterFlag & "'"
    ElseIf CrateLetterFlag = "CO" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter In ('" & CrateLetterFlag & "', 'BE', 'FG', 'CB')"
    ElseIf CrateLetterFlag = "GS" Then
         stLinkCriteria = "[Store: Crate Register].CrateLetter In ('" & CrateLetterFlag & "', 'NS', 'Depot')"
    ElseIf CrateLetterFlag = "FB" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter In ('" & CrateLetterFlag & "', 'Depot')"
    ElseIf CrateLetterFlag = "BE" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter = '" & CrateLetterFlag & "'"
    ElseIf CrateLetterFlag = "char" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter In ('" & CrateLetterFlag & "', 'or')"
    ElseIf CrateLetterFlag = "S" Then
        stLinkCriteria = "[Store: Crate Register].CrateLetter In ('" & CrateLetterFlag & "', 'BE')"
    End If
    ' otherwise load the whole empanada
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    

Exit_Building_Sheet_Click:
    Exit Sub

Err_Building_Sheet_Click:
    MsgBox Err.Description
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

Private Sub Command18_Click()
Open_priority_Click
End Sub

Private Sub Feature_Sheet_Button_Click()
On Error GoTo Err_Feature_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Feature_Sheet_Button_Click:
    Exit Sub

Err_Feature_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Feature_Sheet_Button_Click
End Sub


Sub Open_priority_Click()
On Error GoTo Err_Open_priority_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Priority Detail"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Open_priority_Click:
    Exit Sub

Err_Open_priority_Click:
    MsgBox Err.Description
    Resume Exit_Open_priority_Click
    
End Sub



Private Sub Command22_Click()

Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Finds: Unstrat"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command22_Click:
    Exit Sub

Err_Command22_Click:
    MsgBox Err.Description
    Resume Exit_Command22_Click
End Sub

Private Sub Return_to_Master_Con_Click()

DoCmd.DoMenuItem acFormBar, acFileMenu, 14, , acMenuVer70

End Sub

Private Sub cmdAdmin_Click()
'new season 2006 saj
On Error GoTo err_cmdAdmin

If GetGeneralPermissions = "Admin" Then
    DoCmd.OpenForm "Finds: Admin_MaterialGroupSubGroupLOV"
Else
    MsgBox "Only adminstrators can access this page", vbExclamation, "Administrators Only"
End If
Exit Sub

err_cmdAdmin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCrateCodes_Click()
'just a quick view of the crate codes
On Error GoTo err_codes

    DoCmd.OpenForm "frm_pop_cratecodes", acNormal, , , acFormReadOnly, acDialog
    
Exit Sub

err_codes:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCrateDescr_Click()
'new season 2008 will show count of distinct description entries in crate reg
On Error GoTo err_cratemat
    DoCmd.OpenQuery "Julie_Count_Description_in_Crate_Reg", acViewNormal, acReadOnly
Exit Sub

err_cratemat:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCrateMaterials_Click()
'new season 2008 will show count of distinct material entries in crate reg
On Error GoTo err_cratemat
    DoCmd.OpenQuery "Julie_Count_Materials_in_Crate_Reg", acViewNormal, acReadOnly
Exit Sub

err_cratemat:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDBLink_Click()
On Error GoTo err_dblink

    DoCmd.OpenForm "Finds: Database Link", acNormal
Exit Sub

err_dblink:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdLetters_Click()
'just a quick view of the letter codes v3.1
On Error GoTo err_letters

    DoCmd.OpenForm "frm_pop_letter_prefixes", acNormal, , , acFormReadOnly, acDialog
    
Exit Sub

err_letters:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMaterials_Click()
'just a quick view of the material groups v3.1
On Error GoTo err_letters

    DoCmd.OpenForm "frm_pop_Materials_with_subgroups", acNormal, , , acFormReadOnly, acDialog
    
Exit Sub

err_letters:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'new season 2006 saj
On Error GoTo err_open

If GetGeneralPermissions = "Admin" Then
    Me![cmdAdmin].Visible = True
Else
    Me![cmdAdmin].Visible = False
End If
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Space_Sheet_Button_Click()
On Error GoTo Err_Space_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Space_Sheet_Button_Click:
    Exit Sub

Err_Space_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Space_Sheet_Button_Click

End Sub


Private Sub Unit_Sheet_Click()
On Error GoTo Err_Unit_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord acForm, stDocName, acLast

Exit_Unit_Sheet_Click:
    Exit Sub

Err_Unit_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Unit_Sheet_Click

End Sub


Sub Command19_Click()
On Error GoTo Err_Command19_Click


    DoCmd.Close

Exit_Command19_Click:
    Exit Sub

Err_Command19_Click:
    MsgBox Err.Description
    Resume Exit_Command19_Click
    
End Sub
Sub Finds_Print_Click()
On Error GoTo Err_Finds_Print_Click

    Dim stDocName As String

    stDocName = "Finds: Sheets printout"
    DoCmd.OpenReport stDocName, acPreview

Exit_Finds_Print_Click:
    Exit Sub

Err_Finds_Print_Click:
    MsgBox Err.Description
    Resume Exit_Finds_Print_Click
    
End Sub
Sub Print_Cratereg_Click()
On Error GoTo Err_Print_Cratereg_Click

    Dim stDocName As String

    stDocName = "Finds Store: Crate Register"
    DoCmd.OpenReport stDocName, acPreview

Exit_Print_Cratereg_Click:
    Exit Sub

Err_Print_Cratereg_Click:
    MsgBox Err.Description
    Resume Exit_Print_Cratereg_Click
    
End Sub
Sub Print_unstrat_Click()
On Error GoTo Err_Print_unstrat_Click

    Dim stDocName As String

    stDocName = "Finds: Unstrat Printout"
    DoCmd.OpenReport stDocName, acPreview

Exit_Print_unstrat_Click:
    Exit Sub

Err_Print_unstrat_Click:
    MsgBox Err.Description
    Resume Exit_Print_unstrat_Click
    
End Sub
Sub Print_conserv_Click()
On Error GoTo Err_Print_conserv_Click

    Dim stDocName As String

    stDocName = "Conserv: Full Printout"
    DoCmd.OpenReport stDocName, acPreview

Exit_Print_conserv_Click:
    Exit Sub

Err_Print_conserv_Click:
    MsgBox Err.Description
    Resume Exit_Print_conserv_Click
    
End Sub
Sub Command28_Click()
On Error GoTo Err_Command28_Click


    DoCmd.Close

Exit_Command28_Click:
    Exit Sub

Err_Command28_Click:
    MsgBox Err.Description
    Resume Exit_Command28_Click
    
End Sub

Private Sub X_Finds_Sheet_Click()

End Sub
Private Sub OpenUnitSheet_Click()
On Error GoTo Err_OpenUnitSheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OpenUnitSheet_Click:
    Exit Sub

Err_OpenUnitSheet_Click:
    MsgBox Err.Description
    Resume Exit_OpenUnitSheet_Click
    
End Sub
