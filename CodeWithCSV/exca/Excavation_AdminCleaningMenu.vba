Option Compare Database   'Use database order for string comparisons
Option Explicit
'******************************************************************************
' this sub menu is new in 2008
' SAJ
'******************************************************************************



Private Sub cmdBuildings_Click()
On Error GoTo err_Four
    MsgBox "Generating data..."
    Call CheckFeatureSpaceBuildingRelationships
    DoCmd.OpenReport "LocalCheckFeatureSpaceBuildingRels", acViewPreview
    

Exit Sub

err_Four:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_cmdClose_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation:AdminMenu"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Excavation:ListsMenu"
Exit Sub

err_cmdClose_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDistinctFeatures_Click()
On Error GoTo err_Distinct
    DoCmd.OpenReport "R_Cleaning_Distinct_FeatureTypes", acViewPreview

Exit Sub
err_Distinct:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdFeatureList_Click()
On Error GoTo err_feature
    DoCmd.OpenReport "R_Features_and _SubTypes", acViewPreview

Exit Sub
err_feature:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMissing_Click()
On Error GoTo err_cmdMissing

    DoCmd.OpenForm "Exca: Admin_Subform_MissingNumbers", acNormal, , , , acDialog
    
Exit Sub

err_cmdMissing:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOne_Click()
On Error GoTo err_One
    
    Call CheckFeatureSpaceUnitSpaceRelationships
    DoCmd.OpenReport "LocalCheckFeatureSpaceUnitSpaceRels", acViewPreview
    

Exit Sub

err_One:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSpace_Click()
On Error GoTo err_Three
    MsgBox "Generating data..."
    Call CheckUnitFeatureBuildingRelationships
    DoCmd.OpenReport "LocalCheckUnitFeatureBuildingRels", acViewPreview
    

Exit Sub

err_Three:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTwo_Click()
On Error GoTo err_Two
    MsgBox "Generating data..."
    Call CheckUnitSpaceBuildingRelationships
    DoCmd.OpenReport "LocalCheckUnitSpaceBuildingRels", acViewPreview
    

Exit Sub

err_Two:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command35_Click()
On Error GoTo err_QOne

    DoCmd.OpenQuery "Q_Cleaning_Units_where_Building_NoSpace", acViewNormal, acReadOnly

Exit Sub

err_QOne:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command36_Click()
On Error GoTo err_QTwo

    DoCmd.OpenQuery "Q_Cleaning_Units_where_Feature_NoSpace", acViewNormal, acReadOnly

Exit Sub

err_QTwo:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command37_Click()
On Error GoTo err_QThree

    DoCmd.OpenQuery "Q_Cleaning_Features_where_NoBuilding_NoSpace", acViewNormal, acReadOnly

Exit Sub

err_QThree:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command38_Click()
On Error GoTo err_QFour

    DoCmd.OpenQuery "Q_Cleaning_Features_where_Building_NoSpace", acViewNormal, acReadOnly

Exit Sub

err_QFour:
    Call General_Error_Trap
    Exit Sub
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
