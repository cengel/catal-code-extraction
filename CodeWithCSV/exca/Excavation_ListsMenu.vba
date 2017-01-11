Option Compare Database   'Use database order for string comparisons
Option Explicit
'******************************************************************************
' this sub menu is new with v9.1, it takes two buttons that were on the
' the original main menu: list all feature types, list interpertative categories
' SAJ v9.1
'******************************************************************************

Private Sub Button13_Click()
interpret_Click
End Sub

Private Sub cmdBuildingReport_Click()
'new for 2008
On Error GoTo err_cmdBuilding

    Dim resp, both
    resp = InputBox("If you wish to only report on a certain building please enter the number below, otherwise leave All for all buildings.", "Specify Building?", "All")
    
    'bug fix 2010 - wasnt capturing cancel so have put in default of All
    If resp <> "" Then
        both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        If resp <> "All" Then
            DoCmd.OpenReport "R_BuildingSheet", acViewPreview, , "[Number] = " & resp
            If both = vbYes Then DoCmd.OpenReport "R_Units_in_Buildings", acViewPreview, , "[In_Building] = " & resp
        Else
            DoCmd.OpenReport "R_BuildingSheet", acViewPreview
            If both = vbYes Then DoCmd.OpenReport "R_Units_in_Buildings", acViewPreview
        End If
    End If

Exit Sub

err_cmdBuilding:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
'********************************************
' This will close this form and return to main menu
' SAJ v9.1
'********************************************
On Error GoTo err_cmdClose_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Excavation:ListsMenu"
Exit Sub

err_cmdClose_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdFeatureReport_Click()
'new for 2009
On Error GoTo err_cmdFeature

    Dim resp, both
    resp = InputBox("To avoid over printing you can only print one feature at a time. Please enter the feature number below.", "Specify Feature")
    If resp <> "" Then
        DoCmd.OpenReport "R_Feature_Sheet", acViewPreview, , "[Feature Number] = " & resp
    End If

Exit Sub

err_cmdFeature:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSearchUnits_Click()
'open the new unit search screen 2009
On Error GoTo err_units

    DoCmd.OpenForm "frm_search", acNormal

Exit Sub

err_units:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSpaceSheet_Click()
'new for 2008
On Error GoTo err_cmdSpace

    Dim resp, both
    resp = InputBox("If you wish to only report on a certain space please enter the number below, otherwise leave All in place for all spaces.", "Specify Space?", "All")
    
    'bug fix 2010 - wasnt capturing cancel so have put in default of All
    If resp <> "" Then
        both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        If resp <> "All" Then
            DoCmd.OpenReport "R_SpaceSheet", acViewPreview, , "[Space Number] = " & resp
            If both = vbYes Then DoCmd.OpenReport "R_Units_in_Spaces", acViewPreview, , "[In_Space] = " & resp
        Else
            DoCmd.OpenReport "R_SpaceSheet", acViewPreview
            If both = vbYes Then DoCmd.OpenReport "R_Units_in_Spaces", acViewPreview
        End If
    End If

Exit Sub

err_cmdSpace:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnitReport_Click()
'new for 2009
On Error GoTo err_cmdUnit

    Dim resp, both
    resp = InputBox("To avoid over printing you can only print one unit at a time. Please enter the unit number below.", "Specify Unit")
    If resp <> "" Then
        'find out type of unit
        Dim unitcat
        If IsNumeric(resp) Then
            unitcat = DLookup("[Category]", "[Exca: Unit Sheet]", "[Unit number] = " & resp)
            If Not IsNull(unitcat) Then
                Select Case LCase(unitcat)
                Case "cut"
                    DoCmd.OpenReport "R_Unit_Sheet_cut", acViewPreview, , "[Unit Number] = " & resp
                Case "skeleton"
                    DoCmd.OpenReport "R_Unit_Sheet_skeleton", acViewPreview, , "[Unit Number] = " & resp
                Case Else
                    DoCmd.OpenReport "R_Unit_Sheet_layercluster", acViewPreview, , "[Unit Number] = " & resp
                End Select
            Else
                MsgBox "Unit number not present in the database.", vbInformation, "Data not found"
            End If
        Else
            MsgBox "Not a valid unit number", vbInformation, "Invalid entry"
        End If
    End If

Exit Sub

err_cmdUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command18_Click()
Feature_types_Click
End Sub

Private Sub Command24_Click()
cmdClose_Click
End Sub

Private Sub Command27_Click()
cmdSpaceSheet_Click
End Sub

Private Sub Command29_Click()
cmdSearchUnits_Click
End Sub

Private Sub Command34_Click()
cmdFeatureReport_Click
End Sub

Sub Feature_types_Click()
' orig apart from general error trap
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
' orig apart from general error trap
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
