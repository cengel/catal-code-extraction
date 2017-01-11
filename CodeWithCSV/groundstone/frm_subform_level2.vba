Option Compare Database

Private Sub cboBaseDiameterCompleteness_Change()
On Error GoTo err_changeBaseDiameterCompleteness
Dim strText As String

strText = Nz(Me.cboBaseDiameterCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboBaseDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboBaseDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboBaseDiameterCompleteness.Dropdown

Exit Sub

err_changeBaseDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboBaseDiameterCompleteness_Enter()
On Error GoTo err_EnterBaseDiameterCompleteness
    Me.cboBaseDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboBaseDiameterCompleteness.Dropdown
Exit Sub

err_EnterBaseDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboBaseDiameterCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownBaseDiameterCompleteness
    Me.cboBaseDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboBaseDiameterCompleteness.Dropdown
Exit Sub

err_KeyDownBaseDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboBaseType_Change()
On Error GoTo err_changeBaseType
Dim strText As String

strText = Nz(Me.cboBaseType.Text, "")

If Len(strText) > 0 Then

   Me.cboBaseType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Basetype] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboBaseType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Basetype] ORDER BY ScreenOrder"
End If

Me.cboBaseType.Dropdown

Exit Sub

err_changeBaseType:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cboBaseType_Enter()
On Error GoTo err_EnterBaseType
    Me.cboBaseType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Basetype] ORDER BY ScreenOrder"
    Me.cboBaseType.Dropdown
Exit Sub

err_EnterBaseType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboBaseType_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownBaseType
    Me.cboBaseType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Basetype] ORDER BY ScreenOrder"
    Me.cboBaseType.Dropdown
Exit Sub

err_KeyDownBaseType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboBitDamage_Change()
On Error GoTo err_changeBitDamage
Dim strText As String

strText = Nz(Me.cboBitDamage.Text, "")

If Len(strText) > 0 Then

   Me.cboBitDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboBitDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
End If

Me.cboBitDamage.Dropdown

Exit Sub

err_changeBitDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboBitDamage_Enter()
On Error GoTo err_EnterBitDamage
    Me.cboBitDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
    Me.cboBitDamage.Dropdown
Exit Sub

err_EnterBitDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboBitDamage_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownBitDamage
    Me.cboBitDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
    Me.cboBitDamage.Dropdown
Exit Sub

err_KeyDownBitDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtDamage_Change()
On Error GoTo err_changeButtDamage
Dim strText As String

strText = Nz(Me.cboButtDamage.Text, "")

If Len(strText) > 0 Then

   Me.cboButtDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboButtDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
End If

Me.cboButtDamage.Dropdown

Exit Sub

err_changeButtDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtDamage_Enter()
On Error GoTo err_EnterButtDamage
    Me.cboButtDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
    Me.cboButtDamage.Dropdown
Exit Sub

err_EnterButtDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtDamage_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownButtDamage
    Me.cboButtDamage.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Damage] ORDER BY ScreenOrder"
    Me.cboButtDamage.Dropdown
Exit Sub

err_KeyDownButtDamage:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtThicknessCompleteness_Change()
On Error GoTo err_changeButtThicknessCompleteness
Dim strText As String

strText = Nz(Me.cboButtThicknessCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboButtThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboButtThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboButtThicknessCompleteness.Dropdown

Exit Sub

err_changeButtThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboButtThicknessCompleteness_Enter()
On Error GoTo err_EnterButtThicknessCompleteness
    Me.cboButtThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboButtThicknessCompleteness.Dropdown
Exit Sub

err_EnterButtThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtThicknessCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownButtThicknessCompleteness
    Me.cboButtThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboButtThicknessCompleteness.Dropdown
Exit Sub

err_KeyDownButtThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboButtWidthCompleteness_Change()
On Error GoTo err_changeButtWidthCompleteness
Dim strText As String

strText = Nz(Me.cboButtWidthCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboButtWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboButtWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboButtWidthCompleteness.Dropdown

Exit Sub

err_changeButtWidthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtWidthCompleteness_Enter()
On Error GoTo err_EnterButtWidthCompleteness
    Me.cboButtWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboButtWidthCompleteness.Dropdown
Exit Sub

err_EnterButtWidthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboButtWidthCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownButtWidthCompleteness
    Me.cboButtWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboButtWidthCompleteness.Dropdown
Exit Sub

err_KeyDownButtWidthCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboCortWeath_Change()
On Error GoTo err_changeCortWeath
Dim strText As String

strText = Nz(Me.cboCortWeath.Text, "")

If Len(strText) > 0 Then

   Me.cboCortWeath.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Weathered Surface] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboCortWeath.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Weathered Surface] ORDER BY ScreenOrder"
End If

Me.cboCortWeath.Dropdown

Exit Sub

err_changeCortWeath:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboCortWeath_Enter()
On Error GoTo err_EnterCortWeath
    Me.cboCortWeath.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Weathered Surface] ORDER BY ScreenOrder"
    Me.cboCortWeath.Dropdown
Exit Sub

err_EnterCortWeath:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboCortWeath_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownCortWeath
    Me.cboCortWeath.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Weathered Surface] ORDER BY ScreenOrder"
    Me.cboCortWeath.Dropdown
Exit Sub

err_KeyDownCortWeath:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDegreeBurning_Change()
On Error GoTo err_changeDegreeBurning
Dim strText As String

strText = Nz(Me.cboDegreeBurning.Text, "")

If Len(strText) > 0 Then

   Me.cboDegreeBurning.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree of Burning] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboDegreeBurning.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree of Burning] ORDER BY ScreenOrder"
End If

Me.cboDegreeBurning.Dropdown

Exit Sub

err_changeDegreeBurning:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboDegreeBurning_Enter()
On Error GoTo err_EnterDegreeBurning
    Me.cboDegreeBurning.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree of Burning] ORDER BY ScreenOrder"
    Me.cboDegreeBurning.Dropdown
Exit Sub

err_EnterDegreeBurning:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDegreeBurning_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownDegreeBurning
    Me.cboDegreeBurning.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree of Burning] ORDER BY ScreenOrder"
    Me.cboDegreeBurning.Dropdown
Exit Sub

err_KeyDownDegreeBurning:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDegreePolish_Change()
On Error GoTo err_changeDegreePolish
Dim strText As String

strText = Nz(Me.cboDegreePolish.Text, "")

If Len(strText) > 0 Then

   Me.cboDegreePolish.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Polish] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboDegreePolish.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Polish] ORDER BY ScreenOrder"
End If

Me.cboDegreePolish.Dropdown

Exit Sub

err_changeDegreePolish:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDegreePolish_Enter()

On Error GoTo err_EnterDegreePolish
    Me.cboDegreePolish.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Polish] ORDER BY ScreenOrder"
    Me.cboDegreePolish.Dropdown
Exit Sub

err_EnterDegreePolish:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDegreePolish_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownDegreePolish
    Me.cboDegreePolish.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Polish] ORDER BY ScreenOrder"
    Me.cboDegreePolish.Dropdown
Exit Sub

err_KeyDownDegreePolish:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFracturePattern_Change()
On Error GoTo err_changeFracturePattern
Dim strText As String

strText = Nz(Me.cboFracturePattern.Text, "")

If Len(strText) > 0 Then

   Me.cboFracturePattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fracture Pattern] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboFracturePattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fracture Pattern] ORDER BY ScreenOrder"
End If

Me.cboFracturePattern.Dropdown

Exit Sub

err_changeFracturePattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFracturePattern_Enter()
On Error GoTo err_EnterFracturePattern
    Me.cboFracturePattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fracture Pattern] ORDER BY ScreenOrder"
    Me.cboFracturePattern.Dropdown
Exit Sub

err_EnterFracturePattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFracturePattern_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownFracturePattern
    Me.cboFracturePattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fracture Pattern] ORDER BY ScreenOrder"
    Me.cboFracturePattern.Dropdown
Exit Sub

err_KeyDownFracturePattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentation_Change()
On Error GoTo err_changeFragmentation
Dim strText As String

strText = Nz(Me.cboFragmentation.Text, "")

If Len(strText) > 0 Then

   Me.cboFragmentation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboFragmentation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation] ORDER BY ScreenOrder"
End If

Me.cboFragmentation.Dropdown

Exit Sub

err_changeFragmentation:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentation_Enter()

On Error GoTo err_EnterFragmentation
    Me.cboFragmentation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation] ORDER BY ScreenOrder"
    Me.cboFragmentation.Dropdown
Exit Sub

err_EnterFragmentation:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentation_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownFragmentation
    Me.cboFragmentation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation] ORDER BY ScreenOrder"
    Me.cboFragmentation.Dropdown
Exit Sub

err_KeyDownFragmentation:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentationPattern_Change()
On Error GoTo err_changeFragmentationPattern
Dim strText As String

strText = Nz(Me.cboFragmentationPattern.Text, "")

If Len(strText) > 0 Then

   Me.cboFragmentationPattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation Pattern] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboFragmentationPattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation Pattern] ORDER BY ScreenOrder"
End If

Me.cboFragmentationPattern.Dropdown

Exit Sub

err_changeFragmentationPattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentationPattern_Enter()

On Error GoTo err_EnterFragmentationPattern
    Me.cboFragmentationPattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation Pattern] ORDER BY ScreenOrder"
    Me.cboFragmentationPattern.Dropdown
Exit Sub

err_EnterFragmentationPattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFragmentationPattern_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownFragmentationPattern
    Me.cboFragmentationPattern.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Fragmentation Pattern] ORDER BY ScreenOrder"
    Me.cboFragmentationPattern.Dropdown
Exit Sub

err_KeyDownFragmentationPattern:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboGeologicalCategory_Change()
On Error GoTo err_changeGeologicalCategory
Dim strText As String

strText = Nz(Me.cboGeologicalCategory.Text, "")

If Len(strText) > 0 Then

   Me.cboGeologicalCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Geological Category] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboGeologicalCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Geological Category] ORDER BY ScreenOrder"
End If

Me.cboGeologicalCategory.Dropdown

Exit Sub

err_changeGeologicalCategory:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboGeologicalCategory_Enter()

On Error GoTo err_EnterGeologicalCategory
    Me.cboGeologicalCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Geological Category] ORDER BY ScreenOrder"
    Me.cboGeologicalCategory.Dropdown
Exit Sub

err_EnterGeologicalCategory:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboGeologicalCategory_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownGeologicalCategory
    Me.cboGeologicalCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Geological Category] ORDER BY ScreenOrder"
    Me.cboGeologicalCategory.Dropdown
Exit Sub

err_KeyDownGeologicalCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboGrainSize_Change()
On Error GoTo err_changeGrainSize
Dim strText As String

strText = Nz(Me.cboGrainSize.Text, "")

If Len(strText) > 0 Then

   Me.cboGrainSize.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Grain Size] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboGrainSize.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Grain Size] ORDER BY ScreenOrder"
End If

Me.cboGrainSize.Dropdown

Exit Sub

err_changeGrainSize:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboGrainSize_Enter()
On Error GoTo err_EnterGrainSize
    Me.cboGrainSize.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Grain Size] ORDER BY ScreenOrder"
    Me.cboGrainSize.Dropdown
Exit Sub

err_EnterGrainSize:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboGrainSize_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownGrainSize
    Me.cboGrainSize.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Grain Size] ORDER BY ScreenOrder"
    Me.cboGrainSize.Dropdown
Exit Sub

err_KeyDownGrainSize:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompleteness_Change()
On Error GoTo err_changeLengthCompleteness
Dim strText As String

strText = Nz(Me.cboLengthCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboLengthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboLengthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboLengthCompleteness.Dropdown

Exit Sub

err_changeLengthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompleteness_Enter()
On Error GoTo err_EnterLengthCompleteness
    Me.cboLengthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompleteness.Dropdown
Exit Sub

err_EnterLengthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownLengthCompleteness
    Me.cboLengthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompleteness.Dropdown
Exit Sub

err_KeyDownLengthCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboLengthCompletenessUseFace1_Change()
On Error GoTo err_changeLengthCompletenessUseFace1
Dim strText As String

strText = Nz(Me.cboLengthCompletenessUseFace1.Text, "")

If Len(strText) > 0 Then

   Me.cboLengthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboLengthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboLengthCompletenessUseFace1.Dropdown

Exit Sub

err_changeLengthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompletenessUseFace1_Enter()
On Error GoTo err_EnterLengthCompletenessUseFace1
    Me.cboLengthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompletenessUseFace1.Dropdown
Exit Sub

err_EnterLengthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompletenessUseFace1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownLengthCompletenessUseFace1
    Me.cboLengthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompletenessUseFace1.Dropdown
Exit Sub

err_KeyDownLengthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboLengthCompletenessUseFace2_Change()
On Error GoTo err_changeLengthCompletenessUseFace2
Dim strText As String

strText = Nz(Me.cboLengthCompletenessUseFace2.Text, "")

If Len(strText) > 0 Then

   Me.cboLengthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboLengthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboLengthCompletenessUseFace2.Dropdown

Exit Sub

err_changeLengthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboLengthCompletenessUseFace2_Enter()
On Error GoTo err_EnterLengthCompletenessUseFace2
    Me.cboLengthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompletenessUseFace2.Dropdown
Exit Sub

err_EnterLengthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLengthCompletenessUseFace2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownLengthCompletenessUseFace2
    Me.cboLengthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboLengthCompletenessUseFace2.Dropdown
Exit Sub

err_KeyDownLengthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboLettercode_AfterUpdate()
Dim checkLevel1, checkLevel2, checkOldGST
Dim ctl As Control
'update the GID
On Error GoTo err_lc

If Me![Lettercode] <> "K" And Me![Lettercode] <> "X" Then
    MsgBox "Are you sure you want to assign the Letter Code " & Me![Lettercode] & "?"
End If

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

' if GID is complete, check for duplicates in Level1 and Level2
If Me![Unit] <> "" And Me![Lettercode] <> "" And Me![FindNumber] <> "" Then
    'check that GID not exists
    checkLevel1 = DLookup("[GID]", "[dbo_Groundstone Level 1_2014]", "[GID] = '" & Me![GID] & "'")
    checkLevel2 = DLookup("[GID]", "[dbo_Groundstone Level 2_2014]", "[GID] = '" & Me![GID] & "'")
    checkOldGST = DLookup("[GID]", "[GroundStone: Basic_Data]", "[GID] = '" & Me![GID] & "'")

    If Not IsNull(checkLevel1) Then
        MsgBox "GID Number " & Me![GID] & " already exists in Level 1 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![Lettercode].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![Lettercode].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    ElseIf Not IsNull(checkLevel2) Then
    MsgBox "GID Number " & Me![GID] & " already exists in Level 2 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![Lettercode].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![Lettercode].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry (except the fields from Exca)
        'unlock all fields - CE June 2014
        For Each ctl In Me.Controls
            If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) Then
                ctl.Locked = False
            End If
        Next ctl
        'if GID is in old Groundstone DB just give a warning, but dont disable entry
        If Not IsNull(checkOldGST) Then
            MsgBox "GID Number " & Me![GID] & " already exists in the Old Groundstone table. Make sure you know what you are doing.", vbInformation, "Duplicate GID Number"
        End If

    End If
End If

Exit Sub

err_lc:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesBit_Change()
On Error GoTo err_changeManifactTechniquesBit
Dim strText As String

strText = Nz(Me.cboManifactTechniquesBit.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactTechniquesBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactTechniquesBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
End If

Me.cboManifactTechniquesBit.Dropdown

Exit Sub

err_changeManifactTechniquesBit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesBit_Enter()
On Error GoTo err_EnterManifactTechniquesBit
    Me.cboManifactTechniquesBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesBit.Dropdown
Exit Sub

err_EnterManifactTechniquesBit:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboManifactTechniquesBit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactTechniquesBit
    Me.cboManifactTechniquesBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesBit.Dropdown
Exit Sub

err_KeyDownManifactTechniquesBit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesBody_Change()
On Error GoTo err_changeManifactTechniquesBody
Dim strText As String

strText = Nz(Me.cboManifactTechniquesBody.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactTechniquesBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactTechniquesBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
End If

Me.cboManifactTechniquesBody.Dropdown

Exit Sub

err_changeManifactTechniquesBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesBody_Enter()
On Error GoTo err_EnterManifactTechniquesBody
    Me.cboManifactTechniquesBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesBody.Dropdown
Exit Sub

err_EnterManifactTechniquesBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesBody_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownManifactTechniquesBody
    Me.cboManifactTechniquesBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesBody.Dropdown
Exit Sub

err_KeyDownManifactTechniquesBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesButt_Change()
On Error GoTo err_changeManifactTechniquesButt
Dim strText As String

strText = Nz(Me.cboManifactTechniquesButt.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactTechniquesButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactTechniquesButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
End If

Me.cboManifactTechniquesButt.Dropdown

Exit Sub

err_changeManifactTechniquesButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesButt_Enter()
On Error GoTo err_EnterManifactTechniquesButt
    Me.cboManifactTechniquesButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesButt.Dropdown
Exit Sub

err_EnterManifactTechniquesButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesButt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactTechniquesButt
    Me.cboManifactTechniquesButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesButt.Dropdown
Exit Sub

err_KeyDownManifactTechniquesButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesMargins_Change()
On Error GoTo err_changeManifactTechniquesMargins
Dim strText As String

strText = Nz(Me.cboManifactTechniquesMargins.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactTechniquesMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactTechniquesMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
End If

Me.cboManifactTechniquesMargins.Dropdown

Exit Sub

err_changeManifactTechniquesMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesMargins_Enter()
On Error GoTo err_EnterManifactTechniquesMargins
    Me.cboManifactTechniquesMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesMargins.Dropdown
Exit Sub

err_EnterManifactTechniquesMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesMargins_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactTechniquesMargins
    Me.cboManifactTechniquesMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesMargins.Dropdown
Exit Sub

err_KeyDownManifactTechniquesMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesUseFace1_Change()
On Error GoTo err_changeManifactTechniquesUseFace1
Dim strText As String

strText = Nz(Me.cboManifactTechniquesUseFace1.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactTechniquesUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactTechniquesUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
End If

Me.cboManifactTechniquesUseFace1.Dropdown

Exit Sub

err_changeManifactTechniquesUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesUseFace1_Enter()
On Error GoTo err_EnterManifactTechniquesUseFace1
    Me.cboManifactTechniquesUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesUseFace1.Dropdown
Exit Sub

err_EnterManifactTechniquesUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactTechniquesUseFace1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactTechniquesUseFace1
    Me.cboManifactTechniquesUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Techniques] ORDER BY ScreenOrder"
    Me.cboManifactTechniquesUseFace1.Dropdown
Exit Sub

err_KeyDownManifactTechniquesUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBit_Change()
On Error GoTo err_changeManifactWearBit
Dim strText As String

strText = Nz(Me.cboManifactWearBit.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactWearBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactWearBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
End If

Me.cboManifactWearBit.Dropdown

Exit Sub

err_changeManifactWearBit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBit_Enter()
On Error GoTo err_EnterManifactWearBit
    Me.cboManifactWearBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearBit.Dropdown
Exit Sub

err_EnterManifactWearBit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactWearBit
    Me.cboManifactWearBit.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearBit.Dropdown
Exit Sub

err_KeyDownManifactWearBit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBody_Change()
On Error GoTo err_changeManifactWearBody
Dim strText As String

strText = Nz(Me.cboManifactWearBody.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactWearBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactWearBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
End If

Me.cboManifactWearBody.Dropdown

Exit Sub

err_changeManifactWearBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBody_Enter()
On Error GoTo err_EnterManifactWearBody
    Me.cboManifactWearBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearBody.Dropdown
Exit Sub

err_EnterManifactWearBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearBody_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactWearBody
    Me.cboManifactWearBody.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearBody.Dropdown
Exit Sub

err_KeyDownManifactWearBody:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearButt_Change()
On Error GoTo err_changeManifactWearButt
Dim strText As String

strText = Nz(Me.cboManifactWearButt.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactWearButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactWearButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
End If

Me.cboManifactWearButt.Dropdown

Exit Sub

err_changeManifactWearButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearButt_Enter()
On Error GoTo err_EnterManifactWearButt
    Me.cboManifactWearButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearButt.Dropdown
Exit Sub

err_EnterManifactWearButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearButt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactWearButt
    Me.cboManifactWearButt.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearButt.Dropdown
Exit Sub

err_KeyDownManifactWearButt:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearMargins_Change()
On Error GoTo err_changeManifactWearMargins
Dim strText As String

strText = Nz(Me.cboManifactWearMargins.Text, "")

If Len(strText) > 0 Then

   Me.cboManifactWearMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboManifactWearMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
End If

Me.cboManifactWearMargins.Dropdown

Exit Sub

err_changeManifactWearMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearMargins_Enter()
On Error GoTo err_EnterManifactWearMargins
    Me.cboManifactWearMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearMargins.Dropdown
Exit Sub

err_EnterManifactWearMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboManifactWearMargins_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownManifactWearMargins
    Me.cboManifactWearMargins.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Manufacturing Wear] ORDER BY ScreenOrder"
    Me.cboManifactWearMargins.Dropdown
Exit Sub

err_KeyDownManifactWearMargins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategory_Change()
On Error GoTo err_changeObjectCategory
Dim strText As String

strText = Nz(Me.cboObjectCategory.Text, "")

If Len(strText) > 0 Then

   Me.cboObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
End If

Me.cboObjectCategory.Dropdown

Exit Sub

err_changeObjectCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategory_Enter()

On Error GoTo err_EnterObjectCategory
    Me.cboObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
    Me.cboObjectCategory.Dropdown
Exit Sub

err_EnterObjectCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategory_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownObjectCategory
    Me.cboObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
    Me.cboObjectCategory.Dropdown
Exit Sub

err_KeyDownObjectCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategoryPrimary_Change()
On Error GoTo err_changeObjectCategoryPrimary
Dim strText As String

strText = Nz(Me.cboObjectCategoryPrimary.Text, "")

If Len(strText) > 0 Then

   Me.cboObjectCategoryPrimary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboObjectCategoryPrimary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
End If

Me.cboObjectCategoryPrimary.Dropdown

Exit Sub

err_changeObjectCategoryPrimary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategoryPrimary_Enter()
On Error GoTo err_EnterObjectCategoryPrimary
    Me.cboObjectCategoryPrimary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboObjectCategoryPrimary.Dropdown
Exit Sub

err_EnterObjectCategoryPrimary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategoryPrimary_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownObjectCategoryPrimary
    Me.cboObjectCategoryPrimary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboObjectCategoryPrimary.Dropdown
Exit Sub

err_KeyDownObjectCategoryPrimary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategorySecondary_Change()
On Error GoTo err_changeObjectCategorySecondary
Dim strText As String

strText = Nz(Me.cboObjectCategorySecondary.Text, "")

If Len(strText) > 0 Then

   Me.cboObjectCategorySecondary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboObjectCategorySecondary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
End If

Me.cboObjectCategorySecondary.Dropdown

Exit Sub

err_changeObjectCategorySecondary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategorySecondary_Enter()
On Error GoTo err_EnterObjectCategorySecondary
    Me.cboObjectCategorySecondary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboObjectCategorySecondary.Dropdown
Exit Sub

err_EnterObjectCategorySecondary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectCategorySecondary_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownObjectCategorySecondary
    Me.cboObjectCategorySecondary.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboObjectCategorySecondary.Dropdown
Exit Sub

err_KeyDownObjectCategorySecondary:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRawMaterial_Change()
On Error GoTo err_changeRawMaterial
Dim strText As String

strText = Nz(Me.cboRawMaterial.Text, "")

If Len(strText) > 0 Then

   Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] ORDER BY ScreenOrder"
End If

Me.cboRawMaterial.Dropdown

Exit Sub

err_changeRawMaterial:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRawMaterial_Enter()
On Error GoTo err_RawMat_Enter
    Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] ORDER BY ScreenOrder"
    Me.cboRawMaterial.Dropdown
Exit Sub

err_RawMat_Enter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRawMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_RawMat_KeyDown
    Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] ORDER BY ScreenOrder"
    Me.cboRawMaterial.Dropdown
Exit Sub

err_RawMat_KeyDown:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboResidue_Change()
On Error GoTo err_changeResidue
Dim strText As String

strText = Nz(Me.cboResidue.Text, "")

If Len(strText) > 0 Then

   Me.cboResidue.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder "
Else
    Me.cboResidue.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue] ORDER BY ScreenOrder"
End If

Me.cboResidue.Dropdown

Exit Sub

err_changeResidue:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboResidue_Enter()
On Error GoTo err_EnterResidue
    Me.cboResidue.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue] ORDER BY ScreenOrder"
    Me.cboResidue.Dropdown
Exit Sub

err_EnterResidue:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboResidue_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownResidue
    Me.cboResidue.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue] ORDER BY ScreenOrder"
    Me.cboResidue.Dropdown
Exit Sub

err_KeyDownResidue:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboResidueType_Change()
On Error GoTo err_changeResidueType
Dim strText As String

strText = Nz(Me.cboResidueType.Text, "")

If Len(strText) > 0 Then

   Me.cboResidueType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue Type] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboResidueType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue Type] ORDER BY ScreenOrder"
End If

Me.cboResidueType.Dropdown

Exit Sub

err_changeResidueType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboResidueType_Enter()
On Error GoTo err_EnterResidueType
    Me.cboResidueType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue Type] ORDER BY ScreenOrder"
    Me.cboResidueType.Dropdown
Exit Sub

err_EnterResidueType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboResidueType_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownResidueType
    Me.cboResidueType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Residue Type] ORDER BY ScreenOrder"
    Me.cboResidueType.Dropdown
Exit Sub

err_KeyDownResidueType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRimDiameterCompleteness_Change()
On Error GoTo err_changeRimDiameterCompleteness
Dim strText As String

strText = Nz(Me.cboRimDiameterCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboRimDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboRimDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboRimDiameterCompleteness.Dropdown

Exit Sub

err_changeRimDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimDiameterCompleteness_Enter()
On Error GoTo err_EnterRimDiameterCompleteness
    Me.cboRimDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboRimDiameterCompleteness.Dropdown
Exit Sub

err_EnterRimDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRimDiameterCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownRimDiameterCompleteness
    Me.cboRimDiameterCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboRimDiameterCompleteness.Dropdown
Exit Sub

err_KeyDownRimDiameterCompleteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimThicknessCompleteteness_Change()
On Error GoTo err_changeRimThicknessCompleteteness
Dim strText As String

strText = Nz(Me.cboRimThicknessCompleteteness.Text, "")

If Len(strText) > 0 Then

   Me.cboRimThicknessCompleteteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboRimThicknessCompleteteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboRimThicknessCompleteteness.Dropdown

Exit Sub

err_changeRimThicknessCompleteteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimThicknessCompleteteness_Enter()
On Error GoTo err_EnterRimThicknessCompleteteness
    Me.cboRimThicknessCompleteteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboRimThicknessCompleteteness.Dropdown
Exit Sub

err_EnterRimThicknessCompleteteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimThicknessCompleteteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownRimThicknessCompleteteness
    Me.cboRimThicknessCompleteteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboRimThicknessCompleteteness.Dropdown
Exit Sub

err_KeyDownRimThicknessCompleteteness:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimType_Change()
On Error GoTo err_changeRimType
Dim strText As String

strText = Nz(Me.cboRimType.Text, "")

If Len(strText) > 0 Then

   Me.cboRimType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV RimType] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboRimType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV RimType] ORDER BY ScreenOrder"
End If

Me.cboRimType.Dropdown

Exit Sub

err_changeRimType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRimType_Enter()

On Error GoTo err_EnterRimType
    Me.cboRimType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV RimType] ORDER BY ScreenOrder"
    Me.cboRimType.Dropdown
Exit Sub

err_EnterRimType:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboRimType_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownRimType
    Me.cboRimType.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV RimType] ORDER BY ScreenOrder"
    Me.cboRimType.Dropdown
Exit Sub

err_KeyDownRimType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSampled_Change()
On Error GoTo err_changeSampled
Dim strText As String

strText = Nz(Me.cboSampled.Text, "")

If Len(strText) > 0 Then

   Me.cboSampled.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Sampled] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder "
Else
    Me.cboSampled.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Sampled] ORDER BY ScreenOrder"
End If

Me.cboSampled.Dropdown

Exit Sub

err_changeSampled:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSampled_Enter()
On Error GoTo err_EnterSampled
    Me.cboSampled.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Sampled] ORDER BY ScreenOrder"
    Me.cboSampled.Dropdown
Exit Sub

err_EnterSampled:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSampled_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownSampled
    Me.cboSampled.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Sampled] ORDER BY ScreenOrder"
    Me.cboSampled.Dropdown
Exit Sub

err_KeyDownSampled:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapePlan_Change()
On Error GoTo err_changeShapePlan
Dim strText As String

strText = Nz(Me.cboShapePlan.Text, "")

If Len(strText) > 0 Then

   Me.cboShapePlan.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Plan] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboShapePlan.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Plan] ORDER BY ScreenOrder"
End If

Me.cboShapePlan.Dropdown

Exit Sub

err_changeShapePlan:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapePlan_Enter()
On Error GoTo err_EnterShapePlan
    Me.cboShapePlan.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Plan] ORDER BY ScreenOrder"
    Me.cboShapePlan.Dropdown
Exit Sub

err_EnterShapePlan:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapePlan_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownShapePlan
    Me.cboShapePlan.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Plan] ORDER BY ScreenOrder"
    Me.cboShapePlan.Dropdown
Exit Sub

err_KeyDownShapePlan:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeSection_Change()
On Error GoTo err_changeShapeSection
Dim strText As String

strText = Nz(Me.cboShapeSection.Text, "")

If Len(strText) > 0 Then

   Me.cboShapeSection.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Section] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboShapeSection.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Section] ORDER BY ScreenOrder"
End If

Me.cboShapeSection.Dropdown

Exit Sub

err_changeShapeSection:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeSection_Enter()
On Error GoTo err_EnterShapeSection
    Me.cboShapeSection.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Section] ORDER BY ScreenOrder"
    Me.cboShapeSection.Dropdown
Exit Sub

err_EnterShapeSection:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeSection_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownShapeSection
    Me.cboShapeSection.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Section] ORDER BY ScreenOrder"
    Me.cboShapeSection.Dropdown
Exit Sub

err_KeyDownShapeSection:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace1_Change()
On Error GoTo err_changeShapeUseFace1
Dim strText As String

strText = Nz(Me.cboShapeUseFace1.Text, "")

If Len(strText) > 0 Then

   Me.cboShapeUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboShapeUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
End If

Me.cboShapeUseFace1.Dropdown

Exit Sub

err_changeShapeUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace1_Enter()
On Error GoTo err_EnterShapeUseFace1
    Me.cboShapeUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
    Me.cboShapeUseFace1.Dropdown
Exit Sub

err_EnterShapeUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownShapeUseFace1
    Me.cboShapeUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
    Me.cboShapeUseFace1.Dropdown
Exit Sub

err_KeyDownShapeUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace2_Change()
On Error GoTo err_changeShapeUseFace2
Dim strText As String

strText = Nz(Me.cboShapeUseFace2.Text, "")

If Len(strText) > 0 Then

   Me.cboShapeUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboShapeUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
End If

Me.cboShapeUseFace2.Dropdown

Exit Sub

err_changeShapeUseFace2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace2_Enter()

On Error GoTo err_EnterShapeUseFace2
    Me.cboShapeUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
    Me.cboShapeUseFace2.Dropdown
Exit Sub

err_EnterShapeUseFace2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboShapeUseFace2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownShapeUseFace2
    Me.cboShapeUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Shape Use Face] ORDER BY ScreenOrder"
    Me.cboShapeUseFace2.Dropdown
Exit Sub

err_KeyDownShapeUseFace2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSizeCategory_Change()
On Error GoTo err_changeSizeCategory
Dim strText As String

strText = Nz(Me.cboSizeCategory.Text, "")

If Len(strText) > 0 Then

   Me.cboSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Size Category] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Size Category] ORDER BY ScreenOrder"
End If

Me.cboSizeCategory.Dropdown

Exit Sub

err_changeSizeCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSizeCategory_Enter()
On Error GoTo err_EnterSizeCategory
    Me.cboSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Size Category] ORDER BY ScreenOrder"
    Me.cboSizeCategory.Dropdown
Exit Sub

err_EnterSizeCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSizeCategory_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownSizeCategory
    Me.cboSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Size Category] ORDER BY ScreenOrder"
    Me.cboSizeCategory.Dropdown
Exit Sub

err_KeyDownSizeCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSourceRockSizeCategory_Change()
On Error GoTo err_changeSourceRockSizeCategory
Dim strText As String

strText = Nz(Me.cboSourceRockSizeCategory.Text, "")

If Len(strText) > 0 Then

   Me.cboSourceRockSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Source Rock Size Categories] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboSourceRockSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Source Rock Size Categories] ORDER BY ScreenOrder"
End If

Me.cboSourceRockSizeCategory.Dropdown

Exit Sub

err_changeSourceRockSizeCategory:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboSourceRockSizeCategory_Enter()
On Error GoTo err_EnterSourceRockSizeCategory
    Me.cboSourceRockSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Source Rock Size Categories] ORDER BY ScreenOrder"
    Me.cboSourceRockSizeCategory.Dropdown
Exit Sub

err_EnterSourceRockSizeCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSourceRockSizeCategory_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_KeyDownSourceRockSizeCategory
    Me.cboSourceRockSizeCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Source Rock Size Categories] ORDER BY ScreenOrder"
    Me.cboSourceRockSizeCategory.Dropdown
Exit Sub

err_KeyDownSourceRockSizeCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSurfaceCondition_Change()
On Error GoTo err_changeSurfaceCondition
Dim strText As String

strText = Nz(Me.cboSurfaceCondition.Text, "")

If Len(strText) > 0 Then

   Me.cboSurfaceCondition.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Surface Condition] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboSurfaceCondition.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Surface Condition] ORDER BY ScreenOrder"
End If

Me.cboSurfaceCondition.Dropdown

Exit Sub

err_changeSurfaceCondition:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSurfaceCondition_Enter()

On Error GoTo err_EnterSurfaceCondition
    Me.cboSurfaceCondition.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Surface Condition] ORDER BY ScreenOrder"
    Me.cboSurfaceCondition.Dropdown
Exit Sub

err_EnterSurfaceCondition:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboSurfaceCondition_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownSurfaceCondition
    Me.cboSurfaceCondition.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Surface Condition] ORDER BY ScreenOrder"
    Me.cboSurfaceCondition.Dropdown
Exit Sub

err_KeyDownSurfaceCondition:
    Call General_Error_Trap
End Sub

Private Sub cboTexture_Change()
On Error GoTo err_changeTexture
Dim strText As String

strText = Nz(Me.cboTexture.Text, "")

If Len(strText) > 0 Then

   Me.cboTexture.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Texture] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboTexture.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Texture] ORDER BY ScreenOrder"
End If

Me.cboTexture.Dropdown

Exit Sub

err_changeTexture:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboTexture_Enter()

On Error GoTo err_EnterTexture
    Me.cboTexture.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Texture] ORDER BY ScreenOrder"
    Me.cboTexture.Dropdown
Exit Sub

err_EnterTexture:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboTexture_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownTexture
    Me.cboTexture.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Texture] ORDER BY ScreenOrder"
    Me.cboTexture.Dropdown
Exit Sub

err_KeyDownTexture:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboThicknessCompleteness_Change()
On Error GoTo err_changeThicknessCompleteness
Dim strText As String

strText = Nz(Me.cboThicknessCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboThicknessCompleteness.Dropdown

Exit Sub

err_changeThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboThicknessCompleteness_Enter()
On Error GoTo err_EnterThicknessCompleteness
    Me.cboThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboThicknessCompleteness.Dropdown
Exit Sub

err_EnterThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboThicknessCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownThicknessCompleteness
    Me.cboThicknessCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboThicknessCompleteness.Dropdown
Exit Sub

err_KeyDownThicknessCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboToolSurfaceModif_Change()
On Error GoTo err_changeToolSurfaceModif
Dim strText As String

strText = Nz(Me.cboToolSurfaceModif.Text, "")

If Len(strText) > 0 Then

   Me.cboToolSurfaceModif.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Modification Tool Surface] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder "
Else
    Me.cboToolSurfaceModif.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Modification Tool Surface] ORDER BY ScreenOrder"
End If

Me.cboToolSurfaceModif.Dropdown

Exit Sub

err_changeToolSurfaceModif:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboToolSurfaceModif_Enter()
On Error GoTo err_EnterToolSurfaceModif
    Me.cboToolSurfaceModif.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Modification Tool Surface] ORDER BY ScreenOrder"
    Me.cboToolSurfaceModif.Dropdown
Exit Sub

err_EnterToolSurfaceModif:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboToolSurfaceModif_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownToolSurfaceModif
    Me.cboToolSurfaceModif.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Modification Tool Surface] ORDER BY ScreenOrder"
    Me.cboToolSurfaceModif.Dropdown
Exit Sub

err_KeyDownToolSurfaceModif:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseFaceNumberLocation_Change()
On Error GoTo err_changeUseFaceNumberLocation
Dim strText As String

strText = Nz(Me.cboUseFaceNumberLocation.Text, "")

If Len(strText) > 0 Then

   Me.cboUseFaceNumberLocation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Faces] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboUseFaceNumberLocation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Faces] ORDER BY ScreenOrder"
End If

Me.cboUseFaceNumberLocation.Dropdown

Exit Sub

err_changeUseFaceNumberLocation:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboUseFaceNumberLocation_Enter()

On Error GoTo err_EnterUseFaceNumberLocation
    Me.cboUseFaceNumberLocation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Faces] ORDER BY ScreenOrder"
    Me.cboUseFaceNumberLocation.Dropdown
Exit Sub

err_EnterUseFaceNumberLocation:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseFaceNumberLocation_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownUseFaceNumberLocation
    Me.cboUseFaceNumberLocation.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Faces] ORDER BY ScreenOrder"
    Me.cboUseFaceNumberLocation.Dropdown
Exit Sub

err_KeyDownUseFaceNumberLocation:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearSecondUse_Change()
On Error GoTo err_changeUseWearSecondUse
Dim strText As String

strText = Nz(Me.cboUseWearSecondUse.Text, "")

If Len(strText) > 0 Then

   Me.cboUseWearSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder "
Else
    Me.cboUseWearSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
End If

Me.cboUseWearSecondUse.Dropdown

Exit Sub

err_changeUseWearSecondUse:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearSecondUse_Enter()
On Error GoTo err_EnterUseWearSecondUse
    Me.cboUseWearSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearSecondUse.Dropdown
Exit Sub

err_EnterUseWearSecondUse:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cboUseWearSecondUse_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownUseWearSecondUse
    Me.cboUseWearSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearSecondUse.Dropdown
Exit Sub

err_KeyDownUseWearSecondUse:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearUF1_Change()
On Error GoTo err_changeUseWearUF1
Dim strText As String

strText = Nz(Me.cboUseWearUF1.Text, "")

If Len(strText) > 0 Then

   Me.cboUseWearUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboUseWearUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
End If

Me.cboUseWearUF1.Dropdown

Exit Sub

err_changeUseWearUF1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearUF1_Enter()
On Error GoTo err_EnterUseWearUF1
    Me.cboUseWearUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearUF1.Dropdown
Exit Sub

err_EnterUseWearUF1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearUF1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownUseWearUF1
    Me.cboUseWearUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearUF1.Dropdown
Exit Sub

err_KeyDownUseWearUF1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearUF2_Change()
Dim strText As String

strText = Nz(Me.cboUseWearUF2.Text, "")

If Len(strText) > 0 Then

   Me.cboUseWearUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboUseWearUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
End If

Me.cboUseWearUF2.Dropdown

Exit Sub

err_changeUseWearUF2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboUseWearUF2_Enter()
On Error GoTo err_EnterUseWearUF2
    Me.cboUseWearUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearUF2.Dropdown
Exit Sub

err_EnterUseWearUF2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboUseWearUF2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownUseWearUF2
    Me.cboUseWearUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Use Wear] ORDER BY ScreenOrder"
    Me.cboUseWearUF2.Dropdown
Exit Sub

err_KeyDownUseWearUF2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreePrimUseUF1_Change()
On Error GoTo err_changeWearDegreePrimUseUF1
Dim strText As String

strText = Nz(Me.cboWearDegreePrimUseUF1.Text, "")

If Len(strText) > 0 Then

   Me.cboWearDegreePrimUseUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWearDegreePrimUseUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
End If

Me.cboWearDegreePrimUseUF1.Dropdown

Exit Sub

err_changeWearDegreePrimUseUF1:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboWearDegreePrimUseUF1_Enter()
On Error GoTo err_EnterWearDegreePrimUseUF1
    Me.cboWearDegreePrimUseUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreePrimUseUF1.Dropdown
Exit Sub

err_EnterWearDegreePrimUseUF1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreePrimUseUF1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWearDegreePrimUseUF1
    Me.cboWearDegreePrimUseUF1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreePrimUseUF1.Dropdown
Exit Sub

err_KeyDownWearDegreePrimUseUF1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreePrimUseUF2_Change()
On Error GoTo err_changeWearDegreePrimUseUF2
Dim strText As String

strText = Nz(Me.cboWearDegreePrimUseUF2.Text, "")

If Len(strText) > 0 Then

   Me.cboWearDegreePrimUseUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWearDegreePrimUseUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
End If

Me.cboWearDegreePrimUseUF2.Dropdown

Exit Sub

err_changeWearDegreePrimUseUF2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreePrimUseUF2_Enter()
On Error GoTo err_EnterWearDegreePrimUseUF2
    Me.cboWearDegreePrimUseUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreePrimUseUF2.Dropdown
Exit Sub

err_EnterWearDegreePrimUseUF2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreePrimUseUF2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWearDegreePrimUseUF2
    Me.cboWearDegreePrimUseUF2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreePrimUseUF2.Dropdown
Exit Sub

err_KeyDownWearDegreePrimUseUF2:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreeSecondUse_Change()
On Error GoTo err_changeWearDegreeSecondUse
Dim strText As String

strText = Nz(Me.cboWearDegreeSecondUse.Text, "")

If Len(strText) > 0 Then

   Me.cboWearDegreeSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWearDegreeSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
End If

Me.cboWearDegreeSecondUse.Dropdown

Exit Sub

err_changeWearDegreeSecondUse:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreeSecondUse_Enter()
On Error GoTo err_EnterWearDegreeSecondUse
    Me.cboWearDegreeSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreeSecondUse.Dropdown
Exit Sub

err_EnterWearDegreeSecondUse:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWearDegreeSecondUse_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWearDegreeSecondUse
    Me.cboWearDegreeSecondUse.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Degree Wear] ORDER BY ScreenOrder"
    Me.cboWearDegreeSecondUse.Dropdown
Exit Sub

err_KeyDownWearDegreeSecondUse:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompleteness_Change()
On Error GoTo err_changeWidthCompleteness
Dim strText As String

strText = Nz(Me.cboWidthCompleteness.Text, "")

If Len(strText) > 0 Then

   Me.cboWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboWidthCompleteness.Dropdown

Exit Sub

err_changeWidthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompleteness_Enter()
On Error GoTo err_EnterWidthCompleteness
    Me.cboWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompleteness.Dropdown
Exit Sub

err_EnterWidthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompleteness_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWidthCompleteness
    Me.cboWidthCompleteness.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompleteness.Dropdown
Exit Sub

err_KeyDownWidthCompleteness:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompletenessUseFace1_Change()
On Error GoTo err_changeWidthCompletenessUseFace1
Dim strText As String

strText = Nz(Me.cboWidthCompletenessUseFace1.Text, "")

If Len(strText) > 0 Then

   Me.cboWidthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWidthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboWidthCompletenessUseFace1.Dropdown

Exit Sub

err_changeWidthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboWidthCompletenessUseFace1_Enter()
On Error GoTo err_EnterWidthCompletenessUseFace1
    Me.cboWidthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompletenessUseFace1.Dropdown
Exit Sub

err_EnterWidthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompletenessUseFace1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWidthCompletenessUseFace1
    Me.cboWidthCompletenessUseFace1.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompletenessUseFace1.Dropdown
Exit Sub

err_KeyDownWidthCompletenessUseFace1:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboWidthCompletenessUseFace2_Change()
On Error GoTo err_changeWidthCompletenessUseFace2
Dim strText As String

strText = Nz(Me.cboWidthCompletenessUseFace2.Text, "")

If Len(strText) > 0 Then

   Me.cboWidthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboWidthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
End If

Me.cboWidthCompletenessUseFace2.Dropdown

Exit Sub

err_changeWidthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboWidthCompletenessUseFace2_Enter()
On Error GoTo err_EnterWidthCompletenessUseFace2
    Me.cboWidthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompletenessUseFace2.Dropdown
Exit Sub

err_EnterWidthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboWidthCompletenessUseFace2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownWidthCompletenessUseFace2
    Me.cboWidthCompletenessUseFace2.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Completeness] ORDER BY ScreenOrder"
    Me.cboWidthCompletenessUseFace2.Dropdown
Exit Sub

err_KeyDownWidthCompletenessUseFace2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub FindNumber_AfterUpdate()
Dim checkLevel1, checkLevel2, checkOldGST
Dim ctl As Control
Dim retVal, inputname, sql
'update the GID
On Error GoTo err_fn

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]
' if GID is complete, check for duplicates in Level1 and Level2
If Me![Unit] <> "" And Me![Lettercode] <> "" And Me![FindNumber] <> "" Then
    'check that GID not exists
    checkLevel1 = DLookup("[GID]", "[dbo_Groundstone Level 1_2014]", "[GID] = '" & Me![GID] & "'")
    checkLevel2 = DLookup("[GID]", "[dbo_Groundstone Level 2_2014]", "[GID] = '" & Me![GID] & "'")
    checkOldGST = DLookup("[GID]", "[GroundStone: Basic_Data]", "[GID] = '" & Me![GID] & "'")

    If Not IsNull(checkLevel1) Then
        MsgBox "GID Number " & Me![GID] & " already exists in Level 1 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![FindNumber].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![FindNumber].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    ElseIf Not IsNull(checkLevel2) Then
    MsgBox "GID Number " & Me![GID] & " already exists in Level 2 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![FindNumber].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![FindNumber].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry (except the fields from Exca)
        'unlock all fields - CE June 2014
        For Each ctl In Me.Controls
            If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) Then
                ctl.Locked = False
            End If
        Next ctl
        'if GID is in old Groundstone DB just give a warning, but dont disable entry
        'changed 2015 - now the initials of the user who signs an item out of the old groundstone, are fed back there
        If Not IsNull(checkOldGST) Then
            'MsgBox " Make sure you know what you are doing.", vbInformation, "Duplicate GID Number"
            retVal = MsgBox("GID Number " & Me![GID] & " already exists in the Old Groundstone table. Do you want to rerecord it?", vbQuestion + vbYesNo, "GID in old Groundstone")
            If retVal = vbYes Then
                inputname = InputBox("Please enter your initials for singning the item out of the old groundstone DB:", "Analyst Name")
                If inputname <> "" Then
                    sql = "UPDATE [dbo_Groundstone: Basic_Data_Pre2013_WithReRecordingFlag] SET [ReRecorded] = '" & inputname & " " & Date & "' WHERE [GID] = '" & Me![GID] & "';"
                    DoCmd.RunSQL sql
                Else
                End If
            Else
                'SendKeys "{ESC}"
                Me.Undo
            End If
        Else
        End If
    End If
End If

Exit Sub

err_fn:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
' added locks to disable entry of other fields until we control for duplicate GID
' in Level 1 and Level 2 tables when new GID is added
' we remove the locks later when we do the check after updating
' the three fields that allow for entry: Unit, Lettercode, Fieldnumber
' CE June 2014

On Error GoTo err_fbins

Dim ctl As Control
    'lock all fields except Unit, Letter, FindNo - CE June 2014
    For Each ctl In Me.Controls
        If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) And Not (ctl.Name = "Unit" Or ctl.Name = "cboLettercode" Or ctl.Name = "FindNumber") Then
            ctl.Locked = True
        End If
    '    If (ctl.ControlType = acTextBox Or ctl.Name = "cboAnalyst") And Not (ctl.Name = "Unit" Or ctl.Name = "Lettercode" Or ctl.Name = "FindNumber") Then
    '        ctl.Locked = True
    '    End If
    Next ctl
            
    
    
 Exit Sub

err_fbins:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub GSNo_AfterUpdate()
Dim checkLevel1, checkLevel2
On Error GoTo err_gsnum
' check if GSno is already in use
' - not clear why the DSlookup for checkLevel1 needs string as argument
' - but if I don't do it this way I get a type mismatch
If Me![GSNo] <> "" Then
    'checkLevel1 = DLookup("[GSno]", "[dbo_Groundstone Level 1_2014]", "[GSno] = " & Me![GSNo])
    checkLevel1 = DLookup("[GSno]", "[dbo_Groundstone Level 1_2014]", "[GSno] = '" & Me![GSNo] & "'")
    checkLevel2 = DLookup("[GSno]", "[dbo_Groundstone Level 2_2014]", "[GSno] = " & Me![GSNo])
    'checkLevel2 = DLookup("[GSno]", "[dbo_Groundstone Level 2_2014]", "[GSno] = '" & Me![GSNo] & "'")
    
    If Not IsNull(checkLevel1) Then
        MsgBox "GST Number " & Me![GSNo] & " already exists in Level 1 table.", vbExclamation, "Duplicate GST Number"
        
        If Not IsNull(Me![GSNo].OldValue) Then
            'return field to old value if there was one
            Me![GSNo] = Me![GSNo].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "cboAnalyst"
            DoCmd.GoToControl "GSno"
            Me![GSNo].SetFocus
            Me![GSNo] = Null
        End If
    End If
    If Not IsNull(checkLevel2) Then
    MsgBox "GST Number " & Me![GSNo] & " already exists in Level 2 table.", vbExclamation, "Duplicate GST Number"
        
        If Not IsNull(Me![Unit].OldValue) Then
            'return field to old value if there was one
            Me![GSNo] = Me![GSNo].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            'DoCmd.GoToControl "cboAnalyst"
            DoCmd.GoToControl "GSNo"
            Me![GSNo].SetFocus
            Me![GSNo] = Null
        End If
    End If
End If
Exit Sub

err_gsnum:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Unit_AfterUpdate()
Dim checkLevel1, checkLevel2, checkOldGST
Dim ctl As Control
'update the GID
On Error GoTo err_unit

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]
' if GID is complete, check for duplicates in Level1 and Level2
If Me![Unit] <> "" And Me![Lettercode] <> "" And Me![FindNumber] <> "" Then
    'check that GID not exists
    checkLevel1 = DLookup("[GID]", "[dbo_Groundstone Level 1_2014]", "[GID] = '" & Me![GID] & "'")
    checkLevel2 = DLookup("[GID]", "[dbo_Groundstone Level 2_2014]", "[GID] = '" & Me![GID] & "'")
    checkOldGST = DLookup("[GID]", "[GroundStone: Basic_Data]", "[GID] = '" & Me![GID] & "'")

    If Not IsNull(checkLevel1) Then
        MsgBox "GID Number " & Me![GID] & " already exists in Level 1 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![Unit].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![Unit].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    ElseIf Not IsNull(checkLevel2) Then
    MsgBox "GID Number " & Me![GID] & " already exists in Level 2 table.", vbExclamation, "Duplicate GID Number"
        
        If Not IsNull(Me![Unit].OldValue) Then
            'return field to old value if there was one
            Me![Unit] = Me![Unit].OldValue
        Else
            'oh the joys, to keep the focus on unit have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "GID"
            DoCmd.GoToControl "Unit"
            Me![Unit].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry (except the fields from Exca)
        'unlock all fields - CE June 2014
        For Each ctl In Me.Controls
            If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) Then
                ctl.Locked = False
            End If
        Next ctl
        'if GID is in old Groundstone DB just give a warning, but dont disable entry
        If Not IsNull(checkOldGST) Then
            MsgBox "GID Number " & Me![GID] & " already exists in the Old Groundstone table. Make sure you know what you are doing.", vbInformation, "Duplicate GID Number"
        End If

    End If
End If

Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_DblClick(Cancel As Integer)

On Error GoTo Err_cmdUnitDesc_Click

If Me![Unit] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[dbo_Groundstone: Unit Description_2014]", "[Unit] = " & Me![Unit])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [dbo_Groundstone: Unit Description_2014] ([Unit]) VALUES (" & Me![Unit] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "Frm_GS_UnitDescription_2014", acNormal, , "[Unit] = " & Me![Unit], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub

End Sub
