Option Compare Database

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

Private Sub cboGrObjectCategory_Change()
On Error GoTo err_changeGrObjectCategory
Dim strText As String

strText = Nz(Me.cboGrObjectCategory.Text, "")

If Len(strText) > 0 Then

   Me.cboGrObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboGrObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
End If

Me.cboGrObjectCategory.Dropdown

Exit Sub

err_changeGrObjectCategory:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboGrObjectCategory_Enter()
On Error GoTo err_EnterGrObjectCategory
    Me.cboGrObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
    Me.cboGrObjectCategory.Dropdown
Exit Sub

err_EnterGrObjectCategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboGrObjectCategory_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownGrObjectCategory
    Me.cboGrObjectCategory.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category] ORDER BY ScreenOrder"
    Me.cboGrObjectCategory.Dropdown
Exit Sub

err_KeyDownGrObjectCategory:
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

Private Sub cboPrimaryObject_Change()
On Error GoTo err_changePrimaryObject
Dim strText As String

strText = Nz(Me.cboPrimaryObject.Text, "")

If Len(strText) > 0 Then

   Me.cboPrimaryObject.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] WHERE [TextEquivalent] like '" & strText & "%' OR [Code] like '" & strText & "%' ORDER BY ScreenOrder"
Else
    Me.cboPrimaryObject.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
End If

Me.cboPrimaryObject.Dropdown

Exit Sub

err_changePrimaryObject:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboPrimaryObject_Enter()
On Error GoTo err_EnterPrimaryObject
    Me.cboPrimaryObject.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboPrimaryObject.Dropdown
Exit Sub

err_EnterPrimaryObject:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboPrimaryObject_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownPrimaryObject
    Me.cboPrimaryObject.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Object Category Detail] ORDER BY ScreenOrder"
    Me.cboPrimaryObject.Dropdown
Exit Sub

err_KeyDownPrimaryObject:
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
On Error GoTo err_EnterRawMaterial
    Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] ORDER BY ScreenOrder"
    Me.cboRawMaterial.Dropdown
Exit Sub

err_EnterRawMaterial:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRawMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_KeyDownRawMaterial
    Me.cboRawMaterial.RowSource = "SELECT [Code], [TextEquivalent] FROM [dbo_Groundstone LOV Raw Material] ORDER BY ScreenOrder"
    Me.cboRawMaterial.Dropdown
Exit Sub

err_KeyDownRawMaterial:
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

Private Sub FindNumber_AfterUpdate()
'update the GID
On Error GoTo err_fn
Dim checkLevel1, checkLevel2, checkOldGST
Dim ctl As Control
Dim retVal, inputname, sql

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

On Error GoTo err_fbi

Dim ctl As Control
    'lock all fields except Unit, Letter, FindNo - CE June 2014
    For Each ctl In Me.Controls
        If (ctl.ControlType = acTextBox Or ctl.Name = "cboAnalyst") And Not (ctl.Name = "Unit" Or ctl.Name = "Lettercode" Or ctl.Name = "FindNumber") Then
            ctl.Locked = True
        End If
    Next ctl
 Exit Sub

err_fbi:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub GSNo_AfterUpdate()
Dim checkLevel1, checkLevel2
On Error GoTo err_gsno
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
            DoCmd.GoToControl "cboAnalyst"
            DoCmd.GoToControl "GSNo"
            Me![GSNo].SetFocus
            Me![GSNo] = Null
        End If
    End If
End If
Exit Sub

err_gsno:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()
Dim checkLevel1, checkLevel2, checkOldGST
Dim ctl As Control
Dim retVal, inputname, sql

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
        'changed 2015 (DL): when signing out a previously recorded item, the initials are automatically stored
        If Not IsNull(checkOldGST) Then
            'MsgBox " Make sure you know what you are doing.", vbInformation, "Duplicate GID Number"
            retVal = MsgBox("GID Number " & Me![GID] & " already exists in the Old Groundstone table. Do you want to rerecord it?", vbQuestion + vbYesNo, "GID in old Groundstone")
            If retVal = vbYes Then
                inputname = InputBox("Please enter your initials for singning the item out of the old groundstone DB:", "Analyst Name")
                If inputname <> "" Then
                    sql = "UPDATE [Groundstone: Basic_Data_Pre2013_WithReRecordingFlag] SET [ReRecorded] = '" & inputname & " " & Date & "' WHERE [GID] = '" & Me![GID] & "';"
                    'DoCmd.RunSQL sql
                    Debug.Print sql
                Else
                End If
            Else
                'SendKeys "{ESC}"
                Me.Undo
            End If
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
