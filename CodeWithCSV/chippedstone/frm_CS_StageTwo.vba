Option Compare Database
Option Explicit
Private Sub UpdateGID()
'this is called by this form to update the GID field if either the
'Unit, find letter or number fields are altered
' SAJ
On Error GoTo err_UpdateGID

'get old GID before update
Dim oldGID
oldGID = Me![GID]

'update GID to new values
 Me![GID] = [Forms]![frm_CS_stagetwo]![Unit] & "." & [Forms]![frm_CS_stagetwo]![Letter] & [Forms]![frm_CS_stagetwo]![Number]
    
 '2010 - this wasn't updating the subtable! Now make sure all fields in subtable updated
Me![frm_subform_category].Form![GID] = Me![GID]
If Not IsNull(Me![Unit]) Then Me![frm_subform_category].Form![Unit] = Me![Unit]
If Not IsNull(Me![LetterCode]) Then Me![frm_subform_category].Form![LetterCode] = Me![LetterCode]
If Not IsNull(Me![FindNumber]) Then Me![frm_subform_category].Form![FindNumber] = Me![FindNumber]

'and make sure new 2010 flint detail sub table also update
 Dim flintdetail, sql
flintdetail = DLookup("[GID]", "[ChippedStone_StageTwo_Data_FlintDetail]", "[GID] = '" & oldGID & "'")
    If Not IsNull(flintdetail) Then
        'present so must update values
        sql = "UPDATE [ChippedStone_StageTwo_Data_FlintDetail] SET [Unit] = " & [Forms]![frm_CS_stagetwo]![Unit] & ", [Lettercode] ='" & [Forms]![frm_CS_stagetwo]![Letter] & "', [FindNumber] =" & [Forms]![frm_CS_stagetwo]![Number] & ", [GID] ='" & Me![GID] & "' WHERE [GID] = '" & oldGID & "';"
        
        DoCmd.RunSQL sql
    End If

Exit Sub

err_UpdateGID:
    Call General_Error_Trap
    Exit Sub

End Sub
Private Sub cboCategory_AfterUpdate()
On Error GoTo err_cbocategory

Dim retVal, checknum, sql
Me![frm_subform_category].Visible = True
If IsNull(Me![frm_subform_category]![GID]) Then
    'this record has no blade or core specific data so can allow alteration without a problem
    
    If Me![cboCategory] = "Blade/Flake" Then
        Me![frm_subform_category].SourceObject = "Frm_subform_BladeFlake"
        Me![frm_subform_category].Height = "4400"
        Me![lblsubform].Caption = "Blade/Flake"
        Me![frm_subform_category].Form![GID] = Me![GID]
    Else
        Me![frm_subform_category].SourceObject = "Frm_subform_Cores"
        'Me![frm_subform_category].Height = "1250"
        'made same height in 2010 as team wanted fields after subforms
        Me![frm_subform_category].Height = "4400"
        Me![lblsubform].Caption = "Core"
        Me![frm_subform_category].Form![GID] = Me![GID]
    End If
Else
    If Me![cboCategory] = "Blade/Flake" Then
        'was core
        checknum = DLookup("[GID]", "[ChippedStone_Core]", "[GID] = '" & Me![GID] & "'")
        If Not IsNull(checknum) Then
            retVal = MsgBox("This action means you will lose all of the information entered into the Core fields below. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![cboCategory] = "Core"
                Exit Sub
            Else
                sql = "DELETE FROM [ChippedStone_Core] WHERE [GID] = '" & Me![GID] & "';"
                DoCmd.RunSQL sql
            End If
        End If
        Me![frm_subform_category].SourceObject = "Frm_subform_BladeFlake"
        Me![frm_subform_category].Height = "4400"
        Me![lblsubform].Caption = "Blade/Flake"
        Me![frm_subform_category].Form![GID] = Me![GID]
    Else
        'was blade flake
        checknum = DLookup("[GID]", "[ChippedStone_BladeFlake]", "[GID] = '" & Me![GID] & "'")
        If Not IsNull(checknum) Then
            retVal = MsgBox("This action means you will lose all of the information entered into the Blade/Flake fields below. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![cboCategory] = "Blade/Flake"
                Exit Sub
            Else
                sql = "DELETE FROM [ChippedStone_BladeFlake] WHERE [GID] = '" & Me![GID] & "';"
                DoCmd.RunSQL sql
            End If
        End If
        Me![frm_subform_category].SourceObject = "Frm_subform_Cores"
        'Me![frm_subform_category].Height = "1250"
        'made same height in 2010 as team wanted fields after subforms
        Me![frm_subform_category].Height = "4400"
        Me![lblsubform].Caption = "Core"
        Me![frm_subform_category].Form![GID] = Me![GID]
    End If
End If

'new 2010 - ultimately when debitage category is clean this should become the driver for what is bladeFlake OR core but
'until that is is this category field will remain and filter down debitage category according to the selection:
If Me![cboCategory] = "Blade/Flake" Then
    Me![cboDebitageCat].RowSource = "SELECT [ChippedStoneLOV_DebitageCat].[DebitageCategory] FROM ChippedStoneLOV_DebitageCat WHERE DebitageCategory <> 'Core' ORDER BY [ChippedStoneLOV_DebitageCat].[DebitageCategory]; "
Else
    Me![cboDebitageCat].RowSource = "SELECT [ChippedStoneLOV_DebitageCat].[DebitageCategory] FROM ChippedStoneLOV_DebitageCat WHERE DebitageCategory = 'Core' ORDER BY [ChippedStoneLOV_DebitageCat].[DebitageCategory]; "
End If

Exit Sub

err_cbocategory:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDebitageCat_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cboDeb

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboDebitageCat].Undo
Exit Sub

err_cboDeb:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    Me![GID].Enabled = True
    DoCmd.GoToControl Me![GID].Name
    DoCmd.FindRecord Me![cboFind]
    DoCmd.GoToControl Me![txtBag].Name
    Me![GID].Enabled = False
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboFind_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRawMaterialType_AfterUpdate()
'raw material type is linked to source from this list
'new 2010
On Error GoTo err_cboRawMT

    If Me![cboRawMaterialType].Column(1) <> "" Then
        If Me![cboSource] = "" Or IsNull(Me![cboSource]) Then
            Me![cboSource] = Me![cboRawMaterialType].Column(1)
        ElseIf Me![cboSource] <> Me![cboRawMaterialType].Column(1) Then
            MsgBox "The old source field value of: " & Me![cboSource] & " will now be updated with the new source associated with this Raw Material: " & Me![cboRawMaterialType].Column(1), vbInformation, "Source Update"
            Me![cboSource] = Me![cboRawMaterialType].Column(1)
        End If
    End If
Exit Sub

err_cboRawMT:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboTechnology_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_TechNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboTechnology].Undo
Exit Sub

err_TechNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "txtBag"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
On Error GoTo Err_gofirst_Click


    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_gofirst_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoLast_Click()
On Error GoTo Err_goLast_Click


    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_goLast_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoNext_Click()
On Error GoTo Err_goNext_Click


    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_goNext_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoPrev_Click()
On Error GoTo Err_goPrev_Click


    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_goPrev_Click:
    Call General_Error_Trap
    
End Sub
Private Sub cmdGoToFlint_Click()
'new 2010
'check if record exists in stage two FLINT table and if not put it there ready for data entry
'saj
On Error GoTo err_stagetwo_flint

If Me![GID] <> "" Then
    Dim stagetwo, sql, LetterCode, findnum
    stagetwo = DLookup("[GID]", "[ChippedStone_StageTwo_Data_FlintDetail]", "[GID] = '" & Me![GID] & "'")
    If IsNull(stagetwo) Then
        'not there yet
        sql = "INSERT INTO [ChippedStone_StageTwo_Data_FlintDetail] ([Unit], [LetterCode], [FindNumber], [GID]) VALUES (" & Me![Unit] & ", '" & Me!Letter & "'," & Me![Number] & ",'" & Me![Unit] & "." & Me![Letter] & Me![Number] & "');"
        DoCmd.RunSQL sql
           
    End If
    DoCmd.OpenForm "Frm_pop_StageTwo_FlintDetail", acNormal, , "[GID] = '" & Me![GID] & "'"
    
Else
    MsgBox "Please enter the bag number and the unit number first", vbExclamation, "Insufficient Data"
End If
Exit Sub

err_stagetwo_flint:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![txtGID] <> "" Then
        DoCmd.OpenForm "Frm_Pop_DataOutputOptions", acNormal, , , , acDialog, Me![txtGID] & ";" & Me![worked?]
    Else
        MsgBox "The output options form cannot be shown when there is no GID on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Condition_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_ConditionNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Condition].Undo
Exit Sub

err_ConditionNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'Set up form display
On Error GoTo err_current

If Me![cboCategory] = "Blade/Flake" Then
    Me![frm_subform_category].SourceObject = "frm_subform_BladeFlake"
    Me![frm_subform_category].Height = "4400"
    Me![lblsubform].Caption = "Blade/Flake"
    Me![frm_subform_category].Visible = True
ElseIf Me![cboCategory] = "Core" Then
    Me![frm_subform_category].SourceObject = "frm_subform_Cores"
    'Me![frm_subform_category].Height = "1250"
    'made same height in 2010 as team wanted fields after subforms
    Me![frm_subform_category].Height = "4400"
    Me![lblsubform].Caption = "Core"
    Me![frm_subform_category].Visible = True
Else
    Me![frm_subform_category].Visible = False
End If

'new 2010 - ultimately when debitage category is clean this should become the driver for what is bladeFlake OR core but
'until that is is this category field will remain and filter down debitage category according to the selection:
If Me![cboCategory] = "Blade/Flake" Then
    Me![cboDebitageCat].RowSource = "SELECT [ChippedStoneLOV_DebitageCat].[DebitageCategory] FROM ChippedStoneLOV_DebitageCat WHERE DebitageCategory <> 'Core' ORDER BY [ChippedStoneLOV_DebitageCat].[DebitageCategory]; "
Else
    Me![cboDebitageCat].RowSource = "SELECT [ChippedStoneLOV_DebitageCat].[DebitageCategory] FROM ChippedStoneLOV_DebitageCat WHERE DebitageCategory = 'Core' ORDER BY [ChippedStoneLOV_DebitageCat].[DebitageCategory]; "
End If


'NEW 2010 FOR Adam - and extending stage 2
'when piece of flint is raw material of current record then show Go to Flint button
If Me![RawMaterial] = "Flint" Then
    Me![cmdGoToFlint].Visible = True
    Me![cboRawMaterialType].RowSource = "SELECT [ChippedStoneLOV_RawMaterialType].[RawMaterialType], [ChippedStoneLOV_RawMaterialType].[Source] FROM ChippedStoneLOV_RawMaterialType ORDER BY [ChippedStoneLOV_RawMaterialType].[RawMaterialType];"
    Me![cboRawMaterialType].LimitToList = False
    Me![cboSource].Locked = False
    Me![cboSource].Enabled = True
    Me![cboSource].BackStyle = 1
Else
    Me![cmdGoToFlint].Visible = False
    Me![cboRawMaterialType].RowSource = "SELECT [ChippedStoneLOV_RawMaterialType].[RawMaterialType], [ChippedStoneLOV_RawMaterialType].[Source] FROM ChippedStoneLOV_RawMaterialType WHERE [RawMaterialType] BETWEEN 1 AND 22 ORDER BY [ChippedStoneLOV_RawMaterialType].[RawMaterialType]; "
    Me![cboRawMaterialType].LimitToList = True
    Me![cboSource].Locked = True
    Me![cboSource].Enabled = False
    Me![cboSource].BackStyle = 0
End If

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub Form_Open(Cancel As Integer)
'new 2011 - safer to take focus to find combo
On Error GoTo err_open

    ''DoCmd.GoToControl Me![cboFind].Name

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Letter_AfterUpdate()
Call UpdateGID

'Dim retVal
'If Me![Letter] <> "" Then
'    If Me![Letter].OldValue <> "" And Me![Unit] <> "" And Me![Number] <> "" Then
'        retVal = MsgBox("Altering the Letter effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
'        If retVal = vbYes Then
'            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'        Else
'            Me![Letter] = Me![Letter].OldValue
'        End If
'    Else
'        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'    End If
'
'End If
End Sub

Private Sub Letter_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_Letter_NotInList

Dim retVal, sql

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [ChippedStoneLOV_Letter]([GIDLetter]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
    'DoCmd.RunCommand acCmdSaveRecord
    'Me![Letter].Requery
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_Letter_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Lip_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_LipNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Lip].Undo
Exit Sub

err_LipNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Number_AfterUpdate()
On Error GoTo err_num

Call UpdateGID

'if all gid there triger a save on the record - 2010
If Me![Unit] <> "" And Me![Letter] <> "" And Me![Number] <> "" Then
    DoCmd.RunCommand acCmdSaveRecord
    Me!cboFind.Requery
End If


'Dim retVal
'If Me![Number] <> "" Then
'    If Me![Number].OldValue <> "" And Me![Unit] <> "" And Me![Letter] <> "" Then
'        retVal = MsgBox("Altering the Number effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
'        If retVal = vbYes Then
'            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'        Else
'            Me![Number] = Me![Number].OldValue
'        End If
'    Else
'        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'    End If
'
'End If
Exit Sub

err_num:
    'new 2011
    If Err.Number = 3146 Then
        MsgBox "This GID already exists, please enter another", vbExclamation, "Duplicate GID"
        Me![Number].Undo
        DoCmd.GoToControl Me!Number.Name
        DoCmd.GoToControl Me!RawMaterial.Name
        
        Me!Number.SetFocus
        
        
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub





Private Sub PortionRepresented_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cboport

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![PortionRepresented].Undo
Exit Sub

err_cboport:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub RawMaterial_AfterUpdate()
'new 2010 to make go to flint button visible for Adam
On Error GoTo err_rawmat

    If Me![RawMaterial] = "Flint" Then
        Me![cmdGoToFlint].Visible = True
        Me![cboRawMaterialType].RowSource = "SELECT [ChippedStoneLOV_RawMaterialType].[RawMaterialType], [ChippedStoneLOV_RawMaterialType].[Source] FROM ChippedStoneLOV_RawMaterialType ORDER BY [ChippedStoneLOV_RawMaterialType].[RawMaterialType];"
        Me![cboRawMaterialType].LimitToList = False
        Me![cboSource].Locked = False
        Me![cboSource].Enabled = True
        Me![cboSource].BackStyle = 1
    Else
        Me![cmdGoToFlint].Visible = False
        Me![cboRawMaterialType].RowSource = "SELECT [ChippedStoneLOV_RawMaterialType].[RawMaterialType], [ChippedStoneLOV_RawMaterialType].[Source] FROM ChippedStoneLOV_RawMaterialType WHERE [RawMaterialType] BETWEEN 1 AND 22 ORDER BY [ChippedStoneLOV_RawMaterialType].[RawMaterialType]; "
        Me![cboRawMaterialType].LimitToList = True
        Me![cboSource].Locked = True
        Me![cboSource].Enabled = False
        Me![cboSource].BackStyle = 0
    End If

Exit Sub

err_rawmat:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub RawMaterial_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
''On Error GoTo err_RawMat_NotInList
''
''Dim retVal, sql, inputname
''
''retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
''If retVal = vbYes Then
''    Response = acDataErrAdded
''    sql = "INSERT INTO [ChippedStoneLOV_RawMaterials]([Material]) VALUES ('" & NewData & "');"
''    DoCmd.RunSQL sql
''Else
''    Response = acDataErrContinue
''End If
''
''
''Exit Sub
''
''err_RawMat_NotInList:
''    Call General_Error_Trap
''    Exit Sub

''ALTERATION IN 2010 - only allow whats in list
'stop not in list msg loop
On Error GoTo err_RawMaterialNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![RawMaterial].Undo
Exit Sub

err_RawMaterialNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub ScarPattern_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_ScarPatternNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![ScarPattern].Undo
Exit Sub

err_ScarPatternNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtBag_AfterUpdate()
'link bag number to unit number from basic data
On Error GoTo err_txtbag
Dim retVal

If Me![Unit] <> "" Then
    'there is already a unit number check it against the unit related to this bag on basic data
    If Me![txtBag].Column(1) <> "" Then
        If (CInt(Me![Unit]) <> CInt(Me![txtBag].Column(1))) Then
            retVal = MsgBox("The Unit shown here (" & Me![Unit] & ") does not match the Unit assigned to this Bag on the Basic Data screen (" & Me![txtBag].Column(1) & "). This operation will overwrite the Unit number " & Me![Unit] & " on this screen, continue anyway?", vbCritical + vbYesNo, "Data Error")
            If retVal = vbYes Then
                Me![Unit] = Me![txtBag].Column(1)
            Else
                Me![txtBag] = Me![txtBag].OldValue
            End If
        End If
    End If
Else
    Me![Unit] = Me![txtBag].Column(1)
End If
Exit Sub

err_txtbag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()
Call UpdateGID

'Dim retVal
'If Me![Unit] <> "" Then
'    If Me![Unit].OldValue <> "" And Me![Letter] <> "" And Me![Number] <> "" Then
'        retVal = MsgBox("Altering the Unit effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
'        If retVal = vbYes Then
'            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'        Else
'            Me![Unit] = Me![Unit].OldValue
'        End If
'    Else
'        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
'    End If
'
'End If
End Sub

Private Sub Use_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_UseNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Use].Undo
Exit Sub

err_UseNot:
    Call General_Error_Trap
    Exit Sub
End Sub
