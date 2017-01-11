Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    DoCmd.GoToControl Me![txtGID].Name
    DoCmd.FindRecord Me![cboFind]
    DoCmd.GoToControl Me![Unit].Name
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboStoragePlace_AfterUpdate()
'the storage place determines the whether the museum no or crate num field appears
On Error GoTo err_cboStorage


Dim retVal
If Me![cboStoragePlace].OldValue = 1 And Me![txtCrate] <> "" Then
    retVal = MsgBox("Changing the Storage Location will mean you lose the Crate Number information as you will have to enter a Museum number instead, are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Action")
    If retVal = vbNo Then
        Me![cboStoragePlace] = Me![cboStoragePlace].OldValue
        Exit Sub
    Else
        Me![txtCrate] = Null
    End If
ElseIf Me![cboStoragePlace].OldValue = 2 And Me![txtMuseumNo] <> "" Then
    retVal = MsgBox("Changing the Storage Location will mean you lose the Museum Number information as you will have to enter a Crate number instead, are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Action")
    If retVal = vbNo Then
        Me![cboStoragePlace] = Me![cboStoragePlace].OldValue
        Exit Sub
    Else
        Me![txtMuseumNo] = Null
    End If
End If

If Me![cboStoragePlace] = 1 Then
    Me![lblCrate].Visible = True
    Me![txtCrate].Visible = True
    Me![lblMuseumNo].Visible = False
    Me![txtMuseumNo].Visible = False
ElseIf Me![cboStoragePlace] = 2 Then
    Me![lblCrate].Visible = False
    Me![txtCrate].Visible = False
    Me![lblMuseumNo].Visible = True
    Me![txtMuseumNo].Visible = True
Else
    Me![lblCrate].Visible = False
    Me![txtCrate].Visible = False
    Me![lblMuseumNo].Visible = False
    Me![txtMuseumNo].Visible = False
End If

Exit Sub

err_cboStorage:
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
    DoCmd.GoToControl "Unit"
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

Private Sub cmdGoToPub_Click()
'open the publication form
On Error GoTo err_gotopub

    If Me![Published] = True Then
        DoCmd.OpenForm "Frm_GS_Publications", acNormal, , "[GID] = '" & Me![txtGID] & "'"
    
    Else
        MsgBox "This record is not recorded as published (the check box next to the button). The publication record cannot be shown", vbInformation, "No publication record"
    End If
    

Exit Sub

err_gotopub:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGotoSample_Click()
'open the sample form
On Error GoTo err_gotosample

    If Me![Sampled?] = True Then
        DoCmd.OpenForm "Frm_GS_Samples", acNormal, , "[GID] = '" & Me![txtGID] & "'"
    
    Else
        MsgBox "This record is not recorded as sampled (the check box next to the button). The sample record cannot be shown", vbInformation, "No sample record"
    End If
    

Exit Sub

err_gotosample:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![txtGID] <> "" Then
        DoCmd.OpenForm "Frm_Pop_DataOutputOptions", acNormal, , , , acDialog, Me![txtGID] & ";" & Me![Worked?]
    Else
        MsgBox "The output options form cannot be shown when there is no GID on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'Set up form display
On Error GoTo err_current

If Me![RetrievalMethod] = "Heavy Residue" Then
    Me![txtFlotNo].Enabled = True
    Me![txtFlotNo].BackColor = -2147483643
    Me![cboFraction].Enabled = True
    Me![cboFraction].BackColor = -2147483643
    Me![cboPercent].Enabled = True
    Me![cboPercent].BackColor = -2147483643
    Me![txtVolume].Enabled = True
    Me![txtVolume].BackColor = -2147483643
    Me![txtFlotNo].Locked = False
    Me![cboFraction].Locked = False
    Me![cboPercent].Locked = False
    Me![txtVolume].Locked = False
Else
    Me![txtFlotNo].Enabled = False
    Me![txtFlotNo].BackColor = 9503284
    Me![cboFraction].Enabled = False
    Me![cboFraction].BackColor = 9503284
    Me![cboPercent].Enabled = False
    Me![cboPercent].BackColor = 9503284
    Me![txtVolume].Enabled = False
    Me![txtVolume].BackColor = 9503284
    Me![txtFlotNo].Locked = True
    Me![cboFraction].Locked = True
    Me![cboPercent].Locked = True
    Me![txtVolume].Locked = True
End If

If Me![cboStoragePlace] = 1 Then
    Me![lblCrate].Visible = True
    Me![txtCrate].Visible = True
    Me![lblMuseumNo].Visible = False
    Me![txtMuseumNo].Visible = False
ElseIf Me![cboStoragePlace] = 2 Then
    Me![lblCrate].Visible = False
    Me![txtCrate].Visible = False
    Me![lblMuseumNo].Visible = True
    Me![txtMuseumNo].Visible = True
Else
    Me![lblCrate].Visible = False
    Me![txtCrate].Visible = False
    Me![lblMuseumNo].Visible = False
    Me![txtMuseumNo].Visible = False
End If

If Me![Worked?] = True Then
    Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_WorkedStone"
    
    If Me![subfrmWorkedOrUnworked].Form![Artefact Class] = "Unidentifiable" Then
        Me![subfrmWorkedOrUnworked].Height = "4900"
    Else
        Me![subfrmWorkedOrUnworked].Height = "9800"
    End If
    Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as Worked'"
Else
    Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_UnworkedStone"
    Me![subfrmWorkedOrUnworked].Height = "3444"
    Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as UnWorked'"
End If

If Me![Sampled?] = True Then
    Me![cmdGotoSample].Enabled = True
Else
    Me![cmdGotoSample].Enabled = False
End If

If Me![Published] = True Then
    Me![cmdGoToPub].Enabled = True
Else
    Me![cmdGoToPub].Enabled = False
End If

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub go_to_first_Click()

End Sub

Private Sub GSAnalyst_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_GSAnalyst_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    inputname = InputBox("Please enter the full name of the analyst to go along with these initials:", "Analyst Name")
    If inputname <> "" Then
        sql = "INSERT INTO [GroundStone List of Values: GSAnalyst]([GSAnalystInitials], [GSAnalystName]) VALUES ('" & NewData & "', '" & inputname & "');"
        DoCmd.RunSQL sql
    Else
        Response = acDataErrContinue
    End If
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_GSAnalyst_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Letter_AfterUpdate()
Dim retVal
If Me![Letter] <> "" Then
    If Me![Letter].OldValue <> "" And Me![Unit] <> "" And Me![Number] <> "" Then
        retVal = MsgBox("Altering the Letter effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
        Else
            Me![Letter] = Me![Letter].OldValue
        End If
    Else
        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
    End If

End If
End Sub

Private Sub Letter_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_Letter_NotInList

Dim retVal, sql

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [GroundStone List of Values: GSLetter]([GIDLetter]) VALUES ('" & NewData & "');"
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

Private Sub Number_AfterUpdate()
Dim retVal
If Me![Number] <> "" Then
    If Me![Number].OldValue <> "" And Me![Unit] <> "" And Me![Letter] <> "" Then
        retVal = MsgBox("Altering the Number effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
        Else
            Me![Number] = Me![Number].OldValue
        End If
    Else
        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
    End If

End If
End Sub

Private Sub Published_AfterUpdate()
'set up go to button
On Error GoTo err_Published

If Me![Published] = True Then
    
    Dim pubnum, sql
    sql = "INSERT INTO [Groundstone: Publications] ([GID], [Unit], [Letter], [Number]) VALUES ('" & Me![txtGID] & "'," & Me![Unit] & ",'" & Me![Letter] & "'," & Me![Number] & ");"
    DoCmd.RunSQL sql
    Me![cmdGoToPub].Enabled = True
    
Else
    'don't allow the pub to be unchecked if pub details exist for this GID
    Dim checknum
    checknum = DLookup("[GID]", "[Groundstone: Publications]", "[GID] = " & Me![txtGID])
    If IsNull(checknum) Then
        Me![cmdGoToPub].Enabled = False
    Else
        MsgBox "Publication records exist for this GID, sorry but you cannot uncheck this box whilst these exist", vbInformation, "Invalid Action"
        Me![Published] = True
        Me![cmdGoToPub].Enabled = True
    End If
End If
Exit Sub

err_Published:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub RawMaterial_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_RawMat_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [Groundstone List of Values: Raw Materials]([stone type]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_RawMat_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub RetrievalMethod_AfterUpdate()
Dim retVal
If Me![RetrievalMethod].OldValue = "Heavy Residue" Then
    If Not IsNull(Me![txtFlotNo]) Or Me![cboFraction] <> "" Or Not IsNull(Me![txtVolume]) Then
        retVal = MsgBox("Changing the Retrieval Method from Heavy Residue will mean you will lose all the Flotation data. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtFlotNo] = Null
            Me![cboFraction] = Null
            Me![txtVolume] = Null
        Else
            Me![RetrievalMethod] = "Heavy Residue"
            Exit Sub
        End If
    End If
End If

If Me![RetrievalMethod] = "Heavy Residue" Then
    Me![txtFlotNo].Enabled = True
    Me![txtFlotNo].BackColor = -2147483643
    Me![cboFraction].Enabled = True
    Me![cboFraction].BackColor = -2147483643
    Me![cboPercent].Enabled = True
    Me![cboPercent].BackColor = -2147483643
    Me![txtVolume].Enabled = True
    Me![txtVolume].BackColor = -2147483643
    Me![txtFlotNo].Locked = False
    Me![cboFraction].Locked = False
    Me![cboPercent].Locked = False
    Me![txtVolume].Locked = False
Else
    Me![txtFlotNo].Enabled = False
    Me![txtFlotNo].BackColor = 9503284
    Me![cboFraction].Enabled = False
    Me![cboFraction].BackColor = 9503284
    Me![cboPercent].Enabled = False
    Me![cboPercent].BackColor = 9503284
    Me![txtVolume].Enabled = False
    Me![txtVolume].BackColor = 9503284
    Me![txtFlotNo].Locked = True
    Me![cboFraction].Locked = True
    Me![cboPercent].Locked = True
    Me![txtVolume].Locked = True
End If

End Sub

Private Sub RetrievalMethod_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_RetrievalMethod_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [Groundstone List of Values: RetrievalMethod]([RetrievalMethod]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_RetrievalMethod_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Sampled__AfterUpdate()
'set up go to button
On Error GoTo err_sampled

If Me![Sampled?] = True Then
    
    Dim SampleNum, sql
    SampleNum = InputBox("Please enter the unique sample number that you have allocated to this sample:", "GS Sample Number")
    If SampleNum = "" Then
        MsgBox "You must enter a unique groundstone sample number to identify this sample before proceeding", vbInformation, "Action Cancelled"
        Me![Sampled?] = False
        Me![cmdGotoSample].Enabled = False
    Else
        sql = "INSERT INTO [Groundstone 6: Samples] ([GID], [Unit], [Letter], [Number], [GSSample Number]) VALUES ('" & Me![txtGID] & "'," & Me![Unit] & ",'" & Me![Letter] & "'," & Me![Number] & ",'" & SampleNum & "');"
        DoCmd.RunSQL sql
        Me![cmdGotoSample].Enabled = True
    End If
Else
    'don't allow the sample to be unchecked if sample details exist for this GID
    Dim checknum
    checknum = DLookup("[GID]", "[Groundstone 6: Samples]", "[GID] = " & Me![txtGID])
    If IsNull(checknum) Then
        Me![cmdGotoSample].Enabled = False
    Else
        MsgBox "Samples exist for this GID, sorry but you cannot uncheck this box whilst these exist", vbInformation, "Invalid Action"
        Me![Sampled?] = True
        Me![cmdGotoSample].Enabled = True
    End If
End If
Exit Sub

err_sampled:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()

Dim retVal
If Me![Unit] <> "" Then
    If Me![Unit].OldValue <> "" And Me![Letter] <> "" And Me![Number] <> "" Then
        retVal = MsgBox("Altering the Unit effects the GID, are you sure you want to do this?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
        Else
            Me![Unit] = Me![Unit].OldValue
        End If
    Else
        Me![txtGID] = Me![Unit] & "." & Me![Letter] & Me![Number]
    End If

End If
End Sub

Private Sub Worked__AfterUpdate()
Dim retVal, checknum, sql
If IsNull(Me![subfrmWorkedOrUnworked]![GID]) Then
    'this record has no worked/unworked specific data so can allow alteration without a problem
    If Me![Worked?] = True Then
        Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_WorkedStone"
        Me![subfrmWorkedOrUnworked].Height = "6885"
        Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as Worked'"
        Me![subfrmWorkedOrUnworked].Form![txtGID] = Me![txtGID]
    Else
        Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_UnworkedStone"
        Me![subfrmWorkedOrUnworked].Height = "3444"
        Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as UnWorked'"
        Me![subfrmWorkedOrUnworked].Form![txtGID] = Me![txtGID]
    End If
Else
    If Me![Worked?] = False Then
        'was worked
        checknum = DLookup("[GID]", "[GroundStone 3: Worked Stone Basics]", "[GID] = '" & Me![txtGID] & "'")
        If Not IsNull(checknum) Then
            retVal = MsgBox("This action means you will lose all of the information entered into the Worked fields below. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![Worked?] = True
                Exit Sub
            Else
                sql = "DELETE FROM [GroundStone 3: Worked Stone Basics] WHERE [GID] = '" & Me![txtGID] & "';"
                DoCmd.RunSQL sql
            End If
        End If
        Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_UnworkedStone"
        Me![subfrmWorkedOrUnworked].Height = "3444"
        Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as UnWorked'"
        Me![subfrmWorkedOrUnworked].Form![txtGID] = Me![txtGID]
    Else
        'was unworked
        checknum = DLookup("[GID]", "[GroundStone 2: UnWorked Stone Basics]", "[GID] = '" & Me![txtGID] & "'")
        If Not IsNull(checknum) Then
            retVal = MsgBox("This action means you will lose all of the information entered into the UnWorked fields below. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![Worked?] = False
                Exit Sub
            Else
                sql = "DELETE FROM [GroundStone 2: UnWorked Stone Basics] WHERE [GID] = '" & Me![txtGID] & "';"
                DoCmd.RunSQL sql
            End If
        End If
        Me![subfrmWorkedOrUnworked].SourceObject = "Frm_subform_WorkedStone"
        Me![subfrmWorkedOrUnworked].Height = "6885"
        Me![txtWorkUnWorlLBL].ControlSource = "='This record is currently marked as Worked'"
        Me![subfrmWorkedOrUnworked].Form![txtGID] = Me![txtGID]
    End If
End If
End Sub

Private Sub cmdUnitDesc_Click()
On Error GoTo Err_cmdUnitDesc_Click

If Me![Unit] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[Groundstone: GS Unit Description]", "[Unit] = " & Me![Unit])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [Groundstone: GS Unit Description] ([Unit]) VALUES (" & Me![Unit] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "Frm_GS_UnitDescription", acNormal, , "[Unit] = " & Me![Unit], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub
