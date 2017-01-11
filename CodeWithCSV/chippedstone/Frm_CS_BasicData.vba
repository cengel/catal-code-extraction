Option Compare Database
Option Explicit

Private Sub SetUpFields()
'set up display dependant on fields selected

If Me![RetrievalMethod] = "Heavy Residue" Then
    'for HR make all flot and sample fields etc avail
    Me![txtFlotNo].Enabled = True
    Me![txtFlotNo].BackColor = -2147483643
    Me![txtFlotNo].Locked = False
    Me![txtSampleNum].Enabled = True
    Me![txtSampleNum].BackColor = -2147483643
    Me![txtSampleNum].Locked = False
    Me![cboFraction].Enabled = True
    Me![cboFraction].BackColor = -2147483643
    Me![cboFraction].Locked = False
    Me![cboPercent].Enabled = True
    Me![cboPercent].BackColor = -2147483643
    Me![cboPercent].Locked = False
    Me![txtVolume].Enabled = True
    Me![txtVolume].BackColor = -2147483643
    Me![txtVolume].Locked = False
    Me![txtWgt/L].Enabled = True
    Me![txtWgt/L].BackColor = -2147483643
    Me![txtWgt/L].Locked = False
    Me![txtCount/L].Enabled = True
    Me![txtCount/L].BackColor = -2147483643
    Me![txtCount/L].Locked = False
ElseIf Me![RetrievalMethod] = "Fast Track" Then
    'fast track doesn't need wght/l and count/l
    '17/7/06 TC ask also to blank out Flot no, sample no, faction and %
    Me![txtFlotNo].Enabled = False
    Me![txtFlotNo].BackColor = 8421504
    Me![txtFlotNo].Locked = True
    Me![txtSampleNum].Enabled = False
    Me![txtSampleNum].BackColor = 8421504
    Me![txtSampleNum].Locked = True
    Me![cboFraction].Enabled = False
    Me![cboFraction].BackColor = 8421504
    Me![cboFraction].Locked = True
    Me![cboPercent].Enabled = False
    Me![cboPercent].BackColor = 8421504
    Me![cboPercent].Locked = True
    Me![txtVolume].Enabled = True
    Me![txtVolume].BackColor = -2147483643
    Me![txtVolume].Locked = False
    Me![txtWgt/L].Enabled = False
    Me![txtWgt/L].BackColor = 8421504
    Me![txtWgt/L].Locked = True
    Me![txtCount/L].Enabled = False
    Me![txtCount/L].BackColor = 8421504
    Me![txtCount/L].Locked = True
ElseIf Me![RetrievalMethod] = "Dry Sieve" Then
    'only allow vol and count/l weight/l
    Me![txtFlotNo].Enabled = False
    Me![txtFlotNo].BackColor = 8421504
    Me![txtFlotNo].Locked = True
    Me![txtSampleNum].Enabled = False
    Me![txtSampleNum].BackColor = 8421504
    Me![txtSampleNum].Locked = True
    Me![cboFraction].Enabled = False
    Me![cboFraction].BackColor = 8421504
    Me![cboFraction].Locked = True
    Me![cboPercent].Enabled = False
    Me![cboPercent].BackColor = 8421504
    Me![cboPercent].Locked = True
    Me![txtVolume].Enabled = True
    Me![txtVolume].BackColor = -2147483643
    Me![txtVolume].Locked = False
    Me![txtWgt/L].Enabled = True
    Me![txtWgt/L].BackColor = -2147483643
    Me![txtWgt/L].Locked = False
    Me![txtCount/L].Enabled = True
    Me![txtCount/L].BackColor = -2147483643
    Me![txtCount/L].Locked = False
Else
    Me![txtFlotNo].Enabled = False
    Me![txtFlotNo].BackColor = 8421504
    Me![txtFlotNo].Locked = True
    Me![txtSampleNum].Enabled = False
    Me![txtSampleNum].BackColor = 8421504
    Me![txtSampleNum].Locked = True
    Me![cboFraction].Enabled = False
    Me![cboFraction].BackColor = 8421504
    Me![cboFraction].Locked = True
    Me![cboPercent].Enabled = False
    Me![cboPercent].BackColor = 8421504
    Me![cboPercent].Locked = True
    Me![txtVolume].Enabled = False
    Me![txtVolume].BackColor = 8421504
    Me![txtVolume].Locked = True
    Me![txtWgt/L].Enabled = False
    Me![txtWgt/L].BackColor = 8421504
    Me![txtWgt/L].Locked = True
    Me![txtCount/L].Enabled = False
    Me![txtCount/L].BackColor = 8421504
    Me![txtCount/L].Locked = True
End If

End Sub
Private Function CheckValidRecord() As Boolean
'checks if ok to leave the record
On Error GoTo err_check
Dim msg
    If Me![txtBag] = "" Or Me![Unit] = "" Or Me![RawMaterial] = "" Or Me![Count] = "" Or Me![Weight] = "" Or Me![RetrievalMethod] = "" Then
        msg = "Please make sure all the following fields have been filled out:" & Chr(13) & Chr(13)
        msg = msg & "Bag Number" & Chr(13) & "Unit" & Chr(13) & "Raw Material" & Chr(13) & "Count" & Chr(13) & "Weight" & Chr(13) & "Retrieval Method"
        MsgBox msg, vbExclamation, "Incomplete Record"
        CheckValidRecord = False
    Else
        CheckValidRecord = True
     End If
Exit Function

err_check:
    Call General_Error_Trap
    Exit Function

End Function


Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    DoCmd.GoToControl Me![txtBag].Name
    DoCmd.FindRecord Me![cboFind]
    DoCmd.GoToControl Me![Unit].Name
    Me![cboFind] = ""
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboFind_NotInList(NewData As String, Response As Integer)
On Error GoTo err_not
    
    MsgBox "Bag number not found", vbInformation, "Not In List"
    Response = acDataErrContinue
    Me![cboFind].Undo

Exit Sub

err_not:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_cboFindUnit

If Me![cboFindUnit] <> "" Then
    DoCmd.GoToControl Me![Unit].Name
    DoCmd.FindRecord Me![cboFindUnit]
    Me![cboFindUnit] = ""
End If

Exit Sub

err_cboFindUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_NotInList(NewData As String, Response As Integer)
On Error GoTo err_notUnit
    
    MsgBox "Unit number not found", vbInformation, "Not In List"
    Response = acDataErrContinue
    Me![cboFindUnit].Undo

Exit Sub

err_notUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboPercent_AfterUpdate()
On Error GoTo err_cboP

Call CalcCountL(Me)
Call CalcWgtL(Me)

Exit Sub

err_cboP:
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





Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![txtBag] <> "" Then
        DoCmd.OpenForm "Frm_Pop_DataOutputOptions", acNormal, , , , acDialog, Me![txtBag] & ";basic"
    Else
        MsgBox "The output options form cannot be shown when there is no Bag Number on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdStageTwo_Click()
'check if record exists in stage two table and if not put it there ready for data entry
'saj
'2011 refined as a unit has more than one bag so this check on bag doesn't always work
'plus the user has no idea of the next number so can get a PK error which they dont understand so
'need to make clever. This feature was never used as Stringy insisted on using Excel and the data was only
'imported in 2010
On Error GoTo err_stagetwo

If Me![txtBag] <> "" And Me![Unit] <> "" Then
    Dim stagetwo, sql, LetterCode, findnum, Response
   '' stagetwo = DLookup("[Bag]", "[ChippedStone_StageTwo_Data]", "[Bag] = '" & Me![txtBag] & "'")
   '' If IsNull(stagetwo) Then
   ''     'not there yet
   ''     LetterCode = InputBox("Please enter the letter code of the first piece:", "Letter Code for Piece", "A")
   ''     If LetterCode <> "" Then
   ''         findnum = InputBox("Please enter the number of the first piece:", "Number for Piece", "1")
   ''         If findnum <> "" Then
   ''             sql = "INSERT INTO [ChippedStone_StageTwo_Data] ([Unit], [LetterCode], [FindNumber], [Bag], [GID]) VALUES (" & Me![Unit] & ", '" & LetterCode & "'," & findnum & ",'" & Me![txtBag] & "', '" & Me![Unit] & "." & LetterCode & findnum & "');"
   ''             DoCmd.RunSQL sql
   ''         Else
   ''             MsgBox "Sorry but a Find Number is required to enter a new record", vbExclamation, "Insufficient Data"
   ''             Exit Sub
   ''         End If
   ''     Else
   ''         MsgBox "Sorry but a letter code is required to enter a new record", vbExclamation, "Insufficient Data"
   ''         Exit Sub
   ''     End If
   '' End If
   '' DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[Unit] = " & Me![Unit]
   
   'find out if this unit has any A numbers already
   stagetwo = DLookup("[Unit]", "[ChippedStone_StageTwo_Data]", "[Unit] = " & Me![Unit])
    If IsNull(stagetwo) Then
        'unit not there yet at all therefore A1
        Response = MsgBox("The database does not have any A numbers allocated for this unit so the system will create " & Me![Unit] & ".A1" & Chr(13) & Chr(13) & "Is this OK?", vbYesNo, "New A Number")
        If Response = vbNo Then
            Response = MsgBox("Would you still like to move to Stage 2 and allocate a number there yourself?", vbYesNo, "Continue?")
            If Response = vbNo Then
                Exit Sub
            Else
               ''MsgBox "Use the Add New button on the next screen to enter your record", vbInformation, "Entering your record"
               DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[Unit] = " & Me![Unit], acFormAdd
            End If
        Else
            'user want to continue
            sql = "INSERT INTO [ChippedStone_StageTwo_Data] ([Unit], [LetterCode], [FindNumber], [Bag], [GID]) VALUES (" & Me![Unit] & ", 'A',1,'" & Me![txtBag] & "', '" & Me![Unit] & ".A1');"
            DoCmd.RunSQL sql
            DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[GID] = '" & Me![Unit] & ".A1'"
        End If
    Else
        'unit is in there - check what user wants to do maybe find last number to create new record
        Response = MsgBox("Do you want to add a new A number record or simply view existing records for this record?" & Chr(13) & Chr(13) & "To simply view press Yes", vbQuestion + vbYesNo, "Confirm Action")
        If Response = vbYes Then
            DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[Unit] = " & Me![Unit]
        Else
            'does not want to view, wants to add
            sql = "SELECT ChippedStone_StageTwo_Data.Unit, ChippedStone_StageTwo_Data.LetterCode, Last(ChippedStone_StageTwo_Data.FindNumber) AS LastOfFindNumber "
            sql = sql & "FROM ChippedStone_StageTwo_Data  "
            sql = sql & "GROUP BY ChippedStone_StageTwo_Data.Unit, ChippedStone_StageTwo_Data.LetterCode "
            sql = sql & "HAVING (((ChippedStone_StageTwo_Data.Unit)=" & Me![Unit] & "));"
            
            Dim mydb As DAO.Database, myrs As DAO.Recordset, lastnum
            Set mydb = CurrentDb
            Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
            
                If Not (myrs.BOF And myrs.EOF) Then
                    lastnum = myrs!LastOfFindNumber
                Else
                    'this should not happen as if here there will be a number but JUST IN CASE
                    lastnum = 0
                End If
            myrs.Close
            Set myrs = Nothing
            mydb.Close
            Set mydb = Nothing
            
            lastnum = lastnum + 1
            
            Response = MsgBox("The database will allocate the next A number available for this unit which will be: " & Me![Unit] & ".A" & lastnum & Chr(13) & Chr(13) & "Is this OK?", vbYesNo, "New A Number")
            If Response = vbNo Then
                Response = MsgBox("Would you still like to move to Stage 2 and allocate a number there yourself?", vbYesNo, "Continue?")
                If Response = vbNo Then
                    Exit Sub
                Else
                   ''MsgBox "Use the Add New button on the next screen to enter your record", vbInformation, "Entering your record"
                   DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[Unit] = " & Me![Unit], acFormAdd
                End If
            Else
                'user want to continue
                sql = "INSERT INTO [ChippedStone_StageTwo_Data] ([Unit], [LetterCode], [FindNumber], [Bag], [GID]) VALUES (" & Me![Unit] & ", 'A'," & lastnum & ",'" & Me![txtBag] & "', '" & Me![Unit] & ".A" & lastnum & "');"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "frm_CS_StageTwo", acNormal, , "[GID] = '" & Me![Unit] & ".A" & lastnum & "'"
            End If
            
            
        End If
        
    End If
Else
    MsgBox "Please enter the bag number and the unit number first", vbExclamation, "Insufficient Data"
End If
Exit Sub

err_stagetwo:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Count_AfterUpdate()
On Error GoTo err_count

If Me![RetrievalMethod] = "Heavy Residue" Then
    Call CalcCountL(Me)
ElseIf Me![RetrievalMethod] = "Dry Sieve" Then
    Call CalcCountLDrySeive(Me)
End If
Exit Sub

err_count:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Close()
'If CheckValidRecord = False Then
'    MsgBox "no"
'    DoCmd.CancelEvent
'End If
End Sub

Private Sub Form_Current()
'Set up form display
On Error GoTo err_current

'If Me![RetrievalMethod] = "Heavy Residue" Then
'    'let all flot num, sample num etc fields be avail
'    Me![txtFlotNo].Enabled = True
'    Me![txtFlotNo].BackColor = -2147483643
'    Me![txtSampleNum].Enabled = True
'    Me![txtSampleNum].BackColor = -2147483643
'    Me![cboFraction].Enabled = True
'    Me![cboFraction].BackColor = -2147483643
'    Me![cboPercent].Enabled = True
'    Me![cboPercent].BackColor = -2147483643
'    Me![txtVolume].Enabled = True
'    Me![txtVolume].BackColor = -2147483643
'    Me![txtWgt/L].Enabled = True
'    Me![txtWgt/L].BackColor = -2147483643
'    Me![txtCount/L].Enabled = True
'    Me![txtCount/L].BackColor = -2147483643
'    Me![txtFlotNo].Locked = False
'    Me![cboFraction].Locked = False
'    Me![cboPercent].Locked = False
'    Me![txtVolume].Locked = False
'    Me![txtSampleNum].Locked = False
'    Me![txtWgt/L].Locked = False
'    Me![txtCount/L].Locked = False
'ElseIf Me![RetrievalMethod] = "Fast Track" Then
'    'fast track won't have a weight/l or count/l
'    Me![txtVolume].Enabled = False
'    Me![txtVolume].BackColor = -2147483643
'    Me![txtVolume].Locked = True
'    Me![txtWgt/L].Enabled = False
'    Me![txtWgt/L].BackColor = -2147483643
'    Me![txtWgt/L].Locked = True
'    Me![txtCount/L].Enabled = False
'    Me![txtCount/L].BackColor = -2147483643
'    Me![txtCount/L].Locked = True
'ElseIf Me![RetrievalMethod] = "Dry Sieve" Then
'    'only allow vol and count/l weight/l
'    Me![txtVolume].Enabled = True
'    Me![txtVolume].BackColor = -2147483643
'    Me![txtVolume].Locked = False
'    Me![txtWgt/L].Enabled = True
'    Me![txtWgt/L].BackColor = -2147483643
'    Me![txtWgt/L].Locked = False
'    Me![txtCount/L].Enabled = True
'    Me![txtCount/L].BackColor = -2147483643
'    Me![txtCount/L].Locked = False
'Else
'    Me![txtFlotNo].Enabled = False
'    Me![txtFlotNo].BackColor = 8421504
'    Me![txtSampleNum].Enabled = False
'    Me![txtSampleNum].BackColor = 8421504
'    Me![cboFraction].Enabled = False
'    Me![cboFraction].BackColor = 8421504
'    Me![cboPercent].Enabled = False
'    Me![cboPercent].BackColor = 8421504
'    Me![txtVolume].Enabled = False
'    Me![txtVolume].BackColor = 8421504
'    Me![txtWgt/L].Enabled = False
'    Me![txtWgt/L].BackColor = 8421504
'    Me![txtCount/L].Enabled = False
 '   Me![txtCount/L].BackColor = 8421504
'    Me![txtFlotNo].Locked = True
'    Me![cboFraction].Locked = True
'    Me![cboPercent].Locked = True
'    Me![txtVolume].Locked = True
'    Me![txtSampleNum].Locked = True
'    Me![txtWgt/L].Locked = True
'    Me![txtCount/L].Locked = True
'End If
Call SetUpFields
'check if data in stage two yet and make button text dependant
Dim stagetwo
stagetwo = DLookup("[Bag]", "[ChippedStone_StageTwo_Data]", "[Bag] = '" & Me![txtBag] & "'")
If IsNull(stagetwo) Then
    'not there yet
    Me![cmdStageTwo].Caption = "Move to Level 2"
Else
    'there
    Me![cmdStageTwo].Caption = "View Level 2"
End If

DoCmd.GoToControl "cboFind"

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub Form_Deactivate()
'If CheckValidRecord = False Then
'    MsgBox "no"
'    DoCmd.CancelEvent
'End If
'checks if ok to leave the record
On Error GoTo err_check
Dim msg
'    If Me![txtBag] = "" Or Me![Unit] = "" Or Me![RawMaterial] = "" Or Me![Count] = "" Or Me![Weight] = "" Or Me![RetrievalMethod] = "" Then
'        msg = "Please make sure all the following fields have been filled out:" & Chr(13) & Chr(13)
'        msg = msg & "Bag Number" & Chr(13) & "Unit" & Chr(13) & "Raw Material" & Chr(13) & "Count" & Chr(13) & "Weight" & Chr(13) & "Retrieval Method"
'        MsgBox msg, vbExclamation, "Incomplete Record"
'        'CheckValidRecord = False
'        DoCmd.CancelEvent
'    Else
'        'CheckValidRecord = True
'     End If

''testing 2011
''If Me!frm_subform_location.Form.RecordsetClone.RecordCount = 0 Then
''    msg = "NO NO NO NO deactivate"
''    MsgBox msg
''Else
''    MsgBox "yes yes deactivate"
''End If
Exit Sub

err_check:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
'try to give a user friendly message to the problem
If DataErr = 3146 Then
    'probably primary key error
    MsgBox "An error has been encountered. Check you have not entered an existing Bag Number by looking at the pull down list. If this is the case you will needto press ESC but will lose your data (sorry!)", vbCritical, "Error"
    Response = acDataErrContinue
    
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If CheckValidRecord = False Then
'    MsgBox "no"
'    DoCmd.CancelEvent
'End If
'checks if ok to leave the record

'2011 - Stringy has left and the remaining team members are no entering data but viewing and requested I take
'this off as it makes navigation from frustrating. I'm also not sure its working quite right.
On Error GoTo err_check
Dim msg
  ''  If IsNull(Me![txtBag]) Or IsNull(Me![Unit]) Or IsNull(Me![RawMaterial]) Or IsNull(Me![Count]) Or IsNull(Me![Weight]) Or IsNull(Me![RetrievalMethod]) Then
  ''      msg = "Please make sure all the following fields have been filled out:" & Chr(13) & Chr(13)
  ''      msg = msg & "Bag Number" & Chr(13) & "Unit" & Chr(13) & "Raw Material" & Chr(13) & "Count" & Chr(13) & "Weight" & Chr(13) & "Retrieval Method"
  ''      MsgBox msg, vbExclamation, "Incomplete Record"
  ''      'CheckValidRecord = False
  ''      DoCmd.CancelEvent
  ''  Else
  ''      'CheckValidRecord = True
  ''   End If
  
''testing in 2011
''If Me!frm_subform_location.Form.RecordsetClone.RecordCount = 0 Then
''    msg = "NO NO NO NO unload"
''    MsgBox msg
''Else
''    MsgBox "yes yes unload"
''End If
Exit Sub

err_check:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Letter_AfterUpdate()
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

Private Sub Number_AfterUpdate()
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
End Sub





Private Sub RawMaterial_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_RawMat_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [ChippedStoneLOV_RawMaterials]([Material]) VALUES ('" & NewData & "');"
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
'set up fields depending on method selected
On Error GoTo err_retMethod

Dim retVal
If Me![RetrievalMethod].OldValue = "Heavy Residue" Then
    If Not IsNull(Me![txtFlotNo]) Or Me![cboFraction] <> "" Or Not IsNull(Me![txtVolume]) Or Not IsNull(Me![txtSampleNum]) Or Not IsNull(Me![txtWgt/L]) Or Not IsNull(Me![txtCount/L]) Then
        retVal = MsgBox("Changing the Retrieval Method from Heavy Residue will mean you will lose all the Flotation data. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtFlotNo] = Null
            Me![cboFraction] = Null
            Me![txtVolume] = Null
            Me![txtSampleNum] = Null
            Me![txtWgt/L] = Null
            Me![txtCount/L] = Null
            Me![cboPercent] = Null
        Else
            Me![RetrievalMethod] = "Heavy Residue"
            Exit Sub
        End If
    End If
ElseIf Me![RetrievalMethod].OldValue = "Dry Sieve" Then
    If IsNull(Me![txtVolume]) Or Not IsNull(Me![txtWgt/L]) Or Not IsNull(Me![txtCount/L]) Then
        retVal = MsgBox("Changing the Retrieval Method from Dry Sieve will mean you will lose the Volume data. Are you sure?", vbQuestion + vbYesNo, "Confirm Action")
        If retVal = vbYes Then
            Me![txtFlotNo] = Null
            Me![cboFraction] = Null
            Me![txtVolume] = Null
            Me![txtSampleNum] = Null
            Me![txtWgt/L] = Null
            Me![txtCount/L] = Null
            Me![cboPercent] = Null
        Else
            Me![RetrievalMethod] = "Dry Sieve"
            Exit Sub
        End If
    End If
End If

Call SetUpFields

'additionally if the method is dry sieve then get volume from unit sheet
If Me![RetrievalMethod] = "Dry Sieve" Then
    Dim getVol
    getVol = DLookup("[Dry sieve volume]", "[Exca: Unit Sheet with relationships]", "[Unit Number] = " & Me![Unit])
    If Not IsNull(getVol) Then
        Me![txtVolume] = getVol
        Call CalcWgtLDrySeive(Me)
        Call CalcCountLDrySeive(Me)
    Else
        MsgBox "Unable to obtain the Dry Sieve Volume from the Unit Sheet, it might not have been entered", vbInformation, "Volume Missing"
    End If
End If

Exit Sub

err_retMethod:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub RetrievalMethod_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_RetrievalMethod_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [ChippedStone_RetrievalMethod]([RetrievalMethod]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_RetrievalMethod_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub txtBag_AfterUpdate()
'check existence of new bag number
On Error GoTo err_txtbag

    If IsNull(Me![txtBag].OldValue) Then
        Dim checknum, bg
        checknum = DLookup("[BagNo]", "[ChippedStone_Basic_Data]", "[BagNo] = '" & Me![txtBag] & "'")
        If Not IsNull(checknum) Then
            'exists
            MsgBox "Sorry this bag number exists already, the system will take you to the record", vbInformation, "Duplicate Bag Number"
            bg = Me![txtBag]
            'Me![txtBag] = ""
            Me.Undo
            DoCmd.GoToControl Me![txtBag].Name
            DoCmd.FindRecord bg
            DoCmd.GoToControl Me![Unit].Name
        End If
    End If

Exit Sub

err_txtbag:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtFlotNo_AfterUpdate()
'check if flot number exists against flot log, get volume
On Error GoTo err_flotnu
    
    Dim checknum
    If Me![txtFlotNo] <> "" Then
        checknum = DLookup("[Soil Volume]", "[view_ArchaeoBotany_Flot_Log]", "[Flot Number] = " & Me![txtFlotNo])
        If IsNull(checknum) Then
            MsgBox "Please note this Flot Number does not exist in the Flot log yet, please double check it.", vbExclamation, "Check Entry"
        
        Else
            'get the volume
            Me![txtVolume] = checknum
            Call CalcCountL(Me)
            Call CalcWgtL(Me)
        End If
    End If
Exit Sub

err_flotnu:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtVolume_AfterUpdate()
On Error GoTo err_txtVol

If Me![RetrievalMethod] = "Heavy Residue" Then
    Call CalcCountL(Me)
    Call CalcWgtL(Me)
ElseIf Me![RetrievalMethod] = "Dry Sieve" Then
    Call CalcCountLDrySeive(Me)
    Call CalcWgtLDrySeive(Me)
End If
Exit Sub

err_txtVol:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()

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
Me.Refresh
End Sub



Private Sub cmdUnitDesc_Click()
On Error GoTo Err_cmdUnitDesc_Click

If Me![Unit] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[ChippedStone_UnitDescription]", "[Unit] = " & Me![Unit])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [ChippedStone_UnitDescription] ([Unit]) VALUES (" & Me![Unit] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "Frm_CS_UnitDescription", acNormal, , "[Unit] = " & Me![Unit], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Weight_AfterUpdate()
'see if wgt/l can be calculated
On Error GoTo err_weight

If Me![RetrievalMethod] = "Heavy Residue" Then
    Call CalcWgtL(Me)
ElseIf Me![RetrievalMethod] = "Dry Sieve" Then
    Call CalcWgtLDrySeive(Me)
End If

Exit Sub

err_weight:
    Call General_Error_Trap
    Exit Sub

End Sub
Private Sub cmdUnitFilter_Click()
'17/7/06 - TC request unit filter so easily see bags for the unit
On Error GoTo Err_cmdUnitFilter_Click


    If Me![Unit] <> "" Then
        Me.Filter = "[Unit] = " & Me![Unit]
        Me.FilterOn = True
        Me![cmdFilterOff].Enabled = True
        DoCmd.GoToControl "cmdFilterOff"
        Me![cmdUnitFilter].Enabled = False
    Else
        MsgBox "No Unit number to Filter on", vbInformation, "No Unit Number"
        Me.FilterOn = False
        Me![cmdFilterOff].Enabled = False
        Me![cmdUnitFilter].Enabled = True
    End If


    Exit Sub

Err_cmdUnitFilter_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub
Private Sub cmdFilterOff_Click()
'remove unit filter - 17/6/06 part of TC filter for a unit request
On Error GoTo Err_cmdFilterOff_Click

Dim bagshown
    bagshown = Me![txtBag]
    Me![cmdUnitFilter].Enabled = True
    Me.FilterOn = False
    Me.Filter = ""
    DoCmd.GoToControl "txtBag"
    DoCmd.FindRecord bagshown
    Me![cmdFilterOff].Enabled = False

    Exit Sub

Err_cmdFilterOff_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub
