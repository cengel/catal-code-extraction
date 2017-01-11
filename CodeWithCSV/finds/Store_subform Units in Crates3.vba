Option Compare Database
Option Explicit
Private Sub KnownFind()
'new season 2006 if this find is known to the basic data table then display its material from there
On Error GoTo err_knownfind

If Me![Unit] <> "" And Me![FindSampleLetter] <> "" And Me![FindNumber] <> "" Then
    Dim getmaterial, getmaterialsub, getobject, GID
    GID = Me![Unit] & "." & Me![FindSampleLetter] & Me![FindNumber]
    
    getmaterial = DLookup("[MaterialGroupID]", "[Q_Basic_Data_Material_and_Type_with_Text]", "[GID] = '" & GID & "'")
    If Not IsNull(getmaterial) Then Me![cboMaterialGroup] = getmaterial
    
    getmaterialsub = DLookup("[MaterialSubGroupID]", "[Q_Basic_Data_Material_and_Type_with_Text]", "[GID] = '" & GID & "'")
    If Not IsNull(getmaterialsub) Then Me![cboMaterialSubgroup] = getmaterialsub

    getobject = DLookup("[ObjectTypeID]", "[Q_Basic_Data_Material_and_Type_with_Text]", "[GID] = '" & GID & "'")
    If Not IsNull(getobject) Then Me![cboDescription] = getobject

End If
Exit Sub

err_knownfind:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cboDescription_GotFocus()
'instead of setting the rowsource in properties set it here to ensure
'previous records fields stay visible
On Error GoTo err_cboDescFocus

'Dim sql
'sql = " SELECT Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID, Finds_Code_ObjectTypes.ObjectTypeText, "
'sql = sql & "Finds_Code_MaterialSubGroup_ObjectType.MaterialSubGroupID FROM "
'sql = sql & "Finds_Code_MaterialSubGroup_ObjectType INNER JOIN Finds_Code_ObjectTypes ON "
'sql = sql & "Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID = Finds_Code_ObjectTypes.ObjectTypeID WHERE "
'sql = sql & "(((Finds_Code_MaterialSubGroup_ObjectType.MaterialSubGroupID)=" & Forms![Store: Crate Register]![Store: subform Units in Crates].Form![cboMaterialSubGroup] & "));"
'Me![cboDescription].RowSource = sql

Exit Sub

err_cboDescFocus:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboDescription_NotInList(NewData As String, Response As Integer)
'new 2009 flag this is new to list but allow
On Error GoTo err_new

    'If GetGeneralPermissions = "Admin" Then 'if only admins can add reinstate
        Response = acDataErrContinue
        Dim retVal
        retVal = MsgBox("This description entry is new to the list, are you sure?", vbQuestion + vbYesNo, "Confirm Entry")
        If retVal = vbYes Then
            Me![cboDescription].LimitToList = False
            Me![cboDescription] = NewData
            Me![cboDescription].LimitToList = True
            DoCmd.GoToControl "Year"
            Me![cboDescription].Requery
        Else
            Response = acDataErrContinue
            Me![cboDescription].Undo
        End If
    'End If

Exit Sub

err_new:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMaterialGroup_AfterUpdate()
On Error GoTo err_cboMat
'replaced by got focus code of material subgroup
'Me![cboMaterialSubGroup].Requery

Exit Sub

err_cboMat:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cboMaterialSubgroup_AfterUpdate()
On Error GoTo err_cboMaterialSubgroup
'replaced by got focus code of description
'Me![cboDescription].Requery

Exit Sub

err_cboMaterialSubgroup:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMaterialSubGroup_GotFocus()
'instead of setting the rowsource in properties set it here to ensure
'previous records fields stay visible
On Error GoTo err_cboMatSubGrp

Me![cboMaterialSubgroup].RowSource = "SELECT Finds_Code_MaterialGroup_Subgroup.MaterialSubGroupID, Finds_Code_MaterialGroup_Subgroup.MaterialSubgroupText, Finds_Code_MaterialGroup_Subgroup.MaterialGroupID FROM Finds_Code_MaterialGroup_Subgroup WHERE (((Finds_Code_MaterialGroup_Subgroup.MaterialGroupID)=" & Forms![Store: Crate Register]![Store: subform Units in Crates].Form![cboMaterialGroup] & "));"

Exit Sub

err_cboMatSubGrp:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMaterialGroup_NotInList(NewData As String, Response As Integer)
'new 2008 flag this is new to list but allow
On Error GoTo err_new

    'If GetGeneralPermissions = "Admin" Then 'if only admins can add reinstate
        Response = acDataErrContinue
        Dim retVal
        retVal = MsgBox("This material entry is new to the list, are you sure?", vbQuestion + vbYesNo, "Confirm Entry")
        If retVal = vbYes Then
            Me![cboMaterialGroup].LimitToList = False
            Me![cboMaterialGroup] = NewData
            Me![cboMaterialGroup].LimitToList = True
            DoCmd.GoToControl "cboDescription"
            Me![cboMaterialGroup].Requery
        Else
            Response = acDataErrContinue
            Me![cboMaterialGroup].Undo
        End If
    'End If

Exit Sub

err_new:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMoveCrate_AfterUpdate()
'new in season 2006 - move an item to a new crate - SAJ
On Error GoTo err_cboMove
    'bad sarah, lazy programming this code is repeated in store: find unit in crate2 - centralised this when time, just leaving 2006
    If Me![cboMoveCrate] <> "" Then
        'the move will need to go into a transaction
        ''2011 reviewing this is seems overly complicated, why not simply change the crate letter/num fields?
        ''commented out with ''
        ''On Error Resume Next
        ''Dim mydb As DAO.Database, wrkdefault As Workspace, sql1, sql2, myq As QueryDef
        ''Set wrkdefault = DBEngine.Workspaces(0)
        ''Set mydb = CurrentDb
        ''
        ''' Start of outer transaction.
        ''wrkdefault.BeginTrans
        ''
        '''insert into new crate
        '''altered to match new table structure 2009
        ''sql1 = "INSERT INTO [Store: Units in Crates] "
        ''sql1 = sql1 & "( [Unit number], Bag, Material, "
        ''sql1 = sql1 & "Description, Notes, [Year], Area, Studied, "
        ''sql1 = sql1 & "CrateNumber, CrateLetter, FindSampleLetter, FindNumber, SampleNumber, "
        ''sql1 = sql1 & "FlotNumber, EtutlukNumber, EnvanterNumber, MuseumAccessionNumber, "
        ''sql1 = sql1 & "ExportLocation ) "
        ''sql1 = sql1 & "SELECT "
        ''sql1 = sql1 & "[Store: Units in Crates].[Unit number], [Store: Units in Crates].Bag, "
        ''sql1 = sql1 & "[Store: Units in Crates].Material, "
        ''sql1 = sql1 & "[Store: Units in Crates].Description, "
        ''sql1 = sql1 & "[Store: Units in Crates].Notes, [Store: Units in Crates].Year, "
        ''sql1 = sql1 & "[Store: Units in Crates].Area, [Store: Units in Crates].Studied, "
        ''sql1 = sql1 & Me![cboMoveCrate].Column(2) & " AS 'CrateNumber', '" & Me![cboMoveCrate].Column(1) & "' AS 'CrateLetter', "
        ''sql1 = sql1 & "[Store: Units in Crates].FindSampleLetter, [Store: Units in Crates].FindNumber, "
        ''sql1 = sql1 & "[Store: Units in Crates].SampleNumber, [Store: Units in Crates].FlotNumber, "
        ''sql1 = sql1 & "[Store: Units in Crates].EtutlukNumber, [Store: Units in Crates].EnvanterNumber, "
        ''sql1 = sql1 & "[Store: Units in Crates].MuseumAccessionNumber, [Store: Units in Crates].ExportLocation "
        ''sql1 = sql1 & "FROM [Store: Units in Crates] "
        ''sql1 = sql1 & "WHERE [Store: Units in Crates].rowID = " & Me![rowID] & ";"
        ''
        ''Set myq = mydb.CreateQueryDef("")
        ''myq.sql = sql1
        ''myq.Execute
        ''
        ''myq.close
        ''Set myq = Nothing
        ''
        '''delete from here
        '''sql2 = "DELETE FROM [Store: Units in Crates] WHERE [Store: Units in Crates].rowID = " & Me![rowID] & ";"
        ''If DeleteCrateRecord(Me![rowID], mydb) = False Then
        ''    MsgBox "The delete part of this operation has failed", vbCritical, "Operation Failed"
        ''End If
        ''
        ''If Err.Number = 0 Then
        ''    wrkdefault.CommitTrans
        ''    'MsgBox "Move has been successful"
        ''    Me.Requery
        ''    Me![cboMoveCrate] = ""
        ''    MsgBox "Move has been successful"
        ''Else
        ''    wrkdefault.Rollback
        ''    MsgBox "A problem has occured and the move has been cancelled. The error message is: " & Err.Description
        ''End If
        ''
        ''mydb.close
        ''Set mydb = Nothing
        ''wrkdefault.close
        ''Set wrkdefault = Nothing
        
        '2011 track movement in the tracker table
        '2015 non-admins should be asked before moving into virtual crates - DL
        Dim sql
        Dim LResponse
        Debug.Print Me![CrateNumber] & " " & Me![cboMoveCrate] & " " & GetGeneralPermissions
        If Me![cboMoveCrate].Column(2) = "5000" Or Me![cboMoveCrate].Column(2) = "0" Then
            If GetGeneralPermissions <> "Admin" Then
            LResponse = MsgBox("Your moving this item into a virtual crate?", vbYesNo, "Continue")
                If LResponse <> vbYes Then
                    Me![cboMoveCrate] = ""
                Else
                sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [SampleNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
                     sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, '" & Me![cboMoveCrate].Column(1) & Me![cboMoveCrate].Column(2) & "' as MovedToCrate, '" & logon & "', '" & Now & "' "
                     sql = sql & " FROM [Store: Units in Crates] "
                     sql = sql & " WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                'change the crate number/letter to the one selected
                sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & Me![cboMoveCrate].Column(2) & ", [CrateLetter] = '" & Me![cboMoveCrate].Column(1) & "' WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                Me.Requery
                Me![cboMoveCrate] = ""
                MsgBox "Move has been successful"
            End If
        Else
                sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [SampleNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
                     sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, '" & Me![cboMoveCrate].Column(1) & Me![cboMoveCrate].Column(2) & "' as MovedToCrate, '" & logon & "', '" & Now & "' "
                     sql = sql & " FROM [Store: Units in Crates] "
                     sql = sql & " WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                'change the crate number/letter to the one selected
                sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & Me![cboMoveCrate].Column(2) & ", [CrateLetter] = '" & Me![cboMoveCrate].Column(1) & "' WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                Me.Requery
                Me![cboMoveCrate] = ""
                MsgBox "Move has been successful"
        End If
        Else
                sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [SampleNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
                     sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, '" & Me![cboMoveCrate].Column(1) & Me![cboMoveCrate].Column(2) & "' as MovedToCrate, '" & logon & "', '" & Now & "' "
                     sql = sql & " FROM [Store: Units in Crates] "
                     sql = sql & " WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                'change the crate number/letter to the one selected
                sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & Me![cboMoveCrate].Column(2) & ", [CrateLetter] = '" & Me![cboMoveCrate].Column(1) & "' WHERE [RowID] = " & Me![rowID] & ";"
                DoCmd.RunSQL sql
                Me.Requery
                Me![cboMoveCrate] = ""
                MsgBox "Move has been successful"
    End If
End If
Exit Sub

err_cboMove:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub chkMove_Click()
'reveal move crate button
On Error GoTo err_chkmove

    If Me!chkMove = True Then
        Me![cboMoveCrate].ColumnHidden = False
    Else
        Me![cboMoveCrate].ColumnHidden = True
    End If

Exit Sub

err_chkmove:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboMoveCrate_DblClick(Cancel As Integer)
On Error GoTo err_tracker
'new 2011 - find out where entry was previous located (if at all)

' wrap a condition around all of this to prevent photo and illustration
' from movign items back- 2013 season

If CrateLetterFlag <> "Illustrate" And CrateLetterFlag <> "PHOTO" Then
    'do a check to see if has moved
    Dim checknum
    checknum = DLookup("[OriginalrowID]", "[Store: Crate Movement by Teams]", "[OriginalrowID] = " & Me![rowID])
        If Not IsNull(checknum) Then
            'it has moved before
            DoCmd.OpenForm "frm_pop_movement_history", acNormal, , "[OriginalRowID] = " & Me![rowID], acFormPropertySettings
            'Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
            Me.Requery
            Me.Refresh
        Else
            MsgBox "This record hasn't a tracking history in the database", vbInformation, "No Tracking Info"
        End If
Else
    MsgBox "you cannot move items back"
End If

Exit Sub

err_tracker:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub FindLetter_AfterUpdate()
If Me![FindSampleLetter] <> "" Or Not IsNull(Me![FindSampleLetter]) Then
    If UCase(Me![FindSampleLetter]) <> "S" Then
        Me![FindSampleLetter] = UCase(Me![FindSampleLetter])
    End If
End If
'Call KnownFind
End Sub

Private Sub FindNumber_AfterUpdate()
'Call KnownFind
'2008 - need a way to ensure all X finds go into the basic data table
'v3.1
On Error GoTo err_findnum_upd
    
    If Me![FindSampleLetter] <> "" Then
        If UCase(Me![FindSampleLetter]) = "X" Then
            'check this exists
            Dim LResponse
            Dim strSQL, sql As String
            Dim resp
            Dim qdf
            Dim rst As DAO.Recordset
            
            resp = DLookup("[GID]", "[Finds: Basic Data]", "[GID] = '" & Me![Unit] & ".X" & Me![FindNumber] & "'")
            If IsNull(resp) Then
                'GID missing display message to user. Would be good to add auto but then would need
                'to prompt for material group and subgroup etc and Jules not keen, so will start with this
                'move on cursor
                'Modified to allow to insert basic data in a new x-find
                LResponse = MsgBox("This X Find does not exist in the X Finds Register. Do you want to create it now?", vbYesNo, "Continue")
                If LResponse <> vbYes Then
                    MsgBox "Please ensure you enter it.", , "Data Validation"
                    DoCmd.GoToControl "SampleNumber"

                Else
                    Dim passunit, passnumber As Integer
                    Dim passletter As String
                    
                    sql = "INSERT INTO [Finds: Basic Data] ([GID], [Unit], [FindLetter], [FindNumber]) VALUES ('" & Me![Unit] & "." & Me![FindSampleLetter] & Me![FindNumber] & "', " & Me![Unit] & ", '" & Me![FindSampleLetter] & "', " & Me![FindNumber] & ");"
                    'Debug.Print sql
                    DoCmd.RunSQL sql
                    
                    MsgBox "X-Find " & Me![Unit] & "." & Me![FindSampleLetter] & Me![FindNumber] & " created with basic data." & Chr(13) & "Use [Finds: Basic Data] for entering detail information.", , "find creation"
                    
                    passunit = Me![Unit]
                    passletter = Me![FindSampleLetter]
                    passnumber = Me![FindNumber]
                    
                    Me.Requery
                    DoCmd.OpenForm "Finds: Basic Data", , , "Unit = " & passunit & " AND FindLetter = '" & passletter & "' AND FindNumber = " & passnumber
                    Forms![Finds: Basic Data]![frm_subform_materialstypes].SetFocus
                    'DoCmd.GoToControl (Forms![Finds: Basic Data]![frm_subform_materialstypes].Form![cboMaterialGroup])
                End If
            Else
                'GID already exists - throw error, and inform about actual location
                
                strSQL = "SELECT [CrateLetter], [CrateNumber], [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber] FROM [Store: Units in Crates] WHERE [Unit number]=" & Me![Unit] & " AND [FindSampleLetter]='" & Me![FindSampleLetter] & "' AND [FindNumber]=" & [FindNumber] & " ORDER BY [CrateLetter], [CrateNumber] ASC;"
                'Debug.Print strSQL
                Set rst = CurrentDb.OpenRecordset(strSQL)
                
                If Not (rst.EOF) Then
                    'Debug.Print rst![CrateLetter] & " " & rst![CrateNumber]
                    MsgBox "This X Find does already exist in the X Finds Register. Location: Crate " & rst![CrateLetter] & rst![CrateNumber], , "Double Record"
                    DoCmd.GoToControl "FindNumber"
                Else
                    Debug.Print "X-Find does not exist"
                End If
                
            End If
        End If
    End If
Exit Sub

err_findnum_upd:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub FindNumber_Change()

End Sub

Private Sub Form_AfterUpdate()
'moved from before update - saj season 2006
On Error GoTo err_afterupdate
    Forms![Store: Crate Register]![Date Changed] = Now()
    
'End If

Exit Sub

err_afterupdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo err_Form_BeforeInsert
Me![LastUpdated] = Now()
Exit Sub

err_Form_BeforeInsert:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'again moved to after update
'Forms![Store: Crate Register]![Date Changed] = Now()

End Sub

Private Sub Form_Current()
'new season 2006 - the fields shown are dependant on the type of crate
On Error GoTo err_current

   ' Me![cboMaterialGroup].Requery
   ' Me![cboMaterialSubGroup].Requery
   ' Me![cboDescription].Requery
   
    Me![EtutlukNumber].ColumnHidden = True
    ''Me![EnvanterNumber].Visible = False
    ''Me![MuseumAccessionNumber].Visible = False
    Me![ExportLocation].ColumnHidden = True
    Me![Bag].ColumnHidden = True
    Me![Studied].ColumnHidden = True
    ''Me![txtNotes2].Visible = False
    ''Me![lblNotes2].Visible = False
    ''Me![txtNotes3].Visible = False
    ''Me![lblNotes3].Visible = False
    ''Me![lblAdditional].caption = "Notes"
    ''Me![lblMuseum].Visible = False
    Me![txtNotes].Visible = True
    'mellaart fields intro 2009
    Me![MellaartID].ColumnHidden = True
    Me![MellaartLocation].ColumnHidden = True
    Me![MellaartNotes].ColumnHidden = True
    
   
    If LCase(Forms![Store: Crate Register]![CrateLetter]) = "et" Then
         'etukluk crates must have an ET prefix
        Me![EtutlukNumber].ColumnHidden = False
        Me![Bag].ColumnHidden = False
        ''Me![lblAdditional].caption = "Etukluk No."
        ''Me![txtNotes].Visible = False
        ''Me![txtNotes3].Visible = False
        ''Me![txtNotes2].Visible = True
        ''Me![lblNotes2].Visible = True
    ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "envanter" Then
        'envanter crates must have envanter prefix
        ''Me![EnvanterNumber].Visible = True
        ''Me![MuseumAccessionNumber].Visible = True
        ''Me![lblMuseum].Visible = True
        ''Me![lblAdditional].caption = "Evanter No."
        ''Me![txtNotes].Visible = False
        ''Me![txtNotes3].Visible = False
        ''Me![txtNotes2].Visible = True
        ''Me![lblNotes2].Visible = True
    ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "export" Then
        'export crates must have export prefix
        Me![ExportLocation].ColumnHidden = False
        ''Me![lblAdditional].caption = "Export Location"
        ''Me![txtNotes].Visible = False
        ''Me![txtNotes3].Visible = False
        ''Me![txtNotes2].Visible = True
        ''Me![lblNotes2].Visible = True
    Else
        'all other crates have same fields apart from two
        Me![txtNotes].Visible = True
        
        If LCase(Forms![Store: Crate Register]![CrateLetter]) = "ob" Then
            'bag visible for chipped stone
            Me![Bag].ColumnHidden = False
            ''Me![txtNotes3].Visible = True
            ''Me![lblNotes3].Visible = True
            ''Me![txtNotes].Visible = False
            ''Me![lblAdditional].caption = "Bag"
        ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "fb" Then
            'studied visible for faunal
            Me![Studied].ColumnHidden = False
            ''Me![txtNotes3].Visible = True
            ''Me![lblNotes3].Visible = True
            ''Me![txtNotes].Visible = False
            ''Me![lblAdditional].caption = "Studied"
        ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "mellet" Then
            'bag visible for mellet - new v4.4 2009
            Me![Bag].ColumnHidden = False
            Me![MellaartID].ColumnHidden = False
            Me![MellaartLocation].ColumnHidden = False
            Me![MellaartNotes].ColumnHidden = False
        End If
        
        
        'Me![txtNotes2].Visible = False

        
    End If
    
' Added this to only crates labelled for a particular team be displayed in the
' dropdown that shows the options to move units between crates. For example, Faunal team
' only gets the FB crates to choose from.
' CE - 2012 season originally in Form_Open,
' CE - 2013 season updated and moved here to accommodate for more users
' and more differentiated moves
' CE - 2014 amended

'MsgBox [Forms]![Store: Crate Register]![CrateLetter]

If CrateLetterFlag = "Illustrate" Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "PHOTO" Then
         Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "CONS" Then
    If [Forms]![Store: Crate Register]![CrateLetter] = "CONS" Then 'show all other crates
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE (([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else 'show CONS only
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    End If
' these can only move within their domain
ElseIf CrateLetterFlag = "P" Or CrateLetterFlag = "HB" Or CrateLetterFlag = "OB" Or CrateLetterFlag = "PH" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "CO" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'BE', 'FG', 'CB'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "GS" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'NS', 'Depot'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "FB" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'Depot'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "char" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'or'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "S" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'BE'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "BE" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
ElseIf CrateLetterFlag = "*" Then
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS Crate, [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE (([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
End If


'flotnumber should not be entered manually anymore - adding a query to aggregate all related flotnumbers and pull them into a flotnumber field
'DL 07/02/2015

'Dim strSQL, test
'Dim rst As DAO.Recordset
'Dim fld As Field
                
'If Me![SampleNumber] > 0 And Me![Unit] > 0 Then
                
'    strSQL = "SELECT [Flot Number], [Unit Number], [Sample Number] FROM [Bot: Basic Data] WHERE [Unit number]=" & Me![Unit] & " AND [Sample Number]=" & Me![SampleNumber] & " ORDER BY [Sample Number] ASC;"
'    Debug.Print strSQL
'    Set rst = CurrentDb.OpenRecordset(strSQL)
'    test = [Forms]![Store: Crate Register]![Store: subform Units in Crates]![subform_botflot]![Flot Number].Value
'    Debug.Print "test: " & test
    
'    If Not (rst.EOF) Then
'        Debug.Print rst![Flot Number]
        'MsgBox "This X Find does already exist in the X Finds Register. Location: Crate " & rst![CrateLetter] & rst![CrateNumber], , "Double Record"
'    Else
'        Debug.Print "no sample"
'    End If
'Else
'End If

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Form_Open(Cancel As Integer)
'new 2012 only let admins edit this via this form - taken from
'subform find unit in crates2

On Error GoTo err_open


If GetGeneralPermissions = "Admin" Then
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me![cboMoveCrate].Visible = True
    'Me![Label77].Visible = True
    'Me![Text47].Visible = False
' 2013 added RW group here, because we need them to be able to move crates
ElseIf GetGeneralPermissions = "RW" Then
    Me.AllowDeletions = False
    Me.AllowEdits = True
    Me![cboMoveCrate].Visible = True
    'Me![Label77].Visible = True
    'Me![Text47].Visible = False
Else
    Me.AllowDeletions = False
    Me.AllowEdits = False
    Me![cboMoveCrate].Visible = False
    'Me![Label77].Visible = False
    'Me![Text47].Visible = False
    
End If


'added 2012 to allow team leaders to move bags inbetween crates
'If CrateLetterFlag = "FB" Or CrateLetterFlag = "P" Or CrateLetterFlag = "GS" Or CrateLetterFlag = "OB" Or CrateLetterFlag = "BE" Or CrateLetterFlag = "CONSERVATION" Or CrateLetterFlag = "FG" Then
'    Me![cboMoveCrate].Visible = True
'    Me.AllowEdits = True
'    Me.AllowDeletions = False
'    'lock all the other fields to avoid nasty ODBC update error no 229
'    Me![Unit].Locked = True
'    Me![FindLetter].Locked = True
'    Me![FindNumber].Locked = True
'    Me![SampleNumber].Locked = True
'    Me![FlotNumber].Locked = True
'    Me![cboMaterialGroup].Locked = True
'    Me![cboDescription].Locked = True
'    Me![Year].Locked = True
'    Me![Area].Locked = True
'    Me![txtNotes].Locked = True
'    Me![Bag].Locked = True
'    Me![ExportLocation].Locked = True
'    Me![EtutlukNumber].Locked = True
'    Me![Studied].Locked = True
'    Me![MellaartID].Locked = True
'    Me![MellaartLocation].Locked = True
'    Me![MellaartNotes].Locked = True
'End If


Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_AfterUpdate()
'need to get area and year from excavation but can't link them in as a subform
'as this is a continous form. Can't also set the field value on current as this will
'make all records look the same until you move the focus down the row - instead store
'the year and area in the units in crates table but grab it automatically from the exca
'DB when updated
'saj season 2006
On Error GoTo err_unit

'    If Me![Unit] <> "" Then
'        Dim getArea, getyear
'        getArea = DLookup("[Area]", "[Exca: Unit sheet with relationships]", "[Unit number] = " & Me![Unit])
'        If Not IsNull(getArea) Then
'            Me![Area] = getArea
'        Else
'            MsgBox "The Area field has not been automically obtained from the Excavation database, please check the excavation database directly", vbInformation, "Area Field"
'        End If
        
'        getyear = DLookup("[Year]", "[Exca: Unit sheet with relationships]", "[Unit number] = " & Me![Unit])
'        If Not IsNull(getyear) Then
'            Me![Year] = getyear
'        Else
'            MsgBox "The Year field has not been automically obtained from the Excavation database, please check the excavation database directly. The system has defaulted to this current year.", vbInformation, "Year Field"
'        End If
'    End If
Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub
