Option Compare Database
Option Explicit
Private Sub KnownFind()
'new season 2006 if this find is known to the basic data table then display its material from there
On Error GoTo err_knownfind

If Me![Unit] <> "" And Me![FindLetter] <> "" And Me![FindNumber] <> "" Then
    Dim getmaterial, getmaterialsub, getobject, GID
    GID = Me![Unit] & "." & Me![FindLetter] & Me![FindNumber]
    
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

Private Sub cboMoveCrate_AfterUpdate()
'new in season 2006 - move an item to a new crate - SAJ
On Error GoTo err_cboMove
    'bad sarah, lazy programming this code is repeated in store: subform units in crate2 - centralised this when time, just leaving 2006
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
        ''    MsgBox "Move has been successful"
        ''    'NOT THIS LINE HERE - cause write conflict ---
        ''    'Me.Requery
        ''    'aha 2008 (v3.2, saj) cracked the write conflict error that was really annoying as had to cpoy to clipboard
        ''    'to get move to work. Its becuase this is a bound field so simply undo the field value change (as not needed here
        ''    'as code has done it) and works ok
        ''    Me.Undo
        ''    Me.Requery
        ''
        ''    'NOT THIS LINE HERE AS BOUND TO FIELD HERE --- Me![cboMoveCrate] = ""
        ''    'no longer bound so do
        ''    Me![cboMoveCrate] = ""
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
        Dim sql
        Dim LResponse
        
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
        
                Me.Undo
                Me.Requery
                Me![cboMoveCrate] = ""
    
                MsgBox "Move has been successful"
                'requery underlying form as well - 26July11
                Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
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
        
            Me.Undo
            Me.Requery
            Me![cboMoveCrate] = ""
    
            MsgBox "Move has been successful"
            'requery underlying form as well - 26July11
            Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
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
        
        Me.Undo
        Me.Requery
        Me![cboMoveCrate] = ""
    
        MsgBox "Move has been successful"
        'requery underlying form as well - 26July11
        Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
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
        Me![cboMoveCrate].Visible = True
    Else
        Me![cboMoveCrate].Visible = False
    End If

Exit Sub

err_chkmove:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboMoveCrate_DblClick(Cancel As Integer)
On Error GoTo err_tracker
'new 2011 - find out where entry was previous located (if at all)

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

Exit Sub

err_tracker:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMoveCrate_GotFocus()
Me!cboMoveCrate.Requery

'MsgBox UCase(Me![CrateLetter])
' Added conditions here to make movemens possible only between certain crates - 2013

' a bit clunky, but works. CE 2013
Select Case CrateLetterFlag

    Case "FB"
    If (UCase(Me![CrateLetter]) = CrateLetterFlag Or UCase(Me![CrateLetter]) = "Depot") Then
'        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'Depot'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "P"
    If UCase(Me![CrateLetter]) = CrateLetterFlag Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "PH"
    If UCase(Me![CrateLetter]) = CrateLetterFlag Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "HB"
    If UCase(Me![CrateLetter]) = CrateLetterFlag Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "OB"
    If UCase(Me![CrateLetter]) = CrateLetterFlag Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "CONS"
    If UCase(Me![CrateLetter]) <> CrateLetterFlag Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number])) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    End If
    
    Case "CO"
    If ((UCase(Me![CrateLetter]) = "CO" Or UCase(Me![CrateLetter]) = "BE" Or UCase(Me![CrateLetter]) = "FG" Or UCase(Me![CrateLetter]) = "CB")) Then
'        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Store: Crate Register].CrateLetter FROM [Store: Crate Register] WHERE [Store: Crate Register].CrateLetter in ('" & CrateLetterFlag & "', 'BE', 'FG', 'CB') ORDER BY [Store: Crate Register].CrateLetter;"
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'BE', 'FG', 'CB'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
        Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "S"
    If ((UCase(Me![CrateLetter]) = "S" Or UCase(Me![CrateLetter]) = "BE")) Then
'        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Store: Crate Register].CrateLetter FROM [Store: Crate Register] WHERE [Store: Crate Register].CrateLetter in ('" & CrateLetterFlag & "', 'BE', 'FG', 'CB') ORDER BY [Store: Crate Register].CrateLetter;"
    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'BE'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
        Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "BE"
    If (UCase(Me![CrateLetter]) = CrateLetterFlag) Then
'        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "GS"
    If ((UCase(Me![CrateLetter]) = "GS" Or UCase(Me![CrateLetter]) = "NS" Or UCase(Me![CrateLetter]) = "Depot")) Then
'        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Store: Crate Register].CrateLetter FROM [Store: Crate Register] WHERE [Store: Crate Register].CrateLetter in ('" & CrateLetterFlag & "', 'NS', 'Depot') ORDER BY [Store: Crate Register].CrateLetter;"
         Me![cboMoveCrate].RowSource = "SELECT DISTINCT [Crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([CrateLetter] & [CrateNumber]) <> [Forms]![Store: Crate Register]![txtFullCrateName]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'NS', 'Depot'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
        Me![cboMoveCrate].RowSource = ""
    End If
    
    Case "Illustrate"
            Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    
    Case "Photo"
            Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    
    Case "char"
    If (UCase(Me![CrateLetter]) = CrateLetterFlag Or UCase(Me![CrateLetter]) = "or") Then
        Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) in ('" & CrateLetterFlag & "', 'or'))) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
    Else
         Me![cboMoveCrate].RowSource = ""
    End If




End Select


End Sub

Private Sub FindLetter_AfterUpdate()
If Me![FindLetter] <> "" Or Not IsNull(Me![FindLetter]) Then
    If UCase(Me![FindLetter]) <> "S" Then
        Me![FindLetter] = UCase(Me![FindLetter])
    End If
End If
'Call KnownFind
End Sub

Private Sub FindNumber_AfterUpdate()
'Call KnownFind
End Sub

Private Sub Form_AfterUpdate()
'moved from before update - saj season 2006
On Error GoTo err_afterupdate

    Forms![Store: Crate Register]![Date Changed] = Now()
    'new 2011 28/7/11 - Lisa said intermittent refresh of underlying form - hope this solves it
    Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
Exit Sub

err_afterupdate:
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
   
    Me![EtutlukNumber].Visible = False
    Me![EnvanterNumber].Visible = False
    Me![MuseumAccessionNumber].Visible = False
    Me![ExportLocation].Visible = False
    Me![Bag].Visible = False
    Me![Studied].Visible = False
    Me![txtNotes2].Visible = False
    Me![lblNotes2].Visible = False
    Me![txtNotes3].Visible = False
    Me![lblNotes3].Visible = False
    Me![lblAdditional].caption = "Notes"
    Me![lblMuseum].Visible = False
    
    If LCase(Forms![Store: Crate Register]![CrateLetter]) = "et" Then
         'etukluk crates must have an ET prefix
        Me![EtutlukNumber].Visible = True
        Me![lblAdditional].caption = "Etukluk No."
        Me![Bag].Visible = True
        Me![txtNotes].Visible = False
        Me![txtNotes3].Visible = False
        Me![txtNotes2].Visible = True
        Me![lblNotes2].Visible = True
    ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "envanter" Then
        'envanter crates must have envanter prefix
        Me![EnvanterNumber].Visible = True
        Me![MuseumAccessionNumber].Visible = True
        Me![lblMuseum].Visible = True
        Me![lblAdditional].caption = "Evanter No."
        Me![txtNotes].Visible = False
        Me![txtNotes3].Visible = False
        Me![txtNotes2].Visible = True
        Me![lblNotes2].Visible = True
    ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "export" Then
        'export crates must have export prefix
        Me![ExportLocation].Visible = True
        Me![lblAdditional].caption = "Export Location"
        Me![txtNotes].Visible = False
        Me![txtNotes3].Visible = False
        Me![txtNotes2].Visible = True
        Me![lblNotes2].Visible = True
    Else
        'all other crates have same fields apart from two
        Me![txtNotes].Visible = True
        
        If LCase(Forms![Store: Crate Register]![CrateLetter]) = "ob" Then
            'bag visible for chipped stone
            Me![Bag].Visible = True
            Me![txtNotes3].Visible = True
            Me![lblNotes3].Visible = True
            Me![txtNotes].Visible = False
            Me![lblAdditional].caption = "Bag"
        ElseIf LCase(Forms![Store: Crate Register]![CrateLetter]) = "fb" Then
            'studied visible for faunal
            Me![Studied].Visible = True
            Me![txtNotes3].Visible = True
            Me![lblNotes3].Visible = True
            Me![txtNotes].Visible = False
            Me![lblAdditional].caption = "Studied"
        End If
        
        
        Me![txtNotes2].Visible = False
        
    End If
    
'Dim rst As DAO.Recordset
'Set rst = Me.Form.Recordset

'Do While Not rst.EOF
'    MsgBox UCase(rst!CrateLetter)
'    If CrateLetterFlag = "OB" Then
'        If UCase(Forms![Store: Crate Register]![CrateLetter]) = "OB" Then
'         If UCase(rst!CrateLetter) = "OB" Then
'            MsgBox "UCase(rst!CrateLetter) is OB"
'            Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
'        Else
'            MsgBox "UCase(rst!CrateLetter) is not OB"
'            Me![cboMoveCrate].RowSource = ""
'        End If
'    End If
'    rst.MoveNext
'Loop



Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'new 2006 only let admins edit this via this form
On Error GoTo err_open

If GetGeneralPermissions = "Admin" Then
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me![cboMoveCrate].Visible = True
    Me![Label50].Visible = True
    Me![Text47].Visible = False
Else
    Me.AllowDeletions = False
    Me.AllowEdits = False
    Me![cboMoveCrate].Visible = True
    Me![Label50].Visible = False
    Me![Text47].Visible = False
End If

' added 2012 to allow team leaders to move bags inbetween crates
' edited 2013
If CrateLetterFlag = "FB" Or CrateLetterFlag = "CONS" Or CrateLetterFlag = "P" Or CrateLetterFlag = "CO" Or CrateLetterFlag = "HB" Or CrateLetterFlag = "OB" Or CrateLetterFlag = "GS" Or CrateLetterFlag = "ILLUSTRATE" Or CrateLetterFlag = "PHOTO" Or CrateLetterFlag = "char" Or CrateLetterFlag = "S" Or CrateLetterFlag = "BE" Or CrateLetterFlag = "PH" Then
    Me![cboMoveCrate].Visible = True
    Me.AllowEdits = True
    Me![Label50].Visible = True
    Me![Text47].Visible = False
End If

'' also added this to display only crates labelled for a particular team to be displayed in the
'' dropdown that shows the options to move units between crates. For example, Faunal team
'' only gets the FB crates to choose from.
'' CE - 2012 season

'' Conservation needs to see all crates, but no the others
'If CrateLetterFlag = "FB" Or CrateLetterFlag = "P" Or CrateLetterFlag = "GS" Or CrateLetterFlag = "OB" Or CrateLetterFlag = "BE" Or CrateLetterFlag = "" Or CrateLetterFlag = "FG" Then
'    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE ((([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) And (([Store: Crate Register].CrateLetter) Like '" & CrateLetterFlag & "')) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
'Else
'    Me![cboMoveCrate].RowSource = "SELECT DISTINCT [crateletter] & [CrateNumber] AS [crate number], [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber FROM [Store: Crate Register] WHERE (([crateletter] & [CrateNumber])<>[forms]![Store: find unit in crate2]![Crate Number]) ORDER BY [Store: Crate Register].CrateLetter, [Store: Crate Register].CrateNumber;"
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

    If Me![Unit] <> "" Then
        Dim getarea, getyear
        getarea = DLookup("[Area]", "[Exca: Unit sheet with relationships]", "[Unit number] = " & Me![Unit])
        If Not IsNull(getarea) Then
            Me![Area] = getarea
        Else
            MsgBox "The Area field has not been automically obtained from the Excavation database, please check the excavation database directly", vbInformation, "Area Field"
        End If
        
        getyear = DLookup("[Year]", "[Exca: Unit sheet with relationships]", "[Year] = " & Me![Year])
        If Not IsNull(getyear) Then
            Me![Year] = getyear
        Else
            MsgBox "The Year field has not been automically obtained from the Excavation database, please check the excavation database directly", vbInformation, "Year Field"
        End If
    End If
Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub
