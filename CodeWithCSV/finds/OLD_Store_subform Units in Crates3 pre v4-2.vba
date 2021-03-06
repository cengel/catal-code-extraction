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
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace, sql1, sql2, myq As QueryDef
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
        
        'insert into new crate
        sql1 = "INSERT INTO [Store: Units in Crates] "
        sql1 = sql1 & "( [Crate Number], [Unit number], Bag, [Letter/Number], Material, "
        sql1 = sql1 & "MaterialSubgroup, TempDescription, Notes, [Year], Area, Studied, "
        sql1 = sql1 & "CrateNumber, CrateLetter, FindLetter, FindNumber, SampleNumber, "
        sql1 = sql1 & "FlotNumber, EtutlukNumber, EnvanterNumber, MuseumAccessionNumber, "
        sql1 = sql1 & "ExportLocation ) "
        sql1 = sql1 & "SELECT '" & Me![cboMoveCrate] & "' AS 'Crate Number', "
        sql1 = sql1 & "[Store: Units in Crates].[Unit number], [Store: Units in Crates].Bag, "
        sql1 = sql1 & "[Store: Units in Crates].[Letter/Number], [Store: Units in Crates].Material, "
        sql1 = sql1 & "[Store: Units in Crates].MaterialSubgroup, [Store: Units in Crates].TempDescription, "
        sql1 = sql1 & "[Store: Units in Crates].Notes, [Store: Units in Crates].Year, "
        sql1 = sql1 & "[Store: Units in Crates].Area, [Store: Units in Crates].Studied, "
        sql1 = sql1 & Me![cboMoveCrate].Column(2) & " AS 'CrateNumber', '" & Me![cboMoveCrate].Column(1) & "' AS 'CrateLetter', "
        sql1 = sql1 & "[Store: Units in Crates].FindLetter, [Store: Units in Crates].FindNumber, "
        sql1 = sql1 & "[Store: Units in Crates].SampleNumber, [Store: Units in Crates].FlotNumber, "
        sql1 = sql1 & "[Store: Units in Crates].EtutlukNumber, [Store: Units in Crates].EnvanterNumber, "
        sql1 = sql1 & "[Store: Units in Crates].MuseumAccessionNumber, [Store: Units in Crates].ExportLocation "
        sql1 = sql1 & "FROM [Store: Units in Crates] "
        sql1 = sql1 & "WHERE [Store: Units in Crates].rowID = " & Me![rowID] & ";"
        
        Set myq = mydb.CreateQueryDef("")
        myq.sql = sql1
        myq.Execute
                
        myq.Close
        Set myq = Nothing
        
        'delete from here
        'sql2 = "DELETE FROM [Store: Units in Crates] WHERE [Store: Units in Crates].rowID = " & Me![rowID] & ";"
        If DeleteCrateRecord(Me![rowID], mydb) = False Then
            MsgBox "The delete part of this operation has failed", vbCritical, "Operation Failed"
        End If
        
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            'MsgBox "Move has been successful"
            Me.Requery
            Me![cboMoveCrate] = ""
            MsgBox "Move has been successful"
        Else
            wrkdefault.Rollback
            MsgBox "A problem has occured and the move has been cancelled. The error message is: " & Err.Description
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    
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
'2008 - need a way to ensure all X finds go into the basic data table
'v3.1
On Error GoTo err_findnum_upd
    
    If Me![FindLetter] <> "" Then
        If UCase(Me![FindLetter]) = "X" Then
            'check this exists
            Dim resp
            resp = DLookup("[GID]", "[Finds: Basic Data]", "[GID] = '" & Me![Unit] & ".X" & Me![FindNumber] & "'")
            If IsNull(resp) Then
                'GID missing display message to user. Would be good to add auto but then would need
                'to prompt for material group and subgroup etc and Jules not keen, so will start with this
                MsgBox "This X Find does not exist in the X Finds Register, please ensure you enter it.", vbCritical, "Data Validation"
                'move on cursor
                DoCmd.GoToControl "SampleNumber"
            End If
        End If
    End If
Exit Sub

err_findnum_upd:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_AfterUpdate()
'moved from before update - saj season 2006
On Error GoTo err_afterupdate

    Forms![Store: Crate Register]![Date Changed] = Now()

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
        End If
        
        
        'Me![txtNotes2].Visible = False
        
    End If
Exit Sub

err_current:
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
