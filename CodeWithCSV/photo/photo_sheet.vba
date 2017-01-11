Option Compare Database




Private Sub Close_Click()
On Error GoTo err_close

    DoCmd.OpenForm "Frm_Photo_MainMenu", acNormal, , , acFormPropertySettings
    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmbtechnical_AfterUpdate()
On Error GoTo err_cmbtechnical_AfterUpdate

If Trim(Me![technical]) = "3D Model" Then
    Me!Building.Enabled = True
    Me!Space.Enabled = True
    Me!Building.Visible = True
    Me!Space.Visible = True
    Me!buildingtitle.Visible = True
    Me!spacetitle.Visible = True
Else
    Me!Building.Enabled = False
    Me!Space.Enabled = False
    Me!Building.Visible = False
    Me!Space.Visible = False
    Me!buildingtitle.Visible = False
    Me!spacetitle.Visible = False
End If
Exit Sub

err_cmbtechnical_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmbtechnical_Change()
    If (Me!technical.Value = "3D Model") Then
        Me!modelname.Enabled = True
    Else
        Me!modelname.Enabled = False
    End If
End Sub

Private Sub cmdAddNew_Click()
'********************************************
'Add a new record
'
'taken from SAJ v9.1 - DL 2015
'********************************************
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "photographer"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub cmdRemoveFilter_Click()
On Error GoTo err_Removefilter

    Me![filterDSCN] = ""
    Me![filterPhotographer] = ""
    Me.Filter = ""
    Me.FilterOn = False
    
    Me![cmdRemoveFilter].Visible = False
   

Exit Sub

err_Removefilter:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub description_AfterUpdate()
On Error GoTo err_description_AfterUpdate

Dim searchndestroy As String

searchndestroy = Replace(Me!Description.Value, "'", "´")
Me!Description.Value = searchndestroy

Exit Sub

err_description_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub dscn_number_AfterUpdate()
Dim cleandscn

cleandscn = Trim(Me![dscn_number])
Me![dscn_number] = cleandscn

End Sub

Private Sub findDSCN_AfterUpdate()
'new 2016 - DL
On Error GoTo err_findDSCN

If Me![findDSCN] <> "" Then
    DoCmd.GoToControl "dscn_number"
    DoCmd.FindRecord Me![findDSCN]
    Me![findDSCN] = ""
End If


Exit Sub

err_findDSCN:
    Call General_Error_Trap
    Exit Sub
End Sub




Private Sub filterDSCN_AfterUpdate()
On Error GoTo err_filterDSCN

    If Me![filterDSCN] <> "" Then
        Me.Filter = "[dscn_number] = '" & Me![filterDSCN] & "'"
        Me.FilterOn = True
        Me![filterDSCN] = ""
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filterDSCN:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub filterPhotographer_AfterUpdate()
On Error GoTo err_filterPhotographer

    If Me![filterPhotographer] <> "" Then
        Me.Filter = "[photographer] = '" & Me![filterPhotographer] & "'"
        Me.FilterOn = True
        Me![filterPhotographer] = ""
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filterPhotographer:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
On Error GoTo err_Form_Current

Dim splittedbulk

If (Me!technical.Value = "3D Model") Then
    Me!modelname.Enabled = True
Else
    Me!modelname.Enabled = False
End If

If Me![dscn_number] <> "" And [dscn_number] <> "-" Then
    splittedbulk = Split(Me![dscn_number], "-")
    Debug.Print splittedbulk(0)
    If (splittedbulk(0) = "" And splittedbulk(1) <> "") Or (splittedbulk(0) <> "" And splittedbulk(1) = "") Or (splittedbulk(0) = splittedbulk(1)) Then
        Me!pushunique.Enabled = False
    Else
        Me!pushunique.Enabled = True
    End If
Else
End If

If Trim(Me![photographer] = "Jason Quinlan") Or (Me!technical.Value = "3D Model") Then
    Me!Building.Enabled = True
    Me!Space.Enabled = True
    Me!Building.Visible = True
    Me!Space.Visible = True
    Me!buildingtitle.Visible = True
    Me!spacetitle.Visible = True
Else
    Me!Building.Enabled = False
    Me!Space.Enabled = False
    Me!Building.Visible = False
    Me!Space.Visible = False
    Me!buildingtitle.Visible = False
    Me!spacetitle.Visible = False
End If



Exit Sub

err_Form_Current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open

Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
            Me![sync_portfolio].Enabled = False
    Else
            Me![sync_portfolio].Enabled = True
    End If
    
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub go_next_Click()
On Error GoTo Err_go_next_Click

    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub

Private Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click

    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    MsgBox Err.Description
    Resume Exit_go_previous2_Click
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click

    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()

On Error GoTo Err_go_to_last_Click

    DoCmd.GoToRecord , , acLast

Exit_go_to_last_Click:
    Exit Sub

Err_go_to_last_Click:
    MsgBox Err.Description
    Resume Exit_go_to_last_Click
End Sub

Private Sub photographer_AfterUpdate()
On Error GoTo err_photographer_AfterUpdate

If Trim(Me![photographer]) = "Jason Quinlan" Then
    Me!Building.Enabled = True
    Me!Space.Enabled = True
    Me!Building.Visible = True
    Me!Space.Visible = True
    Me!buildingtitle.Visible = True
    Me!spacetitle.Visible = True
Else
    Me!Building.Enabled = False
    Me!Space.Enabled = False
    Me!Building.Visible = False
    Me!Space.Visible = False
    Me!buildingtitle.Visible = False
    Me!spacetitle.Visible = False
End If
Exit Sub

err_photographer_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub photographer_NotInList(NewData As String, Response As Integer)

Dim strTmp As String

strTmp = "Add '" & NewData & "' as a new photographer?"
If MsgBox(strTmp, vbYesNo + vbDefaultButton2 + vbQuestion, "Not in List") = vbYes Then

    strTmp = "INSERT INTO [photographer] ([phototographer]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL strTmp

    Response = acDataErrAdded
End If

End Sub

Private Sub pushunique_Click()
On Error GoTo err_pushunique_Click

Dim splittedbulk
Dim generateuniques
Dim dscn, dashdscn, sqlstr, sqlstrunits, sqlstrfeatures, sqlstrkeywords
Dim newid As Object
Dim doubles
Dim rst, rstfeatures, rstunits, rstkeywords, rstsetsexists As DAO.Recordset
Dim dateformatted, newdateform, datecreated, newdatecreate As String
Dim I

If Me![dscn_number] <> "" And [dscn_number] <> "-" Then
    splittedbulk = Split(Me![dscn_number], "-")
    If (splittedbulk(0) = "" And splittedbulk(1) <> "") Or (splittedbulk(0) <> "" And splittedbulk(1) = "") Or (splittedbulk(0) = splittedbulk(1)) Then
    ElseIf (splittedbulk(0) <> "" And splittedbulk(1) <> "") Then
        generateuniques = MsgBox("You have entered an interval of DSCN numbers. Unique Datasets will be generated and aggregate dataset will be deleted. Continue?", vbOKCancel)
        If generateuniques = vbOK Then
        
        'Preparing the data for being fed into the interval dnsc datasets
        
        Set rst = CurrentDb.OpenRecordset("query_photo_features", dbOpenSnapshot, dbSeeChanges)
        With rst
            rst.Filter = "[photo_id]=" & Me.[photocentral_id]
            Set rstfeatures = rst.OpenRecordset
        End With
        Debug.Print rstfeatures.RecordCount
        Set rst = CurrentDb.OpenRecordset("query_photo_units", dbOpenSnapshot, dbSeeChanges)
        With rst
            rst.Filter = "[photo_id]=" & Me.[photocentral_id]
            Set rstunits = rst.OpenRecordset
        End With
        Debug.Print rstunits.RecordCount
        Set rst = CurrentDb.OpenRecordset("query_photo_keyword", dbOpenSnapshot, dbSeeChanges)
        With rst
            rst.Filter = "[photo_id]=" & Me.[photocentral_id]
            Set rstkeywords = rst.OpenRecordset
        End With
        
        For dscn = splittedbulk(0) To splittedbulk(1)
            If dscn < 0 Then
            Else
                If dscn < 10 Then
                dashdscn = dscn
                dashdscn = "-000" & dashdscn
                ElseIf dscn < 100 Then
                dashdscn = dscn
                dashdscn = "-00" & dashdscn
                ElseIf dscn < 1000 Then
                dashdscn = dscn
                dashdscn = "-0" & dashdscn
                Else
                dashdscn = dscn * -1
                End If
            End If
            dateformatted = Format(Me!date_taken, "yyyy-mm-dd hh:mm:ss")
            datecreated = Replace(dateformatted, ".", "-")
            sqlstr = "Select [photocentral_id] FROM [photo_id_by_date] WHERE [datecreated] = '" & datecreated & "' AND [dscn_number] = '" & dashdscn & "'"
            Set rstsetexists = CurrentDb.OpenRecordset(sqlstr, dbOpenSnapshot, dbSeeChanges)

            If Not (rstsetexists.BOF And rstsetexists.EOF) Then
                rstsetexists.MoveFirst
                doubles = doubles & dashdscn & ", "
            Else
                news = news & dashdscn & ", "
                               
                sqlstr = "INSERT INTO [photo_central] ([date_taken],[dscn_number],[area],[space],[building],[assocbuilding],[assocspace],[description],[direction],[rating],[photoview],[technical],[modelname],[photographer]) VALUES ('" & Me!date_taken & "','" & Trim(dashdscn) & "'," & Me!Area.Value & "," & IIf(IsNull(Me!Space.Value), "Null", Me!Space.Value) & "," & IIf(IsNull(Me!Building.Value), "Null", Me!Building.Value) & "," & IIf(IsNull(Me!assocbuilding.Value), "Null", Me!assocbuilding.Value) & "," & IIf(IsNull(Me!assocspace.Value), "Null", Me!assocspace.Value) & ",'" & Me!Description.Value & "','" & Me!direction.Value & "','" & Me!rating.Value & "','" & Me!photoview.Value & "','" & Me!technical.Value & "','" & Me!modelname.Value & "','" & Me!photographer.Value & "');"
                DoCmd.RunSQL sqlstr
                
                newdateform = Format(Me!date_taken, "yyyy-mm-dd hh:mm:ss")
                newdatecreate = Replace(newdateform, ".", "-")
                sqlstr = "Select [photocentral_id] FROM [photo_id_by_date] WHERE [datecreated] = '" & newdatecreate & "' AND [dscn_number] = '" & dashdscn & "'"
                Debug.Print sqlstr
                Set newid = CurrentDb.OpenRecordset(sqlstr, dbOpenSnapshot, dbSeeChanges)
                
                If Not (newid.BOF And newid.EOF) Then
                    I = 0
                    newid.MoveFirst
                    If Not (rstfeatures.BOF And rstfeatures.EOF) Then
                    rstfeatures.MoveLast
                    rstfeatures.MoveFirst
                    For I = 1 To rstfeatures.RecordCount
                        sqlstrfeatures = "INSERT INTO [photo_features]([photo_id],[Feature]) VALUES (" & newid.[photocentral_id] & "," & rstfeatures.[Feature] & ")"
                        DoCmd.RunSQL sqlstrfeatures
                        rstfeatures.MoveNext
                    Next I
                    Else
                    End If
                    I = 0
                    If Not (rstunits.BOF And rstunits.EOF) Then
                    rstunits.MoveLast
                    rstunits.MoveFirst
                    For I = 1 To rstunits.RecordCount
                        sqlstrunits = "INSERT INTO [photo_units]([photo_id],[Unit Number]) VALUES (" & newid.[photocentral_id] & "," & rstunits.[Unit Number] & ")"
                        DoCmd.RunSQL sqlstrunits
                        rstunits.MoveNext
                    Next I
                    Else
                    End If
                    I = 0
                    If Not (rstkeywords.BOF And rstkeywords.EOF) Then
                    rstkeywords.MoveLast
                    rstkeywords.MoveFirst
                    For I = 1 To rstkeywords.RecordCount
                        sqlstrkeywords = "INSERT INTO [photo_keywords]([photo_id],[keyword_id]) VALUES (" & newid.[photocentral_id] & "," & rstkeywords.[keyword_id] & ")"
                        DoCmd.RunSQL sqlstrkeywords
                        rstkeywords.MoveNext
                    Next I
                    Else
                    End If
                Else


                End If
                
            End If
        Next dscn
        
        MsgBox "Datasets inserted: " & news & Chr(13) & "Already existing datasets: " & doubles
        sqlstr = "DELETE FROM [photo_central] WHERE [photocentral_id]= " & Me!photocentral_id
        Debug.Print sqlstr
        DoCmd.RunSQL sqlstr
        Else
        End If
        
        
    Else
    
    End If
    
Else
End If

Me.Requery

Exit Sub

err_pushunique_Click:
    Dim errX As DAO.Error
    
    If Errors.Count > 1 Then
    For Each errX In DAO.Errors
        Debug.Print "ODBC Error: " & errX.Number
        Debug.Print errX.Description
    Next errX
    Else
        Debug.Print "VBA Error: " & Err.Number
        Debug.Print Err.Description
    End If
    Exit Sub
End Sub

Private Sub sync_portfolio_Click()
On Error GoTo err_sync_portfolio_Click

Dim dateformated, datecreated, datecreatedreduced, dateclean
Dim flag3d, flaginitials, sqlstr, sqlupitem, Area, dscnclean, photofile, stringvalue, longstring, sqlflag, flaggederase As String
Dim report_success, report_nonsync, syncstatus
Dim rst, rstnm, rstarea, rstnamedarea, rsp, rstrecord, rstupitem, rstincustom, rstfeatures, rstunits, rstkeywords, rstcheckrecord As DAO.Recordset
Dim qdf, qdfupitem, qdfincustom As QueryDef
Dim fieldid, integervalue, valueorder As Integer
Dim decimalvalue As Double

DoCmd.OpenForm "status_portfoliosync", acNormal
Forms![status_portfoliosync].Form![statuswindow] = "Start Syncing on " & Now() & Chr(13) & Chr(10)


Set rst = CurrentDb.OpenRecordset("photo_flaggedphoto", dbOpenSnapshot, dbSeeChanges)
Set rstarea = CurrentDb.OpenRecordset("dbo_Exca: Area Sheet", dbOpenSnapshot, dbSeeChanges)

If Not (rst.BOF And rst.EOF) Then
    rst.MoveLast
    rst.MoveFirst
    
    Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "Opening " & rst.RecordCount & " records from central register" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    For I = 1 To rst.RecordCount
        If (Trim(rst![technical]) = "3D Model") Then
            flag3d = "3d"
        Else
            flag3d = ""
        End If
        
        If Trim(rst![photographer]) = "Jason Quinlan" Then
            flaginitials = "_jpq_"
        Else
            flaginitials = "_"
        End If
        
        datecreated = Format(rst![date_taken], "yyyy/mm/dd")
        If InStr(datecreated, ".") > 0 Then
            datecreatedreduced = Replace(datecreated, ".", "")
            dateformated = datecreatedreduced
        Else
            datecreatedreduced = Replace(datecreated, "/", "")
            dateformated = datecreatedreduced
        End If
        
        Debug.Print dateformated
        
        With rstarea
            rstarea.Filter = "[area number]=" & rst![Area]
            Set rstnamedarea = rstarea.OpenRecordset
        End With
        dscnclean = Trim(Replace(rst![dscn_number], "-", ""))
        
        If flaginitials = "_jpq_" And flag3d = "" Then
            photofile = dateformated & flaginitials & dscnclean & ".jpg"
        Else
            photofile = dateformated & "_" & rstnamedarea![area name] & flaginitials & dscnclean & ".jpg"
        End If
                
        Debug.Print photofile
        
        Set qdf = CurrentDb.CreateQueryDef("")
        qdf.Connect = CurrentDb.TableDefs("[photo_central]").Connect
        qdf.ReturnsRecords = True

        qdf.sql = "sp_" & flag3d & "photo_get_portfoliorecord '" & photofile & "'"
        Debug.Print qdf.sql
        
        Set rstrecord = qdf.OpenRecordset

        If Not (rstrecord.BOF And rstrecord.EOF) Then
    
            rstrecord.MoveLast
            rstrecord.MoveFirst
            
            Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & I & ": " & photofile & " -> RecordID " & rstrecord![record_id] & " [Images" & flag3d & "]" & Chr(13) & Chr(10)
            
            Set qdfupitem = CurrentDb.CreateQueryDef("")
            With qdfupitem
                .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                .ReturnsRecords = True
                .sql = "exec sp_photo_check_itemtable @recordid=" & rstrecord![record_id] & ", @filename ='" & photofile & "'"
                Debug.Print "exec sp_photo_check_itemtable @recordid=" & rstrecord![record_id] & ", @filename ='" & photofile & "'"
                Set rstcheckrecord = qdfupitem.OpenRecordset
                    
                If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                    .ReturnsRecords = False
                    If rstcheckrecord![File_Description] <> "" Then
                        .sql = "exec sp_" & flag3d & "photo_update_item_table @recordid=" & rstrecord![record_id] & ", @filename='" & photofile & "', @updatedby='" & "photoadmin" & "', @filedescription='" & rstcheckrecord![File_Description] & "'"
                    Else
                        .sql = "exec sp_" & flag3d & "photo_update_item_table @recordid=" & rstrecord![record_id] & ", @filename='" & photofile & "', @updatedby='" & "photoadmin" & "', @filedescription='" & rst![Description] & "'"
                    End If
                    .Execute
                Else
                    Debug.Print "Item does not exist? Really?"
                End If
            End With
            
            'Populating custom data fields for 1:1 datasets related to photo item

            If rst![Area] <> "" Then
                fieldid = 10087
                
                'Work around for areaname DigHouse being Dig House (with a space) in Portfolio - DAL 2016
                    If rstnamedarea![area name] = "DigHouse" Then
                        stringvalue = "Dig House"
                    Else
                        stringvalue = rstnamedarea![area name]
                    End If
                    
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10087, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Debug.Print "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10087, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "area exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10087, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10087, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no area (?)" & Chr(13) & Chr(10)
            End If
            If rst![Space] <> "" Then
                fieldid = 10098
                stringvalue = ""
                integervalue = rst![Space]
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "space exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no space" & Chr(13) & Chr(10)
            End If
            If rst![Building] <> "" Then
                fieldid = 10088
                stringvalue = ""
                integervalue = rst![Building]
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "building exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no building" & Chr(13) & Chr(10)
            End If
            If rst![assocbuilding] <> "" Then
                fieldid = 10116
                stringvalue = ""
                integervalue = rst![assocbuilding]
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "associated building exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no associated building" & Chr(13) & Chr(10)
            End If
            If rst![assocspace] <> "" Then
                fieldid = 10118
                stringvalue = ""
                integervalue = rst![assocspace]
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "associated space exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no associated space" & Chr(13) & Chr(10)
            End If
            If Trim(rst![direction]) <> "" Then
                fieldid = 10089
                stringvalue = Trim(rst![direction])
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10089, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "direction exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10089, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10089, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no direction" & Chr(13) & Chr(10)
            End If
            If Trim(rst![rating]) <> "" Then
                fieldid = 10096
                stringvalue = Trim(rst![rating])
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10096, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "rating exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10096, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10096, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no rating" & Chr(13) & Chr(10)
            End If
            If Trim(rst![photoview]) <> "" Then
                fieldid = 10102
                stringvalue = Trim(rst![photoview])
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10102, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "photoview exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10102, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10102, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no photoview" & Chr(13) & Chr(10)
            End If
            If Trim(rst![technical]) <> "" Then
                fieldid = 10100
                stringvalue = Trim(rst![technical])
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10100, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "technical exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10100, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10100, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no technical" & Chr(13) & Chr(10)
            End If
            'If Trim(rst![status]) <> "" Then
            '    fieldid = 10099
            '    stringvalue = Trim(rst![status])
            '    Set qdfincustom = CurrentDb.CreateQueryDef("")
            '    With qdfincustom
            '        .Connect = CurrentDb.TableDefs("[photo_central]").Connect
            '        .ReturnsRecords = False
            '        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10099, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            '        .Execute
            '    End With
            '    Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10099, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            'Else
            '    Debug.Print "no status"
            'End If
            If Trim(rst![photographer]) <> "" Then
                fieldid = 10095
                stringvalue = Trim(rst![photographer])
                Set qdfincustom = CurrentDb.CreateQueryDef("")
                With qdfincustom
                    .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                    .ReturnsRecords = True
                    .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10095, "Null", "Null") & ", @decimal=" & decimalvalue & ""
                    Set rstcheckrecord = qdfincustom.OpenRecordset
                    
                    If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                        Debug.Print "photographer exists already"
                    Else
                        .ReturnsRecords = False
                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10095, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                        .Execute
                    End If
                End With
                Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & IIf(fieldid = 10095, "Null", "Null") & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
            Else
                Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no photographer" & Chr(13) & Chr(10)
            End If

            'Populating custom data fields for 1:n datasets related to photo item

            Set rstnm = CurrentDb.OpenRecordset("photo_features", dbOpenSnapshot, dbSeeChanges)
            rstnm.Filter = "[photo_id]=" & rst![photo_id]
            Set rstfeatures = rstnm.OpenRecordset
            Debug.Print rstnm.RecordCount
            If Not (rstnm.BOF And rstnm.EOF) Then
                With rstnm
                    If Not (rstfeatures.BOF And rstfeatures.EOF) Then
                        rstfeatures.MoveLast
                        rstfeatures.MoveFirst
                        For k = 1 To rstfeatures.RecordCount
                            fieldid = 10090
                            stringvalue = ""
                            integervalue = rstfeatures![Feature]
                            Set qdfincustom = CurrentDb.CreateQueryDef("")
                            With qdfincustom
                                .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                                .ReturnsRecords = True
                                .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                                Set rstcheckrecord = qdfincustom.OpenRecordset

                                If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                                    Debug.Print "feature exists already"
                                Else
                                    If integervalue <> 0 Then
                                        .ReturnsRecords = False
                                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                                        .Execute
                                    Else
                                    'value is 0: nonsense
                                    End If
                                End If
                            End With
                            Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                            rstfeatures.MoveNext
                        Next k
                    Else
                        Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no features" & Chr(13) & Chr(10)
                    End If
                End With
            Else
            End If
            Debug.Print rstfeatures.RecordCount
            Set rstnm = CurrentDb.OpenRecordset("photo_units", dbOpenSnapshot, dbSeeChanges)
            rstnm.Filter = "[photo_id]=" & rst![photo_id]
            Set rstunits = rstnm.OpenRecordset
            If Not (rstnm.BOF And rstnm.EOF) Then
                With rstnm
                    If Not (rstunits.BOF And rstunits.EOF) Then
                        rstunits.MoveLast
                        rstunits.MoveFirst
                        For k = 1 To rstunits.RecordCount
                            fieldid = 10101
                            stringvalue = ""
                            integervalue = rstunits![Unit Number]
                            Set qdfincustom = CurrentDb.CreateQueryDef("")
                            With qdfincustom
                                .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                                .ReturnsRecords = True
                                .sql = "exec sp_" & flag3d & "photo_check_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ""
                                Set rstcheckrecord = qdfincustom.OpenRecordset

                                If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                                Debug.Print "unit exists already"
                                Else
                                    If integervalue <> 0 Then
                                        .ReturnsRecords = False
                                        .sql = "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                                        .Execute
                                    Else
                                    'unit value is 0: nonsense
                                    End If
                                End If
                            End With
                            Debug.Print "exec sp_" & flag3d & "photo_insert_customdata @recordid=" & rstrecord![record_id] & ", @fieldid=" & fieldid & ", @stringvalue='" & stringvalue & "', @longstring='" & longstring & "', @integer=" & integervalue & ", @decimal=" & decimalvalue & ", @valueorder=" & valueorder & ""
                            rstunits.MoveNext
                        Next k
                    Else
                        Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no units" & Chr(13) & Chr(10)
                    End If
                End With
            Else
            End If
            Debug.Print rstunits.RecordCount
            Set rstnm = CurrentDb.OpenRecordset("photo_keywords", dbOpenSnapshot, dbSeeChanges)
            rstnm.Filter = "[photo_id]=" & rst![photo_id]
            Set rstkeywords = rstnm.OpenRecordset
            If Not (rstnm.BOF And rstnm.EOF) Then
                With rstnm
                    If Not (rstkeywords.BOF And rstkeywords.EOF) Then
                        rstkeywords.MoveLast
                        rstkeywords.MoveFirst
                        For k = 1 To rstkeywords.RecordCount
                            integervalue = rstkeywords![keyword_id]
                            Set qdfincustom = CurrentDb.CreateQueryDef("")
                            With qdfincustom
                                .Connect = CurrentDb.TableDefs("[photo_central]").Connect
                                .ReturnsRecords = True
                                .sql = "exec sp_" & flag3d & "photo_check_keyword @recordid=" & rstrecord![record_id] & ", @keyword_id=" & integervalue & ""
                                Set rstcheckrecord = qdfincustom.OpenRecordset

                                If Not (rstcheckrecord.BOF And rstcheckrecord.EOF) Then
                                    Debug.Print "keyword exists already"
                                Else
                                    .ReturnsRecords = False
                                    .sql = "exec sp_" & flag3d & "photo_insert_keyword @recordid=" & rstrecord![record_id] & ", @keyword_id=" & integervalue & ""
                                    .Execute
                                End If
                            End With
                            Debug.Print "exec sp_" & flag3d & "photo_insert_keyword @recordid=" & rstrecord![record_id] & ", @keyword_id=" & integervalue & ""
                            rstkeywords.MoveNext
                        Next k
                    Else
                        Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "no keywords" & Chr(13) & Chr(10)
                    End If
                End With
            Else
            End If

            If (flag3d = "3d") Then
                sqlflag = "INSERT INTO [photo_flag]([photocentral_id]) VALUES (" & rst![photo_id] & ")"
                DoCmd.RunSQL sqlflag
            Else
                flaggederase = "DELETE FROM [photo_central] WHERE [photocentral_id] = " & rst![photo_id]
                DoCmd.RunSQL flaggederase
            End If
            
            Forms![status_portfoliosync].Form![statuswindow].Value = Forms![status_portfoliosync].Form![statuswindow].Value & "-> dataset synced" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            report_success = report_success + "RecordID: " & rstrecord![record_id] & ", File: " & photofile & Chr(13) & Chr(10)
        Else
            Debug.Print "Kein Datensatz"
            report_nonsync = report_nonsync + "Filename: " & photofile & Chr(13) & Chr(10)
        End If
        rst.MoveNext
    Next I
Else
End If

Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "Syncing was succesful for the following datasets:" & Chr(13) & Chr(10)
Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & report_success & Chr(13) & Chr(10)
Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & "The following images could not be found in Portfolio:" & Chr(13) & Chr(10)
Forms![status_portfoliosync].Form![statuswindow] = Forms![status_portfoliosync].Form![statuswindow] & report_nonsync

Exit Sub

err_sync_portfolio_Click:
    Dim errX As DAO.Error
    
    If Errors.Count > 1 Then
    For Each errX In DAO.Errors
        Debug.Print "ODBC Error: " & errX.Number
        Debug.Print errX.Description
    Next errX
    Else
        Debug.Print "VBA Error: " & Err.Number
        Debug.Print Err.Description
    End If

    'Call General_Error_Trap
    Exit Sub
End Sub

Private Sub tglDataSheet_Click()
'********************************************************************
' The user wants to see the basic data in datasheet view
' DL
'********************************************************************
On Error GoTo Err_tglDataSheet

    Me![date_taken].SetFocus
    DoCmd.RunCommand acCmdDatasheetView

Exit Sub

Err_tglDataSheet:
    Call General_Error_Trap
    Exit Sub

End Sub
