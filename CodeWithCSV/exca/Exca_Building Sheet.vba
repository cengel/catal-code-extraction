Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub cboFindBuilding_AfterUpdate()
'********************************************
'Find the selected building number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindBuilding_AfterUpdate

    If Me![cboFindBuilding] <> "" Then
        'for existing number the field with be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![Number].Enabled = False Then Me![Number].Enabled = True
        DoCmd.GoToControl "Number"
        DoCmd.FindRecord Me![cboFindBuilding]
        Me![cboFindBuilding] = ""
    End If
Exit Sub

err_cboFindBuilding_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboFindBuilding_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - 2009
On Error GoTo err_cbofindNot

    MsgBox "Sorry this Building cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFindBuilding].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'********************************************
'Add a new record
'
'SAJ v9.1
'********************************************
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Number"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoToImage_Click()
'********************************************************************
' New button for 2009 which allows any available images to be
' displayed
' SAJ
'********************************************************************
On Error GoTo err_cmdGoToImage_Click

Dim mydb As DAO.Database
Dim tmptable As TableDef, tblConn, I, msg, fldid
Set mydb = CurrentDb

    'get the field id for unit in the catalog that matches this year
    'NEW 2007 method where by portfolio now uses its own sql database
    Dim myq1 As QueryDef, connStr
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    'myq1.sql = "sp_Portfolio_GetUnitFieldID " & Me![Year] - year here is random as there isn't one
    myq1.sql = "sp_Portfolio_GetBuildingFieldID_2009 2009"
    
    ''WAS GETTING a 2010 TIMEOUT VERY COMMONLY HENCE CODE BELOW WHICH NOT REALLY HELP - when using my
    ''own login it worked fine so tracked it on login. When changed main DB of portfolio login from master
    ''to catalhoyuk is works fine (? also tried main DB to portfolio copy db but no difference from master)
    ''start = Timer    ' Set start time.
    ''        Do While Timer < start + 50
     ''           'DoEvents    ' Yield to other processes.
    ''        Loop
    Dim myrs As Recordset
    Set myrs = myq1.OpenRecordset
    
    
         
    ''MsgBox myrs.Fields(0).Value
    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
        fldid = 0
    Else
        fldid = myrs.Fields(0).Value
    End If
        
    myrs.Close
    Set myrs = Nothing
    myq1.Close
    Set myq1 = Nothing
    
    
    
    For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
             
    If tmptable.Connect <> "" Then
        tblConn = tmptable.Connect
        Exit For
    End If
    Next I
    
    If tblConn <> "" Then
        'If InStr(tblConn, "catalsql") = 0 Then
        If InStr(tblConn, "catalsql") = 0 Then
            'if on site the image can be loaded from the server directly into Access
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, 2009
            
        Else
            'database is running remotely must access images via internet
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Number])
        End If

    Else
        
    End If
    
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
    
Exit Sub

err_cmdGoToImage_Click:
    'If Err.Number = 3146 Then
    '
    '    response = MsgBox("Call to photo catalogue timed out - try again?", vbYesNo + vbQuestion, Err.Description)
    '    If response = vbYes Then
    '        start = Timer    ' Set start time.
    '        Do While Timer < start + 100
    '            'DoEvents    ' Yield to other processes.
    '        Loop
    '      Resume
    '    End If
    'Else
        Call General_Error_Trap
    'End If
    Exit Sub
End Sub

Private Sub cmdPrintBuildingSheet_Click()
'new for 2009
On Error GoTo err_cmdBuilding

    Dim both
    both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        DoCmd.OpenReport "R_BuildingSheet", acViewPreview, , "[Number] = " & Me![Number]
        If both = vbYes Then DoCmd.OpenReport "R_Units_in_Buildings", acViewPreview, , "[In_Building] = " & Me![Number]
Exit Sub

err_cmdBuilding:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReportProblem_Click()
'bring up a popup to allow user to report a problem
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "building;" & Me![Number]
    
Exit Sub

err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewBuildingsketch_Click()
On Error GoTo err_ViewBuildingsketch_Click
    Dim Path
    Dim fname, newfile
    
    'check if can find sketch image
    'using global constanst sktechpath Declared in globals-shared
    'path = "\\catal\Site_Sketches\Features\Sketches"
    Path = sketchpath2015 & "buildings\sketches\"
    Path = Path & "B" & Me![Number] & "*" & ".jpg"
    
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        newfile = fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "buildings\sketches\" & newfile
    
    If Dir(Path) = "" Then
        'directory not exist
        MsgBox "The sketch plan of this building has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_buildingsketch", acNormal, , , acFormReadOnly, , Me![Number]
    End If
 
Exit Sub

err_ViewBuildingsketch_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub EstProportionofBuildingEx_AfterUpdate()
'new 2010
On Error GoTo err_est

If Me![EstProportionofBuildingEx] = "" Or IsNull(Me![EstProportionofBuildingEx]) Then
    Me![chkInfill].Enabled = False
    Me![chkPartOcc].Enabled = False
    Me![chkComplete].Enabled = False
    Me![chkConstruction].Enabled = False
    Me![chkWalls].Enabled = False
    Me![chkOutline].Enabled = False
    Me![chkOther].Enabled = False
    Me![txtPropNotes].Enabled = False
Else
    Me![chkInfill].Enabled = True
    Me![chkPartOcc].Enabled = True
    Me![chkComplete].Enabled = True
    Me![chkConstruction].Enabled = True
    Me![chkWalls].Enabled = True
    Me![chkOutline].Enabled = True
    Me![chkOther].Enabled = True
    Me![txtPropNotes].Enabled = True
End If

Exit Sub

err_est:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
'********************************************
'This is the trowel button to close the form
' Error trap added v9.1
' removed open excavation form (menu) as this form
' can now be called by other forms
' SAJ v9.1
'********************************************
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String

    'stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Building Sheet"
Exit Sub
err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub






Private Sub Field24_AfterUpdate()
'show mound from area combo
Me![Mound] = Me![Field24].Column(1)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'***************************************************************************
' If no building number entered key violation occurs - intercept and provide
' nice msg, plus way to escape msg loop
'
' SAJ v9.1
'***************************************************************************
On Error GoTo err_Form_BeforeUpdate

If IsNull(Me![Number] And (Not IsNull(Me![Field24]) Or Not IsNull(Me![Location]) Or (Me![Description] <> "" And Not IsNull(Me![Description])))) Then
    MsgBox "You must enter a building number otherwise the record cannot be saved." & Chr(13) & Chr(13) & "If you wish to cancel this record being entered and start again completely press ESC", vbInformation, "Incomplete data"
    Cancel = True
    DoCmd.GoToControl "Number"
ElseIf IsNull(Me![Number]) And IsNull(Me![Field24]) And IsNull(Me![Location]) And (IsNull(Me![Description]) Or Me![Description] = "") Then
    'blank record that was edited but all data scrubbed - system still thinks it should
    'try and save the record as it created a shell for it so just tell it to undo
    Me.Undo
End If

If Me.Dirty And (IsNull(Me![LastUpdatedBy]) Or Me![LastUpdatedBy] = "") Then
    MsgBox "You must enter your name in the Last Updated By field", vbInformation, "Last Updated By Field not filled out"
    Cancel = True
    DoCmd.GoToControl "LastUpdatedBy"
End If

Me![LastUpdatedDate] = Date

Exit Sub

err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub


Private Sub Form_Current()
'*************************************************************************
' New requirement that building number cannot be edited after entry. This
' can only be done by an administrator so lock field here
'
' SAJ v9.1
'*************************************************************************
On Error GoTo err_Form_Open

   If Me![Number] <> "" Then
    'building number exists, lock field
        Me![Number].Locked = True
        Me![Number].Enabled = False
        Me![Number].BackColor = Me.Section(0).BackColor
        
        Me![Location].SetFocus
    Else
    'no building number - new record allow entry
        Me![Number].Locked = False
        Me![Number].Enabled = True
        Me![Number].BackColor = 16777215
        
        Me![Number].SetFocus
    End If
    
'See oncurrent of unit sheet for why I've now commented this out - button is enabled and image form deterimes if images
''    'new 2009
''find out is any images available
Dim imageCount, Imgcaption
''
''Dim mydb As DAO.Database
''Dim myq1 As QueryDef, connStr
''    Set mydb = CurrentDb
''    Set myq1 = mydb.CreateQueryDef("")
''    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
''    myq1.ReturnsRecords = True
''    myq1.sql = "sp_Portfolio_CountImagesForBuilding_2009 " & Me![Number]
''
''    Dim myrs1 As Recordset
''    Set myrs1 = myq1.OpenRecordset
''    ''MsgBox myrs.Fields(0).Value
''    If myrs1.Fields(0).Value = "" Or myrs1.Fields(0).Value = 0 Then
''           imageCount = 0
''    Else
''        imageCount = myrs1.Fields(0).Value
''   End If
''
''myrs1.close
''Set myrs1 = Nothing

backhere:
''myq1.close
''Set myq1 = Nothing
''mydb.close
''Set mydb = Nothing
''
''If imageCount > 0 Then
''    Imgcaption = imageCount
''    If imageCount = 1 Then
''        Imgcaption = Imgcaption & " Image to Display"
''    Else
''        Imgcaption = Imgcaption & " Images to Display"
''    End If
''    Me![cmdGoToImage].Caption = Imgcaption
''    Me![cmdGoToImage].Enabled = True
''Else
''    Me![cmdGoToImage].Caption = "No Image to Display"
''    Me![cmdGoToImage].Enabled = False
''End If
Imgcaption = "Images of Building"
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
    
'''OFFSITE 2009 - ignore photos and sketches offsite
'''JUST TAKE THIS LINE OUT ON SITE TO RETRIEVE FUNCTIONALITY
''Me![cmdGoToImage].Enabled = False

'New 2010 post ex fields
If Me![EstProportionofBuildingEx] = "" Or IsNull(Me![EstProportionofBuildingEx]) Then
    Me![chkInfill].Enabled = False
    Me![chkPartOcc].Enabled = False
    Me![chkComplete].Enabled = False
    Me![chkConstruction].Enabled = False
    Me![chkWalls].Enabled = False
    Me![chkOutline].Enabled = False
    Me![chkOther].Enabled = False
    Me![txtPropNotes].Enabled = False
Else
    Me![chkInfill].Enabled = True
    Me![chkPartOcc].Enabled = True
    Me![chkComplete].Enabled = True
    Me![chkConstruction].Enabled = True
    Me![chkWalls].Enabled = True
    Me![chkOutline].Enabled = True
    Me![chkOther].Enabled = True
    Me![txtPropNotes].Enabled = True
End If

'show mound from area combo
Me![Mound] = Me![Field24].Column(1)

Exit Sub

err_Form_Open:
    If Err.Number = 3146 Then 'odbc call failed, crops up every so often on all
    'sheets bar unit and have NO idea why except it always starts with building with no photos
    'but not all time, it occurs on Set myrs1 = myq1.OpenRecordset statement, tried everything
    '    Resume Next
        imageCount = "?"
        GoTo backhere
    Else
        'MsgBox myq1.sql
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'*************************************************************************
' OpenArg may now be used to set up record for dataentry - specific format
' must be used:
' OpenArgs = "NEW,Num:building num to be entered if known,Area:area assoc if known
' eg: "NEW,Num:300,Area:west"
' Then it can be unpicked by code
'
' SAJ v9.1
'*************************************************************************
On Error GoTo err_Form_Open

    If Not IsNull(Me.OpenArgs) Then
        Dim getArgs, whatTodo, NumKnown, AreaKnown
        Dim firstcomma, Action
        getArgs = Me.OpenArgs
        
        If Len(getArgs) > 0 Then
            'get 1st comma to see what action is
            firstcomma = InStr(getArgs, ",")
            If firstcomma <> 0 Then
                'get action word upto 1st comma
                Action = Left(getArgs, firstcomma - 1)
                'if new then create new rec
                If UCase(Action) = "NEW" Then DoCmd.GoToRecord acActiveDataObject, , acNewRec
                
                'check term num is present, getting its starting point
                NumKnown = InStr(UCase(getArgs), "NUM:")
                If NumKnown <> 0 Then
                    'num phrase is there so obtain it between 'num:' (ie start pt of num: plus its 4 chars)
                    'and place of next comma
                    NumKnown = Mid(getArgs, NumKnown + 4, InStr(NumKnown, getArgs, ",") - (NumKnown + 4))
                    Me![Number] = NumKnown 'add it to the number fld
                    Me![Number].Locked = True 'lock the number field
                  '  DoCmd.RunCommand acCmdSaveRecord
                    
                End If
                
                'check term area is present, getting its starting point
                AreaKnown = InStr(UCase(getArgs), "AREA:")
                If AreaKnown <> 0 Then
                    'area phrase is there so obtain it between 'area:' (ie start pt of area: plus its 5 chars)
                    'and end of str
                    AreaKnown = Mid(getArgs, AreaKnown + 5, Len(getArgs))
                    Me![Field24] = AreaKnown 'add it to the area fld
                    Me![Field24].Locked = True
                End If
            End If
        
            
            'disable find and add new in this instance
            Me![cboFindBuilding].Enabled = False
            Me![cmdAddNew].Enabled = False
            Me.AllowAdditions = False
        End If
    End If
    
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        'disable find and add new in this instance
        Me![cboFindBuilding].Enabled = False
        Me![cmdAddNew].Enabled = False
        Me.AllowAdditions = False
    Else
        'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
        DoCmd.GoToControl "cboFindBuilding"
    End If
    
    'now sort out view depending on permissions
    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And (Me.AllowAdditions = True Or Me.AllowDeletions = True Or Me.AllowEdits = True) Then
        'SAJ season 2006 - don't allow deletions from this screen
        ToggleFormReadOnly Me, False, "NoDeletions"
    Else
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
    End If
    
    'new 2010
    If Me![EstProportionofBuildingEx] = "" Or IsNull(Me![EstProportionofBuildingEx]) Then
        Me![chkInfill].Enabled = False
        Me![chkPartOcc].Enabled = False
        Me![chkComplete].Enabled = False
        Me![chkConstruction].Enabled = False
        Me![chkWalls].Enabled = False
        Me![chkOutline].Enabled = False
        Me![chkOther].Enabled = False
        Me![txtPropNotes].Enabled = False
    Else
        Me![chkInfill].Enabled = True
        Me![chkPartOcc].Enabled = True
        Me![chkComplete].Enabled = True
        Me![chkConstruction].Enabled = True
        Me![chkWalls].Enabled = True
        Me![chkOutline].Enabled = True
        Me![chkOther].Enabled = True
        Me![txtPropNotes].Enabled = True
    End If

Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Number_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure duplicate building numbers not entered
' which would result in nasty key violation msg back from sql server
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Number_AfterUpdate

Dim checknum

If Me![Number] <> "" Then
    'check that building num not exist
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but this Building Number already exists, please enter another number.", vbInformation, "Duplicate Building Number"
        
        If Not IsNull(Me![Number].OldValue) Then
            'return field to old value if there was one
            Me![Number] = Me![Number].OldValue
        Else
            'if its a new record and no oldval (ie: null error is try to set it to oldval)
            'then just undo entry - me![number].undo does not undo this field, only
            'me.undo will but this removes all edits in all fields so must gather them first (!!)
            Dim currloc, currarea, currdesc
            currloc = Me![Location]
            currarea = Me![Field24]
            currdesc = Me![Description]
            DoCmd.GoToControl "Number"
            Me.Undo
            
            'reset all fields, for some reason if description is null (and currdesc is null)
            'it won't set ti back to null, instead "". This throws out the form_beforeupdate
            'code, to ensure this not happen have added the if not isnull check, so only updates
            'field if there was an original value
            If Not IsNull(currloc) Then Me![Location] = currloc
            If Not IsNull(currarea) Then Me![Field24] = currarea
            If Not IsNull(currdesc) Then Me![Description] = currdesc
            
            'for some reason have to send focus to another field to bring it back
            'otherwise goes onto area- setfocus not work either
            DoCmd.GoToControl "Description"
            DoCmd.GoToControl "Number"
        End If
    End If
End If

Exit Sub

err_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
On Error GoTo Err_cmdHelp_Click

'either pop up a window or have a message
MsgBox "A help message to explain the Human Burial Assemblage field will appear soon", vbInformation, "Help"
    

Exit_cmdHelp_Click:
    Exit Sub

Err_cmdHelp_Click:
    Resume Exit_cmdHelp_Click
    
End Sub
