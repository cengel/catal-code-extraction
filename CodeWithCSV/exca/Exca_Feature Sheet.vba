Option Compare Database
Option Explicit

Sub Close_Feature_Sheet_Click()
'CONTROL NOT SEEM TO EXIST - SAJ v9.1
'On Error GoTo Err_Close_Feature_Sheet_Click
'
'
'    DoCmd.close
'
'Exit_Close_Feature_Sheet_Click:
'    Exit Sub
'
'Err_Close_Feature_Sheet_Click:
'    MsgBox Err.Description
'    Resume Exit_Close_Feature_Sheet_Click
'
End Sub

Private Sub Building_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure building num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter building num when they can
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Building_AfterUpdate

Dim checknum, msg, retval

If Me![Building] <> "" Then
    'first check its valid
    If IsNumeric(Me![Building]) Then
    
        'check that building num does exist
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
        If IsNull(checknum) Then
            msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Combo27]
            End If
        Else
            'valid number, enable view button
            Me![cmdGoToBuilding].Enabled = True
        End If
    
    Else
        'not a vaild numeric building number
        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboFindFeature_AfterUpdate()
'********************************************
'Find the selected feature number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindFeature_AfterUpdate

    If Me![cboFindFeature] <> "" Then
        'for existing number the field with be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![Feature Number].Enabled = False Then Me![Feature Number].Enabled = True
        DoCmd.GoToControl "Feature Number"
        DoCmd.FindRecord Me![cboFindFeature]
        Me![cboFindFeature] = ""
    End If
Exit Sub

err_cboFindFeature_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindFeature_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - 2009
On Error GoTo err_cbofindNot

    MsgBox "Sorry this Feature cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFindFeature].Undo
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
    DoCmd.GoToControl "Feature Number"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoToBuilding_Click()
'***********************************************************
' Open building form with a filter on the number related
' to the button. Open as readonly.
'
' SAJ v9.1
'***********************************************************
On Error GoTo Err_cmdGoToBuilding_Click
Dim checknum, msg, retval, permiss

If Not IsNull(Me![Building]) Or Me![Building] <> "" Then
    'check that building num does exist
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
    If IsNull(checknum) Then
        'number not exist - now see what permissions user has
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            msg = "This Building Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
            If retval = vbNo Then
                MsgBox "No building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Combo27]
            End If
        Else
            'user is readonly so just tell them record not exist
            MsgBox "Sorry but this building record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Building Record"
        End If
    Else
        'record exists - open it
        Dim stDocName As String
        Dim stLinkCriteria As String

        stDocName = "Exca: Building Sheet"
    
        stLinkCriteria = "[Number]= " & Me![Building]
        'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog, "FILTER"
        'decided against dialog as can go to other forms from building sheet and if so they would open underneath it
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, , "FILTER"
    End If
    
End If

Exit Sub

Err_cmdGoToBuilding_Click:
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
    'myq1.sql = "sp_Portfolio_GetUnitFieldID " & Me![Year]
    myq1.sql = "sp_Portfolio_GetFeatureFieldID_2009 " & Me![Year]
    
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
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Feature Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
            
        Else
            'database is running remotely must access images via internet
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Feature Number])
        End If

    Else
        
    End If
    
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
    
Exit Sub

err_cmdGoToImage_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdPrintFeatureSheet_Click()
On Error GoTo err_print

    If Me![Feature Number] <> "" Then
        DoCmd.OpenReport "R_Feature_Sheet", acViewPreview, , "[feature number] = " & Me![Feature Number]
    End If
Exit Sub

err_print:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReportProblem_Click()
'bring up a popup to allow user to report a problem
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "feature number;" & Me![Feature Number]
    
Exit Sub

err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewFeaturematrix_Click()
'new season 2014 - open the feature matrix
On Error GoTo err_ViewFeaturematrix

    Dim Path
    Dim fname
    
    'check if can find sketch image
    'using global constanst sktechpath Declared in globals-shared
    'path = "\\catal\Site_Sketches\Features\Sketches"
    Path = sketchpath2015 & "features\matrices\"
    Path = Path & "F" & Me![Feature Number] & "*" & ".jpg"
    
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        Debug.Print fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "features\matrices\" & fname
    
    If Dir(Path) = "" Then
        'directory not exist
        MsgBox "The sketch plan of this unit has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_featurematrix", acNormal, , , acFormReadOnly, , Me![Feature Number]
    End If
 
Exit Sub

err_ViewFeaturematrix:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewFeaturesketch_Click()
'new season 2014 - open the feature sketch
'had not really been implemented in 2014 - starting over in 2015 - DL
On Error GoTo err_ViewFeaturesketch
    Dim Path
    Dim fname
    
    'check if can find sketch image
    'using global constanst sktechpath Declared in globals-shared
    'path = "\\catal\Site_Sketches\Features\Sketches"
    Path = sketchpath2015 & "features\sketches\"
    Path = Path & "F" & Me![Feature Number] & "*" & ".jpg"
    fname = Dir(Path & "*", vbNormal)
    
    While fname <> ""
        fname = Dir()
    Wend
    Path = sketchpath2015 & "features\sketches\" & fname
    Debug.Print Path
    
    If Dir(Path) = "" Then
        'directory not exist; convert_all.bat resides in root folder of sketches -
        'necessary to take it out of the equation - DL 2016
        MsgBox "The sketch plan of this unit has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_featuresketch", acNormal, , , acFormReadOnly, , Me![Feature Number]
    End If
 
Exit Sub

err_ViewFeaturesketch:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Combo27_AfterUpdate()
'********************************************
'Update the mound field to reflect the mound
'associated with the area, mound is now a read
'only field and users do not have to enter it
'
'SAJ v9.1
'********************************************
On Error GoTo err_Combo27_AfterUpdate

If Me![Combo27].Column(1) <> "" Then
    Me![Mound] = Me![Combo27].Column(1)
End If

Exit Sub
err_Combo27_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
' Existing close form button
' removed open excavation form (menu) as this form
' can now be called by other forms
' SAJ v9.1
On Error GoTo err_Excavation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    'stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Feature Sheet"
    
Exit_Excavation_Click:
    Exit Sub

err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub




Private Sub Feature_Number_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure duplicate feature numbers not entered
' which would result in nasty key violation msg back from sql server if not
' trapped. Duplicates were previously dealt with by an undo at Feature_number_exit,
' but this undo would blank the whole record with no explaination so trying
' to explain problem to user here.
'
' There is no way to programmatically undo the number field only the whole record
' this means any data the user has filled out before entering the feature number is lost.
' In the space and building sheets this was got round by taking a copy of any data
' in any of the fields, undoing the record to blank the duplicate key and then re-instating
' the field values for the user so they didn't have to re-type them. However those
' tables contained far fewer fields than this feature sheet and it would be unweldy
' to adopt the same approach.
'
' Instead a change of data entry approach has been introduced - make the feature number
' the first value entered and disable all fields until a valid entry is made here.
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Feature_Number_AfterUpdate
Dim checknum

If Me![Feature Number] <> "" Then
    'check that feature num not exist
    checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![Feature Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the Feature Number " & Me![Feature Number] & " already exists, please enter another number.", vbInformation, "Duplicate Feature Number"
        
        If Not IsNull(Me![Feature Number].OldValue) Then
            'return field to old value if there was one
            Me![Feature Number] = Me![Feature Number].OldValue
        Else
            'oh the joys, to keep the focus on feature have to flip to year then back
            'otherwise if will ignore you and go straight to year - dont believe me, comment out the gotocontrol year then!
            DoCmd.GoToControl "Year"
            DoCmd.GoToControl "Feature Number"
            Me![Feature Number].SetFocus
            
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        'the number does not exist so allow rest of data entry
        ToggleFormReadOnly Me, False
    End If
End If

'if after checks the field has a value hide the enter number msg
If Me![Feature Number] <> "" Then Me![lblMsg].Visible = False
Exit Sub

err_Feature_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Feature_Number_Exit(Cancel As Integer)
'*****************************************************
' This existing code is commented out and replaced by
' a handling procedure after update - the reason being
' this blanks all edits to this record done so far with
' no explaination to the user why, it also use legacy
' domenuitem.
' SAJ v9.1
'*****************************************************
'On Error GoTo Err_Feature_Number_Exit
'
'    Me.Refresh
'    'DoCmd.Save acTable, "Exca: Unit Sheet"
'
'Exit_Feature_Number_Exit:
'    Exit Sub
'
'Err_Feature_Number_Exit:
'
'    'MsgBox Err.Description
'
'    'MsgBox "This unit already exists in the database. Use the 'Find' button to go to it.", vbOKOnly, "duplicate"
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 0, , acMenuVer70
'
'    Cancel = True
'
'    Resume Exit_Feature_Number_Exit
End Sub



Private Sub Feature_Type_AfterUpdate()
'v9.2 SAJ - new feature subtype list must be updated to reflect feature type selection
On Error GoTo err_Feature_Type

If Me![Feature Type] <> "" Then
    'remove any previous entry in sub type field as won't match subtypes for new main type
    Me![cboFeatureSubType] = ""
    Me![cboFeatureSubType].RowSource = "SELECT [Exca:FeatureSubTypeLOV].FeatureSubType FROM [Exca:FeatureTypeLOV] INNER JOIN [Exca:FeatureSubTypeLOV] ON [Exca:FeatureTypeLOV].FeatureTypeID = [Exca:FeatureSubTypeLOV].FeatureTypeID WHERE ((([Exca:FeatureTypeLOV].FeatureType)='" & Me![Feature Type] & "')) ORDER BY [Exca:FeatureSubTypeLOV].FeatureSubType; "
    Me![cboFeatureSubType].Requery
End If

'new 2009 for burials
If LCase(Me![Feature Type]) = "burial" Then
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Then
        Me!txtBurialMNI.Enabled = True
        Me!txtBurialMNI.Enabled = False
    Else
        Me!txtBurialMNI.Enabled = False
        Me!txtBurialMNI.Enabled = True
    End If
End If

Exit Sub
err_Feature_Type:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'*******************************************************************
'Existing date change update - added error trap v9.1
'
'Also new requirement - if user edits record but no plan number exists
'then prompt them
' v9.1
'*******************************************************************
On Error GoTo err_Form_BeforeUpdate

If IsNull(Me![Exca: subform Feature Plans].Form![Graphic Number]) Then
    'this event will trigger when move to subform, so do not display then
    'this will mean if a user edits something above dimensions and tab on through
    'then moves to another record they will not get the message.
    'but they will get it if a new record thats entered to bottom
    If Me.ActiveControl.Name <> "Dimensions" And Me.ActiveControl.Name <> "Description" Then
        MsgBox "There is no Plan number entered for this Feature. Please can you enter one soon", vbInformation, "What is the Plan Number?"
    End If
End If

Me![Date changed] = Now()
Forms![Exca: Feature Sheet]![dbo_Exca: FeatureHistory].Form![lastmodify].Value = Now()

Exit Sub

err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'********************************************
'Check state of record, if no feature number its
'ie: new record make all fields read only so user must enter
' valid feature num before proceeding.
'
'This will also be useful when intro more adv security checking
'
' New requirement that feature number cannot be edited after entry. This
' can only be done by an administrator so lock field here
'SAJ v9.1
'********************************************
On Error GoTo err_Form_Current

'If IsNull(Me![Feature Number]) Or Me![Feature Number] = "" Then 'make rest of fields read only
'    ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
'    Me![lblMsg].Visible = True
'Else
'    ToggleFormReadOnly Me, False
'    Me![lblMsg].Visible = False
'End If

'after general formatting deal with any specifics here
'    If Me![Building] = "" Or IsNull(Me![Building]) Then
'        Me![cmdGoToBuilding].Enabled = False
'    Else
'        Me![cmdGoToBuilding].Enabled = True
'    End If

'overall check - is this user RW or admin then set up fields related to if new record or not
Dim permiss
permiss = GetGeneralPermissions
If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
    If IsNull(Me![Feature Number]) Or Me![Feature Number] = "" Then
        'new record so make rest of fields read only
        ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
        Me![lblMsg].Visible = True
        
        'no feature number - new record allow entry
        Me![Feature Number].Locked = False
        Me![Feature Number].Enabled = True
        Me![Feature Number].BackColor = 16777215
        Me![Feature Number].SetFocus
    Else
        'if coming in as a filter thats readonly then send in extra arg
        If Me.FilterOn = True And Me.AllowEdits = False Then
            'when popped up the building form this was allowing new records to be added, altered to fix
            'ToggleFormReadOnly Me, False, "NoAdditions"
            ToggleFormReadOnly Me, True, "NoAdditions"
        Else
            'if a filter is on remember no additions
            If Me.FilterOn Then
                ToggleFormReadOnly Me, False, "NoAdditions"
            Else
                'SAJ season 2006 - don't allow deletions from this screen
                ToggleFormReadOnly Me, False, "NoDeletions"
            End If
            'feature number exists, lock field
            Me![Year].SetFocus
            Me![Feature Number].Locked = True
            Me![Feature Number].Enabled = False
            Me![Feature Number].BackColor = Me.Section(0).BackColor
        End If
        Me![lblMsg].Visible = False
    End If

End If
    'after general formatting deal with any specifics here
    'moved into subform
    'If Me![Building] = "" Or IsNull(Me![Building]) Then
    '    Me![cmdGoToBuilding].Enabled = False
    'Else
    '    Me![cmdGoToBuilding].Enabled = True
    'End If

    If Me.FilterOn = True Or Me.AllowEdits = False Then
        'disable find and add new in this instance
        Me![cboFindFeature].Enabled = False
        Me![cmdAddNew].Enabled = False
    Else
        If Me![cboFindFeature].Enabled Then DoCmd.GoToControl "cboFindFeature"
    End If
    
    'Me![Feature Number].SetFocus
    
'v9.2 SAJ - new feature subtype dependant on feature main type - keep subtype combo linked with main type
Me![cboFeatureSubType].RowSource = "SELECT [Exca:FeatureSubTypeLOV].FeatureSubType FROM [Exca:FeatureTypeLOV] INNER JOIN [Exca:FeatureSubTypeLOV] ON [Exca:FeatureTypeLOV].FeatureTypeID = [Exca:FeatureSubTypeLOV].FeatureTypeID WHERE ((([Exca:FeatureTypeLOV].FeatureType)='" & Me![Feature Type] & "')) ORDER BY [Exca:FeatureSubTypeLOV].FeatureSubType; "
'Me![cboFeatureSubType].Requery

''LATE AUGUST 2009 SEASON
''We have recurring Error 52 Bad File name messages popping up until user UpdateDatabases, it will work a while
''and then reappear - is this related to this network call = timeout/corruption? Taking it out for now
''to see, when user presses button they will take pot luck on there being images
'''new 2009
'''find out is any images available
Dim imageCount, Imgcaption
''
''Dim mydb As DAO.Database
''Dim myq1 As QueryDef, connStr
''    Set mydb = CurrentDb
''    Set myq1 = mydb.CreateQueryDef("")
''    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
''    myq1.ReturnsRecords = True
''   myq1.sql = "sp_Portfolio_CountImagesForFeature_2009 '" & Me![Feature Number] & "', ''"
''
''    Dim myrs As Recordset
''   Set myrs = myq1.OpenRecordset
''    ''MsgBox myrs.Fields(0).Value
''    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
''           imageCount = 0
''    Else
''        imageCount = myrs.Fields(0).Value
''   End If
''
''myrs.close
''Set myrs = Nothing
''

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
Imgcaption = "Images of Feature"
Me![cmdGoToImage].Caption = Imgcaption
Me![cmdGoToImage].Enabled = True

'''OFFSITE 2009 - ignore photos and sketches offsite
'''JUST TAKE THIS LINE OUT ON SITE TO RETRIEVE FUNCTIONALITY
''Me![cmdGoToImage].Enabled = False
    
'new 2009 - lock up burial mni for everyone apart from admin
If permiss = "ADMIN" And LCase(Me![Feature Type]) = "burial" Then
    Me!txtBurialMNI.Enabled = True
    Me!txtBurialMNI.Locked = False
Else
    Me!txtBurialMNI.Enabled = False
    Me!txtBurialMNI.Locked = True
End If
Exit Sub

err_Form_Current:
    If Err.Number = 3146 Then 'odbc call failed, crops up every so often on all
    'sheets bar unit and have NO idea why except it always starts with building with no photos
    'but not all time, it occurs on Set myrs1 = myq1.OpenRecordset statement, tried everything
    '    Resume Next
        imageCount = "?"
        GoTo backhere
    Else
        Call General_Error_Trap
        Exit Sub
    End If
End Sub




Private Sub Form_Error(DataErr As Integer, response As Integer)
'************************************************************************
' This can catch runtime errors and intercept with a nicer message
'
' SAJ v9.1
'************************************************************************
Dim msg

If DataErr = 3162 Then
    'received this msg for invalid field entry of null
    ' eg: enter new feature number then delete it
    msg = "An error has occurred: invalid entry in the current field, probably a null value." & Chr(13) & Chr(13)
    msg = msg & "The system will attempt to resolve this, please re-try the action, but if you continue to get an error press the ESC key."
    MsgBox msg, vbInformation, "Error encountered"
    response = acDataErrContinue
    SendKeys "{ESC}"
    SendKeys "{ESC}"
ElseIf DataErr = 3146 Then
    'FIX: added 2 sendkey esc above - seems to stop this message
    'found that despite doing the above when user tries to move to a different record
    'its still coming back with a sql server error violation of primary
    'key constriant 'aaaaaExca_Features_PK'
    'MsgBox DataErr
    'SendKeys "{ESC}"
    'SendKeys "{ESC}"
    DoCmd.RunCommand acCmdUndo
    response = acDataErrContinue
    'this stops the error coming up but doesn't take user on to record they requested,
    'have to press record navigation again - see FIX Above
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
'*************************************************************************
' OpenArg may now be used to set up record for dataentry - specific format
' must be used:
' OpenArgs = "NEW,Num:feature num to be entered if known,Area:area assoc if known
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
                'if new then create new rec - this will cause on current to run
                If UCase(Action) = "NEW" Then DoCmd.GoToRecord acActiveDataObject, , acNewRec
    
                'check term num is present, getting its starting point
                NumKnown = InStr(UCase(getArgs), "NUM:")
                If NumKnown <> 0 Then
                    'num phrase is there so obtain it between 'num:' (ie start pt of num: plus its 4 chars)
                    'and place of next comma
                    NumKnown = Mid(getArgs, NumKnown + 4, InStr(NumKnown, getArgs, ",") - (NumKnown + 4))
                    Me![Feature Number] = NumKnown 'add it to the number fld
                    Me![Feature Number].Locked = True 'lock the number field
                  '  DoCmd.RunCommand acCmdSaveRecord

                End If

                'check term area is present, getting its starting point
                AreaKnown = InStr(UCase(getArgs), "AREA:")
                If AreaKnown <> 0 Then
                    'area phrase is there so obtain it between 'area:' (ie start pt of area: plus its 5 chars)
                    'and end of str
                    AreaKnown = Mid(getArgs, AreaKnown + 5, Len(getArgs))
                    Me![Combo27] = AreaKnown 'add it to the area fld
                    Me![Combo27].Locked = True
                End If
            End If

            'disable find and add new in this instance
            Me![cboFindFeature].Enabled = False
            Me![cmdAddNew].Enabled = False
            
            'creating the new record above will have called on current and set up form
            'as a new record, it won't have realised a feaure num has gone in and all fields
            'will still be locked so recall
            'Call Form_Current
            ToggleFormReadOnly Me, False
            Me.AllowAdditions = False
            Me![lblMsg].Visible = False
        End If
    Else
        'not a new record when opened so
        'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
        'moved to current post season
        'If Me![cboFindFeature].Enabled = True Then DoCmd.GoToControl "cboFindFeature"
    End If
    
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ' ToggleFormReadOnly Me, False ' on current will set it up for these users
    Else
        'set read only form here, just once
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
        Me![Feature Number].BackColor = Me.Section(0).BackColor
        Me![Feature Number].Locked = True
    End If
    
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub go_next_Click()
'pre-existing button to go to next record
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub


Private Sub go_previous_Click()
'pre-existing button to go to prev record
On Error GoTo Err_go_previous_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous_Click:
    Exit Sub

Err_go_previous_Click:
    MsgBox Err.Description
    Resume Exit_go_previous_Click
End Sub

Private Sub go_to_first_Click()
'pre-existing button to go to first record
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
'pre-existing button to go to last record
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
    
End Sub

Private Sub Master_Control_Click()
'THE FORM 'CATAL DATA ENTRY' NO LONGER EXISTS - what was it? - so made invis
'SAJ v9.1
'On Error GoTo Err_Master_Control_Click
'
'    Dim stDocName As String
'    Dim stLinkCriteria As String
'
'    stDocName = "Catal Data Entry"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
'    DoCmd.close acForm, "Exca: Feature Sheet"
'
'Exit_Master_Control_Click:
'    Exit Sub
'
'Err_Master_Control_Click:
'    MsgBox Err.Description
'    Resume Exit_Master_Control_Click
End Sub


Private Sub New_entry_Click()
'CONTROL NOT SEEM TO EXIST - SAJ v9.1
'On Error GoTo Err_New_entry_Click
'
'
'    DoCmd.GoToRecord , , acNewRec
'    Mound.SetFocus
'
'Exit_New_entry_Click:
'    Exit Sub
'
'Err_New_entry_Click:
'    MsgBox Err.Description
'    Resume Exit_New_entry_Click
End Sub


Sub find_feature_Click()
'REMOVED SAJ v9.1 REPLACE BY CBOFINDFEATURE - DUE TO LEGACY USE OF DOMENUITEM
'AND DANGER WITH FIND/REPLACE BOX
'On Error GoTo Err_find_feature_Click
'
'
'   Screen.PreviousControl.SetFocus
'    Feature_Number.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
'
'Exit_find_feature_Click:
'    Exit Sub
'
'Err_find_feature_Click:
'    MsgBox Err.Description
'    Resume Exit_find_feature_Click
'
End Sub

Private Sub print_bulk_Click()
On Error GoTo Err_print_bulk_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "print_bulk_features"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
   
   '' REMOVED SAJ AS WILL FAIL FOR RO USERS DoCmd.GoToRecord acForm, stDocName, acNewRec

Exit_print_bulk_Click:
    Exit Sub

Err_print_bulk_Click:
    Call General_Error_Trap
    Resume Exit_print_bulk_Click
End Sub
