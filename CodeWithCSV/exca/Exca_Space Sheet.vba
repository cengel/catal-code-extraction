Option Compare Database   'Use database order for string comparisons
Option Explicit

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
        'new 2009 winter - if a user gets number wrong they can't delete it and many then put 0
        'building 0 keeps appearing and we don't want it so put a check in
        If Me![Building] = 0 Then
            MsgBox "Building 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
            Me![Building] = ""
        Else
            'check that building num does exist
            checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
            If IsNull(checknum) Then
                msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
                If retval = vbNo Then
                    MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
                Else
                    DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Field26]
                End If
            Else
                'valid number, enable view button
                Me![cmdGoToBuilding].Enabled = True
            End If
            'building number entered so internal space - new season 2009
            Me![chkExternal] = False
        End If
    Else
        'not a vaild numeric building number
        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
Else
    'no building number entered so external space - new season 2009
    Me![chkExternal] = True
End If

Exit Sub

err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub






Private Sub cboFindSpace_AfterUpdate()
'********************************************
'Find the selected space number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindSpace_AfterUpdate

    If Me![cboFindSpace] <> "" Then
        'if number is disabled then must enable if for the search
        'it will be reset to disabled by the code in form oncurrent
        If Me![Space number].Enabled = False Then Me![Space number].Enabled = True
        DoCmd.GoToControl "Space Number"
        DoCmd.FindRecord Me![cboFindSpace]
        Me![cboFindSpace] = ""
    End If
Exit Sub

err_cboFindSpace_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindSpace_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - 2009
On Error GoTo err_cbofindNot

    MsgBox "Sorry this Space cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFindSpace].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboHodderPhase_AfterUpdate()
'NEW WINTER 2009 - READY FOR INTRO OF NEW PHASING MODEL
On Error GoTo err_phase

    If Me![cboHodderPhase] <> "" Then
        Me![txtHodderPhase] = Me![Area] & "." & Me![cboHodderPhase]
        
        'winter 2010 - bear in mind the timeperiod field of the unit must be updated to reflect any change
        'in 2010 some have got out of sync eg: phased post chalc but still say Neolithic
        Dim sql
        If Me![cboHodderPhase] = "Post-Chalcolithic" Then
            sql = "UPDATE ([Exca: Space Sheet] INNER JOIN [Exca: Units in Spaces] ON [Exca: Space Sheet].[Space number] = "
            sql = sql & "[Exca: Units in Spaces].In_space) INNER JOIN [Exca: Unit Sheet] ON [Exca: Units in Spaces].Unit = "
            sql = sql & "[Exca: Unit Sheet].[Unit Number] SET [Exca: Unit Sheet].TimePeriod = 'Post-Chalcolithic' "
            sql = sql & "WHERE ((([Exca: Space Sheet].[Space number])=" & Me![Space number] & "));"
            DoCmd.RunSQL sql
        ElseIf Me![cboHodderPhase] = "Chalcolithic" Then
            sql = "UPDATE ([Exca: Space Sheet] INNER JOIN [Exca: Units in Spaces] ON [Exca: Space Sheet].[Space number] = "
            sql = sql & "[Exca: Units in Spaces].In_space) INNER JOIN [Exca: Unit Sheet] ON [Exca: Units in Spaces].Unit = "
            sql = sql & "[Exca: Unit Sheet].[Unit Number] SET [Exca: Unit Sheet].TimePeriod = 'Chalcolithic' "
            sql = sql & "WHERE ((([Exca: Space Sheet].[Space number])=" & Me![Space number] & "));"
            DoCmd.RunSQL sql
        ElseIf Len(Me![cboHodderPhase]) < 3 Then 'will be a ?letter or letter = neolithic - this will ignore unknown etc
            sql = "UPDATE ([Exca: Space Sheet] INNER JOIN [Exca: Units in Spaces] ON [Exca: Space Sheet].[Space number] = "
            sql = sql & "[Exca: Units in Spaces].In_space) INNER JOIN [Exca: Unit Sheet] ON [Exca: Units in Spaces].Unit = "
            sql = sql & "[Exca: Unit Sheet].[Unit Number] SET [Exca: Unit Sheet].TimePeriod = 'Neolithic' "
            sql = sql & "WHERE ((([Exca: Space Sheet].[Space number])=" & Me![Space number] & "));"
            DoCmd.RunSQL sql
            
        
        End If
        

    Else
        Dim response
        response = MsgBox("Do you wish the Hodder Level field to be blank?", vbYesNo + vbQuestion, "Action confirmation")
        If response = vbYes Then
            Me![txtHodderPhase] = ""
        End If
        
    End If
    Me![cboHodderPhase] = ""
Exit Sub

err_phase:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub chkExternal_AfterUpdate()
On Error GoTo err_chk
'new 2009 - phasing list here only if external space
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
End If

'new 2010
If Me![ExternalToBuilding] = True Then
    Me![ExternalSpaceInfillingProcess].Enabled = True
    Me![cboOutline].Enabled = True
Else
    Me![ExternalSpaceInfillingProcess].Enabled = False
    Me![cboOutline].Enabled = False
End If

If Me![ExternalToBuilding] = True Then
    Me![ExternalSpaceInfillingProcess].Enabled = True
Else
    If Me![ExternalSpaceInfillingProcess] <> "" Or Me![cboOutline] <> "" Then
        Dim resp
        resp = MsgBox("Only external spaces can have an infilling process / outline, this value will therefore be removed. You must assign it to the building. Are you sure you wish to make this change?", vbYesNo + vbExclamation, "This affects infilling process")
        If resp = vbYes Then
            Me![ExternalSpaceInfillingProcess] = ""
            Me![ExternalSpaceInfillingProcess].Enabled = False
            Me![cboOutline] = ""
            Me![cboOutline].Enabled = False
        Else
            Me![ExternalToBuilding] = True
        End If
    End If
End If
Exit Sub

err_chk:
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
    DoCmd.GoToControl "Space Number"
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
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Field26]
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
        'decided against dialog as you can open other forms from building form and they would appear beneath it
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
    'myq1.sql = "sp_Portfolio_GetUnitFieldID " & Me![Year] - year here is random as there isn't one
    myq1.sql = "sp_Portfolio_GetSpaceFieldID_2009 2009"
    
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
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Space number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, 2009
            
        Else
            'database is running remotely must access images via internet
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Space number])
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

Private Sub cmdHelp_Click()
On Error GoTo Err_cmdHelp_Click

'either pop up a window or have a message
MsgBox "A help message to explain the post excavation fields will appear soon", vbInformation, "Help"
    

Exit_cmdHelp_Click:
    Exit Sub

Err_cmdHelp_Click:
    Resume Exit_cmdHelp_Click
    
End Sub

Private Sub cmdPrintSpaceSheet_Click()
'new for 2009
On Error GoTo err_cmdSpace

    Dim resp, both
    both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        DoCmd.OpenReport "R_SpaceSheet", acViewPreview, , "[Space Number] = " & Me![Space number]
        If both = vbYes Then DoCmd.OpenReport "R_Units_in_Spaces", acViewPreview, , "[In_Space] = " & Me![Space number]

Exit Sub

err_cmdSpace:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReportProblem_Click()
'bring up a popup to allow user to report a problem
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "space;" & Me![Space number]

Exit Sub

err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewSpacesketch_Click()
On Error GoTo err_ViewSpacesketch_Click
    Dim Path
    Dim fname, newfile
    
    'check if can find sketch image
    'using global constanst sktechpath Declared in globals-shared
    'path = "\\catal\Site_Sketches\Features\Sketches"
    Path = sketchpath2015 & "spaces\sketches\"
    Path = Path & "S" & Me![Space number] & "*" & ".jpg"
    
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        newfile = fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "spaces\sketches\" & newfile
    
    If Dir(Path) = "" Then
        'directory not exist
        MsgBox "The sketch plan of this space has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_spacesketch", acNormal, , , acFormReadOnly, , Me![Space number]
    End If
 
Exit Sub

err_ViewSpacesketch_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
' removed open excavation form (menu) as this form
' can now be called by other forms
' SAJ v9.1
    Dim stDocName As String
    Dim stLinkCriteria As String

    'stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Space Sheet"
End Sub



Private Sub Field26_AfterUpdate()
'TALK TO SHAHINA ABOUT THE WAY THIS SHOULD WORK - SHE SAYS LEAVE IT EDITABLE AT PRESENT
'Dim checkBuildingArea
'If Me![Field26].OldValue <> "" And Not IsNull(Me![Building]) Then
'    checkBuildingArea = DLookup("[Area]", "[Exca: Building Details]", "[Number] = " & Me![Building] & "")
'
'    If Not IsNull(checkBuildingArea) Then
'        If checkBuildingArea <> Me![Field26] Then
'            MsgBox "This area alteration means this Space is recorded as being in a different Area to the Building number " & Me![Building]
'        End If
'    End If
'End If
'new 2008 - the mound wasn't getting updated!
Me![Mound] = Me!Field26.Column(1)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'***************************************************************************
' If no space number entered key violation occurs - intercept and provide
' nice msg, plus way to escape msg loop
'
' SAJ v9.1
'***************************************************************************
On Error GoTo err_Form_BeforeUpdate

If IsNull(Me![Space number] And (Not IsNull(Me![Field26]) Or Not IsNull(Me![Building]) Or Not IsNull(Me![Level]) Or (Not IsNull(Me![Description]) And Me![Description] <> ""))) Then
    MsgBox "You must enter a space number otherwise the record cannot be saved." & Chr(13) & Chr(13) & "If you wish to cancel this record being entered and start again completely press ESC", vbInformation, "Incomplete data"
    Cancel = True
    DoCmd.GoToControl "Space Number"
ElseIf IsNull(Me![Space number]) And IsNull(Me![Field26]) And IsNull(Me![Building]) And IsNull(Me![Level]) And (IsNull(Me![Description]) Or Me![Description] = "") Then
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
'***********************************************************************
' Things to check for each record: if no building number then disable goto button
' New requirement that space number cannot be edited after entry. This
' can only be done by an administrator so lock field here
' SAJ v9.1, v9.2
'***********************************************************************
On Error GoTo err_Form_Current

    If Me![Building] = "" Or IsNull(Me![Building]) Then
        Me![cmdGoToBuilding].Enabled = False
    Else
        Me![cmdGoToBuilding].Enabled = True
    End If
    
    If Me![Space number] <> "" Then
    'space number exists, lock field
        Me![Space number].Locked = True
        Me![Space number].Enabled = False
        Me![Space number].BackColor = Me.Section(0).BackColor
        
        Me![Building].SetFocus
    Else
    'no space number - new record allow entry
        Me![Space number].Locked = False
        Me![Space number].Enabled = True
        Me![Space number].BackColor = 16777215
        
        Me![Space number].SetFocus
    End If
    
    'new for v9.2 intro of LevelLOV and certain/uncertain option
    ''MsgBox Me![LevelCertain]
    If Me![MellaartLevelCertain] = True Then
        'if level is certain then enable level list
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    Else
        'level is uncertain, allow edits to level start and end but not level
        Me![Level].Enabled = False
        Me![cboUncertainLevelStart].Enabled = True
        Me![cboUnCertainLevelEnd].Enabled = True
    End If
    
'see unit sheet oncurrent for explaination of why altered - image form now does the check
''    'new 2009
'''find out is any images available
Dim imageCount, Imgcaption
''
''Dim mydb As DAO.Database
''Dim myq1 As QueryDef, connStr
''    Set mydb = CurrentDb
''    Set myq1 = mydb.CreateQueryDef("")
''    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
''    myq1.ReturnsRecords = True
''    myq1.sql = "sp_Portfolio_CountImagesForSpace_2009 '" & Me![Space number] & "', ''"
''
''    Dim myrs As Recordset
''    Set myrs = myq1.OpenRecordset
''    ''MsgBox myrs.Fields(0).Value
''    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
''           imageCount = 0
''    Else
''        imageCount = myrs.Fields(0).Value
''   End If

''myrs.close
''Set myrs = Nothing

backhere:
''myq1.close
''Set myq1 = Nothing
''mydb.close
''Set mydb = Nothing
    
''If imageCount > 0 Then
''    Imgcaption = imageCount
''    If imageCount = 1 Then
''        Imgcaption = Imgcaption & " Image to Display"
''    Else
''       Imgcaption = Imgcaption & " Images to Display"
''    End If
''    Me![cmdGoToImage].Caption = Imgcaption
''    Me![cmdGoToImage].Enabled = True
''Else
''    Me![cmdGoToImage].Caption = "No Image to Display"
''    Me![cmdGoToImage].Enabled = False
''End If

Imgcaption = "Images of Space"
Me![cmdGoToImage].Caption = Imgcaption
Me![cmdGoToImage].Enabled = True

'''OFFSITE 2009 - ignore photos and sketches offsite
'''JUST TAKE THIS LINE OUT ON SITE TO RETRIEVE FUNCTIONALITY
''Me![cmdGoToImage].Enabled = False

'new 2009 - phasing list here only if external space
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
End If

'new 2010
If Me![ExternalToBuilding] = True Then
    Me![ExternalSpaceInfillingProcess].Enabled = True
    Me![cboOutline].Enabled = True
Else
    Me![ExternalSpaceInfillingProcess].Enabled = False
    Me![cboOutline].Enabled = False
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

Private Sub Form_Open(Cancel As Integer)
'******************************************************
' Chck status of form on open to see if filtered or locked
' so determine which controls to enable
'
' SAJ v9.1
'******************************************************
On Error GoTo err_Form_Open

If Me.FilterOn = True Or Me.AllowEdits = False Then
    'disable find and add new in this instance find will not work
    'and should not be able to add records
    Me![cboFindSpace].Enabled = False
    Me![cmdAddNew].Enabled = False
    Me.AllowAdditions = False
Else
    'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
    DoCmd.GoToControl "cboFindSpace"
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

'new 2009 - repeat from oncurrent but seems to be needed here to prevent first record being
'editable - phasing list here only if external space
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
End If

'new 2010
If Me![ExternalToBuilding] = True Then
    Me![ExternalSpaceInfillingProcess].Enabled = True
    Me![cboOutline].Enabled = True
Else
    Me![ExternalSpaceInfillingProcess].Enabled = False
    Me![cboOutline].Enabled = False
End If

Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub frmLevelCertain_AfterUpdate()
'*************************************************************************
' New in v9.2 - levels can now be certain or uncertain. Uncertain levels
' can have a start and end entry. If the user changes their mind the value must
' be moved between these lists.
'SAJ v9.2
'*************************************************************************
On Error GoTo err_frmLevelCertain_AfterUpdate
Dim retval

If Me![frmLevelCertain] = -1 Then
    'user has selected level as certain, allow them to choose a level from the list
    'and disable the start end combos
    If Me![cboUncertainLevelStart] <> "" And Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the Start Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then
            Me![Level] = Me![cboUncertainLevelStart]
        Else
            retval = MsgBox("Do you want the End Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
            If retval = vbYes Then
                Me![Level] = Me![cboUnCertainLevelEnd]
            Else
                retval = MsgBox("The start and end level fields will now be cleared and you will have to select the Certain level from that list. Are you sure you want to continue?", vbQuestion + vbYesNo, "Uncertain Levels will be cleared")
                If retval = vbYes Then
                    Me![cboUncertainLevelStart] = ""
                    Me![cboUnCertainLevelEnd] = ""
                Else
                    Me![frmLevelCertain] = 0
                End If
            End If
        End If
    ElseIf Me![cboUncertainLevelStart] <> "" Then
        retval = MsgBox("Do you want the Start Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then Me![Level] = Me![cboUncertainLevelStart]
        Me![cboUncertainLevelStart] = ""
    ElseIf Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the End Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then Me![Level] = Me![cboUnCertainLevelEnd]
        Me![cboUnCertainLevelEnd] = ""
    End If
    
    If Me![frmLevelCertain] = -1 Then 'they have decide not to change their mind
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    End If
Else
    'level uncertain so allow start end but not certain level
    Me![Level].Enabled = False
    If Me![Level] <> "" Then
        Me![cboUncertainLevelStart] = Me![Level]
        Me![Level] = ""
    End If
    Me![cboUncertainLevelStart].Enabled = True
    Me![cboUnCertainLevelEnd].Enabled = True
End If
Exit Sub

err_frmLevelCertain_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Level_NotInList(NewData As String, response As Integer)
'***********************************************************************
' Intro of a validity check to make users a little more aware of the data
' they are entering here. The combo here is trying to prevent different entries
' that represent the same thing. Users are allowed to enter new values but just made aware
'
' SAJ v9.1
' SAJ v9.2 - now the list is only updateable by the administrator via Admin interface
' rowsource of this combo changed from:
' SELECT [Qry:SpaceSheet_Distinct_Levels].Level FROM [Qry:SpaceSheet_Distinct_Levels];
' to
' Exca:LevelLOV
'***********************************************************************
'On Error GoTo err_Level_NotInList
'
'Dim retVal
'retVal = MsgBox("This level has not been used on this screen before, are you sure you wish to add it to the list?", vbInformation + vbYesNo, "New Level Value")
'If retVal = vbYes Then
'    'allow value, as this is distinct query based list we must save the record
'    'first but need to turn off limittolist first to be able to do so an alternative
'    'way to do this would be to dlookup on entry when not limited
'    'to list but this method is quicker (but messier) as not require DB lookup 1st
'    Response = acDataErrContinue
'    Me![Level].LimitToList = False 'turn off limit to list so record can be saved
'    DoCmd.RunCommand acCmdSaveRecord 'save rec
'    Me![Level].Requery 'requery combo to get new value in list
'    Me![Level].LimitToList = True 'put back on limit to list
'Else
'    'no leave it so they can edit it
'    Response = acDataErrContinue
'End If
'Exit Sub
'
'err_Level_NotInList:
'    Call General_Error_Trap
'    Exit Sub
'
End Sub

Private Sub Space_number_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure duplicate space numbers not entered
' which would result in nasty key violation msg back from sql server
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Space_Number_AfterUpdate

Dim checknum

If Me![Space number] <> "" Then
    'check that space num not exist
    checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![Space number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but this Space Number already exists, please enter another number.", vbInformation, "Duplicate Space Number"
        
        If Not IsNull(Me![Space number].OldValue) Then
            'return field to old value if there was one
            Me![Space number] = Me![Space number].OldValue
        Else
            'if its a new record and no oldval (ie: null error is try to set it to oldval)
            'then just undo entry - me![number].undo does not undo this field, only
            'me.undo will but this removes all edits in all fields so must gather them first (!!)
            Dim currBuild, currarea, currdesc, currLevel
            currBuild = Me![Building]
            currarea = Me![Field26]
            currLevel = Me![Level]
            currdesc = Me![Description]
            DoCmd.GoToControl "Space Number"
            Me.Undo
            
            'reset all fields, for some reason if description is null (and currdesc is null)
            'it won't set it back to null, instead "". This throws out the form_beforeupdate
            'code, to ensure this not happen have added the if not isnull check, so only updates
            'field if there was an original value
            If Not IsNull(currBuild) Then Me![Building] = currBuild
            If Not IsNull(currarea) Then Me![Field26] = currarea
            If Not IsNull(currLevel) Then Me![Level] = currLevel
            If Not IsNull(currdesc) Then Me![Description] = currdesc
            
            'for some reason have to send focus to another field to bring it back
            'otherwise goes onto area- setfocus not work either
            DoCmd.GoToControl "Description"
            DoCmd.GoToControl "Space Number"
        End If
    End If
End If

Exit Sub

err_Space_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
