Option Compare Database
Option Explicit
Private Sub UpdateFullConservRef()
'Season 2006 - update the full conservation ref to keep it in sync with the
'year and Id fields
'SAJ
On Error GoTo err_UpdateFullCOnservRef
Dim theYear, theID
    'txtConservationRef_Year
    'txtConservationRef_ID
    If Not IsNull(txtConservationRef_Year) And Not IsNull(txtConservationRef_ID) Then
        
        theYear = Right(Me![txtConservationRef_Year], 2)
        
        If Len(Me![txtConservationRef_ID]) = 1 Then
            theID = "00" & Me![txtConservationRef_ID]
        ElseIf Len(Me![txtConservationRef_ID]) = 2 Then
            theID = "0" & Me![txtConservationRef_ID]
        Else
            theID = Me![txtConservationRef_ID]
        End If
        Me![FullConservation_Ref] = theYear & "." & theID
    
    End If

Exit Sub

err_UpdateFullCOnservRef:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboFind_AfterUpdate()
'find conservation ref
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
        Me![txtFullRef].Enabled = True
        
        DoCmd.GoToControl "txtFullRef"
        DoCmd.FindRecord Me![cboFind]
        
        DoCmd.GoToControl "cboFind"
        Me![txtFullRef].Enabled = False
    End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub

End Sub



Private Sub cboPackingList_AfterUpdate()
'new for season 2006
'user selects treatment from the list and it fills the text into the treatment field
On Error GoTo err_pack

    If Me![cboPackingList] <> "" Then
        Me![Packing] = Me![Packing] & " " & Me![cboPackingList]
        Me![cboPackingList] = ""
    End If
Exit Sub

err_pack:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboRelatedToID_AfterUpdate()
'new for seaon 2006
'enable subform
On Error GoTo err_cboRelatedToID

    Me![frm_subform_conservation_ref_relatedto].Locked = False
    Me![frm_subform_conservation_ref_relatedto].Enabled = True
    DoCmd.RunCommand acCmdSaveRecord
    Me![frm_subform_conservation_ref_relatedto].Form.Requery
    'Me![frm_subform_conservation_ref_relatedto].Form.Refresh
    'Me![frm_subform_conservation_ref_relatedto].Form.Repaint
    
    'if the relationship is to an object/s show cmdRange button on the relatedto subform
    'so a range of X numbers can be entered automatically
    If Me![cboRelatedToID] = 2 Then
        Me![frm_subform_conservation_ref_relatedto].Form![cmdRange].Visible = True
    Else
        Me![frm_subform_conservation_ref_relatedto].Form![cmdRange].Visible = False
    End If
Exit Sub

err_cboRelatedToID:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboTreatmentList_AfterUpdate()
'new for season 2006
'user selects treatment from the list and it fills the text into the treatment field
On Error GoTo err_treat

    If Me![cboTreatmentList] <> "" Then
        Me![Treatment] = Me![Treatment] & " " & Me![cboTreatmentList]
        Me![cboTreatmentList] = ""
    End If
Exit Sub

err_treat:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'********************************************
'Add a new record
'
'SAJ
'********************************************
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "txtConservationRef_Year"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAdminMenu_Click()
'new for season 2006 - open admin menu
On Error GoTo err_cmdAdmin

    'double check is admin
    Dim permiss
    permiss = GetGeneralPermissions
    
    If permiss <> "ADMIN" Then
        MsgBox "You do not have permission to open this screen", vbInformation, "Permission Denied"
    Else
        DoCmd.OpenForm "frm_Admin_Menu"
        
    End If
    
Exit Sub

err_cmdAdmin:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdDelete_Click()
'new for season 2006 - delete here so can clean up sub tables
On Error GoTo err_cmdDelete

    'double check is admin
    Dim permiss
    permiss = GetGeneralPermissions
    
    If permiss <> "ADMIN" Then
        MsgBox "You do not have permission to delete records. Contact your supervisor.", vbInformation, "Permission Denied"
    Else
        Dim retVal, sql
        retVal = MsgBox("Really delete conservation record: " & Me![txtConservationRef_Year] & "." & txtConservationRef_ID & "?", vbCritical + vbYesNoCancel, "Confirm Delete")
        If retVal = vbYes Then
            sql = "Delete from [Conservation_ConservRef_RelatedTo] WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
            DoCmd.RunSQL sql
            
            sql = "Delete from [Conservation_Photos] WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
            DoCmd.RunSQL sql
        
            sql = "Delete from [Conservation_Basic_Record] WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
            DoCmd.RunSQL sql
            
            Me.Requery
            DoCmd.GoToRecord acActiveDataObject, , acLast
        End If
    End If
    
Exit Sub

err_cmdDelete:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdGoToImage_Click()
'********************************************************************
' New button for season 2006 which allows any available images to be
' displayed - links to the view on Image_Metadata table that has been exported
' from Portfolio
' SAJ
'********************************************************************
On Error GoTo err_cmdGoToImage_Click

'DoCmd.OpenForm "Image_Display", acNormal, , "[Lab Record ID] = '" & Me![FullConservation_Ref] & "'", acFormReadOnly, acDialog
    
Dim mydb As DAO.Database
Dim tmptable As TableDef, tblConn, I, msg, LabTeamID, LabRecordID
Set mydb = CurrentDb

    'new 2009 - get back the portfolio lab team id and lab record id as this might change as portfolio recatalogues
    'this code is the same for all labs
    Dim myq1 As QueryDef, connStr
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    myq1.sql = "sp_Portfolio_Return_Lab_Team_FieldIDs"
    
    Dim myrs As Recordset
    Set myrs = myq1.OpenRecordset
    ''MsgBox myrs.Fields(0).Value
    If myrs!LabRecordID.Value = "" Or myrs!LabRecordID.Value = 0 Then
        LabRecordID = 0
    Else
        LabRecordID = myrs!LabRecordID.Value
    End If
        
    If myrs!LabTeam.Value = "" Or myrs!LabTeam.Value = 0 Then
        LabTeamID = 0
    Else
        LabTeamID = myrs!LabTeam.Value
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
            Dim formsql
           
           formsql = "select record_id, stringvalue from view_Portfolio_Lab_RecordID " & _
                     " where (field_id = " & LabRecordID & ") AND (stringvalue = '" & Me!FullConservation_Ref & "') " & _
                    " AND (record_id IN " & _
                    " (SELECT record_id FROM view_Portfolio_Lab_TeamID " & _
                    " WHERE field_id = " & LabTeamID & " AND stringvalue ='conservation')) "
           
            DoCmd.OpenForm "Image_Display", acNormal
           Forms!Image_Display.RecordSource = formsql
            'DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Unit Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
            
        Else
            'database is running remotely must access images via internet
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=unit&id=" & Me![Unit Number])
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

Private Sub cmdReport_Click()
'new for season 2006 - Anjas report has been altered and offered here on record
'by record basis
On Error GoTo err_cmdReport

    'DoCmd.OpenReport "Conserv: Full Printout", acViewPreview, , "[ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![ConservationRef_ID]
    DoCmd.OpenForm "FrmPopDataOutputOptions", acNormal, , , , acDialog, "Conservation_Basic_Record;" & Me![txtFullRef]

Exit Sub

err_cmdReport:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Command44_Click()
'***************************************************
' Existing close button revamped - text changed from close
' to trowel as in rest of database
'
' SAJ
'***************************************************
On Error GoTo err_Command44_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    DoCmd.Close acForm, "Conserv: Basic Record"
    
Exit_Command44_Click:
    Exit Sub

err_Command44_Click:
    MsgBox Err.Description
    Resume Exit_Command44_Click
End Sub

Private Sub Form_Current()
'new for season 2006 - set up form on open
'SAJ
On Error GoTo err_current
'If Not IsNull(Me![Finds: Basic Data.GID]) Then
'Me![finds].Enabled = True
'Else
'Me![finds].Enabled = False
'End If

Dim permiss
permiss = GetGeneralPermissions

'If Not IsNull(Me![txtConservationRef_Year]) And Not IsNull(Me![txtConservationRef_ID]) Then
'i have now set the default value of the ref year to this year so this must just check on the
'ref field as otherwise it will lock even new records!
If Not IsNull(Me![txtConservationRef_ID]) Then
    'existing record - lock conservation ref for everyone including admin but show admin a
    'an edit number button. If they simply edit the number the link to the sub tables is lost so
    'it must be done in code
    'do not existing conservation ref to be altered
    Me![txtConservationRef_Year].Enabled = False
    Me![txtConservationRef_Year].Locked = True
    Me![txtConservationRef_Year].BackColor = -2147483633
    Me![txtConservationRef_ID].Enabled = False
    Me![txtConservationRef_ID].Locked = True
    Me![txtConservationRef_ID].BackColor = -2147483633
    
    If permiss = "ADMIN" Then
        'enable conservatio ref to be alterd via edit button
        Me![cmdEditNum].Visible = True
    Else
        Me![cmdEditNum].Visible = False
    End If
    
'ElseIf IsNull(Me![txtConservationRef_Year]) And IsNull(Me![txtConservationRef_ID]) Then
ElseIf IsNull(Me![txtConservationRef_ID]) Then
    'a new record allow everyone except RO to update
    If permiss = "RO" Then
        'read only uses can do nothing
        Me![txtConservationRef_Year].Enabled = False
        Me![txtConservationRef_Year].Locked = True
        Me![txtConservationRef_Year].BackColor = -2147483633
        Me![txtConservationRef_ID].Enabled = False
        Me![txtConservationRef_ID].Locked = True
        Me![txtConservationRef_ID].BackColor = -2147483633
    Else 'enable new conservatio ref to be entered
        Me![txtConservationRef_Year].Enabled = True
        Me![txtConservationRef_Year].Locked = False
        Me![txtConservationRef_Year].BackColor = 16777215
        Me![txtConservationRef_ID].Enabled = True
        Me![txtConservationRef_ID].Locked = False
        Me![txtConservationRef_ID].BackColor = 16777215
    End If
End If


If IsNull(Me![cboRelatedToID]) Then
    Me![frm_subform_conservation_ref_relatedto].Locked = True
    Me![frm_subform_conservation_ref_relatedto].Enabled = False
    
    Me![cboRelatedToID].Enabled = True
    Me![cboRelatedToID].Locked = False
    Me![cboRelatedToID].BackColor = 16777215
Else
    Me![frm_subform_conservation_ref_relatedto].Locked = False
    Me![frm_subform_conservation_ref_relatedto].Enabled = True
    
    If permiss <> "ADMIN" Then
        Me![cboRelatedToID].Enabled = True
        Me![cboRelatedToID].Locked = False
        Me![cboRelatedToID].BackColor = -2147483633
    End If
End If


On Error Resume Next
'DC wants to see the excavation area the ref is related to quickly on screen. So this sets
'up the recordsource of a subform (frm_subform_exca_area) to link to the first excavation ID
'in the conservationref_relatedto list - not great but its a start on this functionality.
'This is only valid of records with a building/space/feature or unit number
'when the form is first opened the field below is not known and err 2455:
'expression that has an invalid reference to the property Form/Report is thrown so ignore
'it and carry on this will work for moving between records thereafter
If Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber] <> "" Then
    ''MsgBox Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber]

    If Me![RelatedToID] = 1 Or Me![RelatedToID] = 2 Or Me![RelatedToID] = 3 Then
        'related to an excavation id - building, feature, space or unit
        Me![frm_subform_exca_area].Visible = True
        If Me![frm_subform_conservation_ref_relatedto].Form![RelatedToSubTypeID] = 1 Then
            'building
            Me![frm_subform_exca_area].Form.RecordSource = "SELECT [Area], [Number] from [Exca: Building Details] WHERE [Number] = " & Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber]
            'MsgBox Me![frm_subform_exca_area].Form.RecordSource
        ElseIf Me![frm_subform_conservation_ref_relatedto].Form![RelatedToSubTypeID] = 2 Then
            'space
            Me![frm_subform_exca_area].Form.RecordSource = "SELECT [Area], [Space Number] from [Exca: Space Sheet] WHERE [Space Number] = " & Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber]
            'MsgBox Me![frm_subform_exca_area].Form.RecordSource
        ElseIf Me![frm_subform_conservation_ref_relatedto].Form![RelatedToSubTypeID] = 3 Then
            'feature
            Me![frm_subform_exca_area].Form.RecordSource = "SELECT [Area], [Feature Number] from [Exca: Features] WHERE [Feature Number] = " & Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber]
            'MsgBox Me![frm_subform_exca_area].Form.RecordSource
        Else
            'unit
            Me![frm_subform_exca_area].Form.RecordSource = "SELECT [Area], [Unit Number] from [Exca: Unit Sheet with Relationships] WHERE [Unit Number] = " & Me![frm_subform_conservation_ref_relatedto].Form![txtExcavationIDNumber]
            'MsgBox Me![frm_subform_exca_area].Form.RecordSource
        End If
        Me![frm_subform_exca_area].Requery
    Else
        Me![frm_subform_exca_area].Visible = False
    End If
Else
    Me![frm_subform_exca_area].Visible = False
End If


'maintain error resume next status on this for on open reason (as above)
'if the relationship is to an object/s show cmdRange button on the relatedto subform
'so a range of X numbers can be entered automatically
If Me![RelatedToID] = 2 Then
    Me![frm_subform_conservation_ref_relatedto].Form![cmdRange].Visible = True
Else
    Me![frm_subform_conservation_ref_relatedto].Form![cmdRange].Visible = False
End If


On Error GoTo err_current


'COMMENT OUT FOR OFF SITE
 'new season 2006 - link in portfolio image info
 'find out is any images available
 Dim imageCount, Imgcaption
 '2006 link
 'imageCount = DCount("[Lab Record ID]", "view_Conservation_Image_Metadata", "[Lab Record ID] = '" & Me![FullConservation_Ref] & "'")
 Dim mydb As DAO.Database
 Dim myq1 As QueryDef, connStr
 
    Set mydb = CurrentDb

    Set myq1 = mydb.CreateQueryDef("")

    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    
    myq1.sql = "sp_Web_Images_Count_for_Specific_Lab_Team_Entity '" & Me![txtFullRef] & "', 'conservation'"
 

    Dim myrs As Recordset
    
    Set myrs = myq1.OpenRecordset
    
    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
           imageCount = 0
    Else

        imageCount = myrs.Fields(0).Value
   End If

 myrs.Close
 Set myrs = Nothing
 myq1.Close
 Set myq1 = Nothing
 mydb.Close
 Set mydb = Nothing



 If imageCount > 0 Then
    Imgcaption = imageCount
    If imageCount = 1 Then
        Imgcaption = Imgcaption & " Image to Display"
    Else
        Imgcaption = Imgcaption & " Images to Display"
    End If
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
 Else
    Me![cmdGoToImage].Caption = "No Image to Display"
    Me![cmdGoToImage].Enabled = False
 End If

'''OFF SITE - RESURRECT THIS CODE - COMMENT OUT ALL ABOVE FROM 'COMMENT OUT FOR OFF SITE
'''Me![cmdGoToImage].Caption = "No Image Link Offsite"
'''Me![cmdGoToImage].Enabled = False

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub

End Sub



Private Sub Form_Open(Cancel As Integer)
'new season 2006
On Error GoTo err_open


    DoCmd.GoToRecord acActiveDataObject, "Conserv: Basic record", acLast
    
    Dim permiss
    permiss = GetGeneralPermissions

    If permiss = "ADMIN" Then
        Me![cmdAdminMenu].Visible = True
        Me![cmdDelete].Visible = True
    Else
        Me![cmdAdminMenu].Visible = False
        Me![cmdDelete].Visible = False
    End If

    DoCmd.Maximize
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub find_Click()
'hijacking a button, that has been already here but invisible and seemingly unused'
'sending to the basic data list form for searching mutiple attributes - DL 2015

On Error GoTo err_find_Click

DoCmd.OpenForm "Conserv_BasicRecord_List", acNormal

Forms![Conserv_BasicRecord_List].Form![queryfullconserv].SetFocus

Exit_find_Click:
    Exit Sub

err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub

Sub close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
Sub finds_Click()
On Error GoTo Err_finds_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Finds: Basic Data"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_finds_Click:
    Exit Sub

Err_finds_Click:
    MsgBox Err.Description
    Resume Exit_finds_Click
    
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

On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
    
End Sub

Private Sub txtConservationRef_ID_AfterUpdate()
Call UpdateFullConservRef
End Sub

Private Sub txtConservationRef_Year_AfterUpdate()
'check 4 digit year entered
On Error GoTo err_txtRef

    If Len(Me![txtConservationRef_Year]) <> 4 Then
        'all years must be 4 digits
        MsgBox "All years must be entered as a four digit number eg: 2006, your entry has been altered to this year.", vbExclamation, "Entry Altered"
        Me![txtConservationRef_Year] = Year(Date)
'       MsgBox Me![txtConservationRef_Year].OldValue
    End If
    Call UpdateFullConservRef
    
Exit Sub

err_txtRef:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEditNum_Click()
'control conservation reference edits
On Error GoTo Err_cmdEditNum_Click

Dim retVal, newvalueYear, newvalueID, sql, fullref, theYear, theID

retVal = MsgBox("This button enables you to change the conservation reference number: " & Me![txtConservationRef_Year] & "." & Me![txtConservationRef_ID] & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbExclamation + vbYesNo, "Confirm Action")
If retVal = vbYes Then
    newvalueYear = InputBox("Please enter the new Conservation Year below (even if it matches the current year you must enter it):", "Conservation Year")
    If newvalueYear <> "" Then
        If Len(newvalueYear) <> 4 Then
            MsgBox "Sorry but this does not represent a valid year, please enter the year like this: 2006. Action cancelled", vbCritical, "Action Cancelled"
        Else
            newvalueID = InputBox("Please enter the new Conservation Number below (just the number part):", "Conservation Number")
            If newvalueID <> "" Then
                retVal = MsgBox("The existing reference: " & Me![txtConservationRef_Year] & "." & Me![txtConservationRef_ID] & " will now be altered to: " & newvalueYear & "." & newvalueID & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbExclamation + vbYesNo, "Confirm Action")
                If retVal = vbYes Then
                
                    theYear = Right(newvalueYear, 2)
                
                    If Len(newvalueID) = 1 Then
                        theID = "00" & newvalueID
                    ElseIf Len(newvalueID) = 2 Then
                        theID = "0" & newvalueID
                    Else
                        theID = newvalueID
                    End If
                
                    fullref = theYear & "." & theID
                    'first update an subtable references to this number
                    sql = "UPDATE [Conservation_ConservedBy] SET [ConservationRef_Year] = '" & newvalueYear & "', [ConservationRef_ID] = " & newvalueID & " WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
                    DoCmd.RunSQL sql
                
                    sql = "UPDATE [Conservation_ConservRef_RelatedTo] SET [ConservationRef_Year] = '" & newvalueYear & "', [ConservationRef_ID] = " & newvalueID & " WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
                    DoCmd.RunSQL sql
                
                    sql = "UPDATE [Conservation_Basic_Record] SET [ConservationRef_Year] = '" & newvalueYear & "', [ConservationRef_ID] = " & newvalueID & ", [FullConservation_Ref] = '" & fullref & "' WHERE [ConservationRef_Year] = '" & Me![txtConservationRef_Year] & "' AND [ConservationRef_ID] = " & Me![txtConservationRef_ID] & ";"
                    DoCmd.RunSQL sql
                
                    Me.Requery
                    Me![txtFullRef].Enabled = True
        
                    DoCmd.GoToControl "txtFullRef"
                    DoCmd.FindRecord fullref
                    Me![cboFind].Requery
                    DoCmd.GoToControl "cboFind"
                    Me![txtFullRef].Enabled = False
                Else
                    MsgBox "Action cancelled, no change has been made", vbCritical, "Action Cancelled"
                End If
        Else
            MsgBox "No Number entered, action cancelled", vbCritical, "Action Cancelled"
        End If
      End If
    Else
        MsgBox "No Year entered, action cancelled", vbCritical, "Action Cancelled"
    End If
    
End If


Exit_cmdEditNum_Click:
    Exit Sub

Err_cmdEditNum_Click:
    Call General_Error_Trap
    Resume Exit_cmdEditNum_Click
    
End Sub
