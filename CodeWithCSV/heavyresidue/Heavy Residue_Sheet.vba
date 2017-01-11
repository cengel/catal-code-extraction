Option Compare Database
Option Explicit
Private Sub Update_GID()
Me![GID] = Me![Unit] & "." & Me![Sample] & "." & Me![Flot Number]
End Sub

Private Sub cboFilterUnit_AfterUpdate()
'filter - new 2011
On Error GoTo err_filter

    If Me![cboFilterUnit] <> "" Then
        Me.Filter = "[Unit] = " & Me![cboFilterUnit]
        Me.FilterOn = True
        Me![cboFilterUnit] = ""
        Me![cmdRemoveFilter].Visible = True
    End If

Exit Sub

err_filter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFilterUnit_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - new 2011
On Error GoTo err_cbofilterNot

    MsgBox "Sorry this Unit does not exist in this database yet", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFilterUnit].Undo
Exit Sub

err_cbofilterNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
'new 2011
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    DoCmd.GoToControl "GID"
    DoCmd.FindRecord Me![cboFind]
    ''DoCmd.GoToControl "Analyst"
    Me![cboFind] = ""
End If


Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop - new 2011
On Error GoTo err_cbofindNot

    MsgBox "Sorry this GID does not exist in the database", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub chkMerged_AfterUpdate()
'new 2011
On Error GoTo err_chk

If Me!chkMerged = True Then
    DoCmd.OpenForm "frm_pop_mergedFlot", , , , , , Me![Flot Sample number]
    Me!cmdMergeFlot.Visible = True
Else
    'is false must check if anything exists
    Dim checkit
    checkit = DLookup("[FlotRecordedInHR]", "[Heavy Residue: Flot Merge Log]", "[FlotRecordedInHR] = " & Me![Flot Sample number])
    If checkit <> "" Then
        'there are merge records
        MsgBox "Flot numbers are recorded as being merged into this one. You cannot uncheck this box until this asociation is removed." & Chr(13) & Chr(13) & "Use the button to the right of the check box and delete the numbers there", vbExclamation, "Action Cancelled"
        Me!chkMerged = True
    Else
        Me!cmdMergeFlot.Visible = False
    End If
    
End If
Exit Sub

err_chk:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdAddNew_Click()
'********************************************************************
' Create new record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgonew_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acNewRec
    ''DoCmd.GoToControl Me![frm_subform_basic].Form![Analyst].Name
    DoCmd.GoToControl Me![Unit].Name
    Exit Sub

Err_cmdgonew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'new 2011 - control the delete of a record to ensure both tables are clear
On Error GoTo err_del

Dim response
    response = MsgBox("Do you really want to remove GID " & Me!GID & " and all its related material records from your database?", vbYesNo + vbQuestion, "Remove Record")
    If response = vbYes Then
        Dim sql
        sql = "Delete FROM [Heavy Residue: Basic Details] WHERE [GID] = '" & Me![GID] & "';"
        DoCmd.RunSQL sql
        
        sql = "Delete from [Heavy Residue: Material] WHERE [GID] = '" & Me![GID] & "';"
        DoCmd.RunSQL sql
        
        sql = "Delete from [Heavy Residue: Flot Merge Log] WHERE [FlotRecordedInHr] = " & Me![Flot Sample number] & ";"
        DoCmd.RunSQL sql
        
        Me.Requery
        MsgBox "Deletion completed", vbInformation, "Done"
        
        Me![cboFind].Requery
        Me![cboFilterUnit].Requery
        
    End If
Exit Sub

err_del:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_cmdgofirst_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoLast_Click()
'********************************************************************
' Go to last record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoLast_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_cmdgoLast_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoNext_Click()
'********************************************************************
' Go to next record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoNext_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_cmdgoNext_Click:
    If Err.Number = 2105 Then
        MsgBox "No more records to show", vbInformation, "End of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdGoPrev_Click()
'********************************************************************
' Go to previous record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoprevious_Click

    ''DoCmd.GoToControl Me![frm_subform_basic].Name
    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_cmdgoprevious_Click:
    If Err.Number = 2105 Then
        MsgBox "Already at the begining of the recordset", vbInformation, "Begining of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdMergeFlot_Click()
'new 2011
On Error GoTo Err_Command62_Click

    DoCmd.OpenForm "frm_pop_MergedFlot", , , , , , Me![Flot Sample number]

    Exit Sub

Err_Command62_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdRemoveFilter_Click()
'remove unit filter - new 2011
On Error GoTo err_Removefilter

    Me![cboFilterUnit] = ""
    Me.Filter = ""
    Me.FilterOn = False
    
    DoCmd.GoToControl "cboFind"
    Me![cmdRemoveFilter].Visible = False
   

Exit Sub

err_Removefilter:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Flot_Sample_number_AfterUpdate()
'new season 2006 - get soil vol from flot log
On Error GoTo err_sample

Update_GID

'august 2008 check this record exists in Bots - request from Betsa - just a prompt, don't stop entry
'NOTE THIS QUERY (Q_Bots_GID)IS HIDDEN IN THE QUERY LIST TO PREVENT ACCIDENTAL DELETION - saj
Dim getBots, getBotsUnit, getBotsSample, getBotsFlot
'first see if gid matches as a whole in bots db
getBots = DLookup("[GID]", "Q_Bots_GID", "[GID] = '" & Me![GID] & "'")
If IsNull(getBots) Then
    'no exact GID match - so does flot number exist and with what unit?
    getBotsFlot = DLookup("[Flot Number]", "view_ArchaeoBotany_Flot_Log", "[Flot Number] = " & Me![Flot Sample number])
    If IsNull(getBotsFlot) Then
        'flot not present at all in Bots
        MsgBox "This flotation number cannot be found in the Botany database. Please check it is correct", vbExclamation, "Data mismatch"
    Else
        'flot exists so check unit number, see if it matches one entered here
        getBotsUnit = DLookup("[Unit Number]", "view_ArchaeoBotany_Flot_Log", "[Flot Number] = " & Me![Flot Sample number])
        If getBotsUnit <> Me![Unit] Then
            'the unit number here does not match entry in bots db for this flot number
            MsgBox "This Flot number is entered into the Bots database for Unit " & getBotsUnit & " not the unit you have entered. Please check it.", vbExclamation, "Data Mismatch"
        Else
            'the unit number matches for this flot but there is still a problem as the GID did not, this means the sample number must be wrong
            MsgBox "This GID does not match a GID in the Bots database, the sample number appears to be incorrect. Please check it.", vbExclamation, "Data Mismatch"
        End If
    End If
    
End If

Dim getVol
getVol = DLookup("[Soil Volume]", "view_ArchaeoBotany_Flot_Log", "[Flot Number] = " & Me![Flot Sample number])
If Not IsNull(getVol) Then
    Me![Flot Volume] = getVol
End If

Me![cboFind].Requery
Me![cboFilterUnit].Requery

Exit Sub

err_sample:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Flot_Sample_number_Change()
'comment out saj
'Update_GID
End Sub


Private Sub Flot_Sample_number_Enter()
'SAJ before versioning - this causes sql update error to be returned to user even
'they have not tried to edit anything, most confusing and unnecessary so removed
' 11/01/05
'Update_GID
End Sub


Private Sub Form_Current()
'new 2011
On Error GoTo err_current

    Me![cboFind].Requery
    Me![cboFilterUnit].Requery
    
    If Me![chkMerged] = True Then
        Me!cmdMergeFlot.Visible = True
    Else
        Me!cmdMergeFlot.Visible = False
    End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Sample_AfterUpdate()
Update_GID
End Sub

Private Sub Sample_Change()
'comment out saj
'Update_GID
End Sub


Private Sub Sample_Enter()
'SAJ before versioning - this causes sql update error to be returned to user even
'they have not tried to edit anything, most confusing and unnecessary so removed
' 11/01/05
'Update_GID
End Sub


Private Sub Unit_AfterUpdate()
Update_GID
End Sub

Private Sub Unit_Change()
'Update_GID
End Sub


Private Sub Unit_Enter()
'SAJ before versioning - this causes sql update error to be returned to user even
'they have not tried to edit anything, most confusing and unnecessary so removed
' 11/01/05
'Update_GID
End Sub



