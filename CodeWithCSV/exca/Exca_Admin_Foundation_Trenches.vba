Option Compare Database   'Use database order for string comparisons
Option Explicit







Private Sub cboFindFT_AfterUpdate()
'********************************************
'Find the selected space number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindFT_AfterUpdate

    If Me![cboFindFT] <> "" Then
        'if number is disabled then must enable if for the search
        'it will be reset to disabled by the code in form oncurrent
        If Me![FTName].Enabled = False Then Me![FTName].Enabled = True
        DoCmd.GoToControl "FTName"
        DoCmd.FindRecord Me![cboFindFT]
        Me![cboFindFT] = ""
        '2009 dont move focus back so not cause accidental overwrite
        DoCmd.GoToControl "cboFindFT"
    End If
Exit Sub

err_cboFindFT_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
'********************************************
'This is the trowel button to close the form
' Error trap added v9.1
'********************************************
On Error GoTo err_Excavation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, Me.Name
Exit Sub

err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Form_Current()

If Me![FTName] <> "" Then
    'FT exists, lock field
        Me![FTName].Locked = True
        Me![FTName].Enabled = False
        Me![FTName].BackColor = Me.Section(0).BackColor
        
        
    Else
    'no FT - new record allow entry
        Me![FTName].Locked = False
        Me![FTName].Enabled = True
        Me![FTName].BackColor = 16777215
        
        Me![FTName].SetFocus
    End If
'new for v11.1 intro of LevelLOV and certain/uncertain option
    ''MsgBox Me![LevelCertain]
    If Me![LevelCertain] = True Then
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
End Sub

Private Sub Form_Open(Cancel As Integer)
'*************************************************************
' Check permissions of user to define how to display form
' v9.1 SAJ
'*************************************************************
On Error GoTo err_Form_Open

If Me.FilterOn = True Or Me.AllowEdits = False Then
    'disable find and add new in this instance find will not work
    'and should not be able to add records
    Me![cboFindFT].Enabled = False
    ''Me![cmdAddNew].Enabled = False
    Me.AllowAdditions = False
Else
    'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
    DoCmd.GoToControl "cboFindFT"
End If
   
'now sort out view depending on permissions
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN") And (Me.AllowAdditions = True Or Me.AllowDeletions = True Or Me.AllowEdits = True) Then
    'SAJ season 2006 - don't allow deletions from this screen
    ToggleFormReadOnly Me, False, "NoDeletions"
Else
    ToggleFormReadOnly Me, True
    ''Me![cmdAddNew].Enabled = False
End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub frmLevelCertain_AfterUpdate()
'*************************************************************************
' levels can now be certain or uncertain. Uncertain levels
' can have a start and end entry. If the user changes their mind the value must
' be moved between these lists.
'SAJ v11.1
'*************************************************************************
On Error GoTo err_frmLevelCertain_AfterUpdate
Dim retval

If Me![frmLevelCertain] = -1 Then
    'user has selected level as certain, allow them to choose a level from the list
    'and disable the start end combos
    If Me![cboUncertainLevelStart] <> "" And Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the Start Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then
            Me![Level] = Me![cboUncertainLevelStart]
        Else
            retval = MsgBox("Do you want the End Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
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
        retval = MsgBox("Do you want the Start Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then Me![Level] = Me![cboUncertainLevelStart]
        Me![cboUncertainLevelStart] = ""
    ElseIf Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the End Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
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

Private Sub FTName_AfterUpdate()
'*************************************************
' As foundation trenches names are actually stored in the Unit sheet table
' don't allow them to altered after a unit sheet has been allocated to the name
'
' On screen msg tells user to contact DBA who can explain
' what a problem the change could have, unless it really is just
' an error done very recently.
'
' SAJ v11.1
'*************************************************
On Error GoTo err_FTNAME

Dim resp

If Not IsNull(Me!FTName.OldValue) Then
    'a change - check FT name not used in unit sheet
    resp = DLookup("[FoundationTrench]", "[Exca: Unit Sheet]", "[FoundationTrench] = '" & Me!FTName.OldValue & "' AND [Area] = '" & Me![cboArea] & "'")
    If Not IsNull(resp) Then
        MsgBox "This FT is assocated with a Unit so the name cannot be altered. Please enter this change as a new FT name and then re-allocate the units to the new record", vbExclamation, "Changed Cancelled"
        Me!FTName = Me!FTName.OldValue
    End If
End If
Exit Sub

err_FTNAME:
    Call General_Error_Trap
    Exit Sub


End Sub
