Option Compare Database
Option Explicit
Private Sub Update_GID()
If Not IsNull(Me![Unit]) And Not IsNull(Me![Sample]) And Not IsNull(Me![Flot_Number]) Then
    Me![GID] = Me![Unit] & "." & Me![Sample] & "." & Me![Flot_Number]
Else
    Debug.Print "GID is not valid"
End If
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



Private Sub charcoal_weight_AfterUpdate()
On Error GoTo err_density

If [subfrm_HeavyResidue_FlotVolume].[Form]![Soil Volume].Value <> 0 Then
Me![density].Value = Round(Me![charcoal_weight].Value / [subfrm_HeavyResidue_FlotVolume].[Form]![Flot Volume].Value, 2)
End If

Exit Sub

err_density:
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
        sql = "Delete FROM [Anthracology: Basic Details] WHERE [GID] = '" & Me![GID] & "';"
        DoCmd.RunSQL sql
        
        sql = "Delete from [Anthracology: Dendro] WHERE [GID] = '" & Me![GID] & "';"
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


Exit Sub

err_sample:
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



