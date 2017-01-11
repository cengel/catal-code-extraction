Option Compare Database   'Use database order for string comparisons
Option Explicit



Private Sub Area_name_AfterUpdate()
'***********************************************************************************
' Area names are taken and stored in lots of other tables, therefore they should not
' be altered once they have been entered. Although there is an Area number field this is
' not stored by the other tables (it should be but thats a huge alteration).
'
' Area names will also be stored in tables that aren't used here so cannot be updated
' from here anyway so the safest thing to do is to tell the user (an Admin) the impact changing
' will have and let them put the area in the historical areas table or cancel the alteration
'
' originally had this code in beforeupdate as that allows cancel of input but it would
' not allow the recordset to be requeried after the old area record was deleted
'
' SAJ v9.1
'***********************************************************************************

On Error GoTo err_Area_name_afterupdate
Dim msg, retval

    If Not IsNull(Me![Area Name].OldValue) Or (Me![Area Name].OldValue <> Me![Area Name]) Then
        'check its not a new entry which means the oldvalue is null
        'and only act if its an edit that alters the text
        msg = "Sorry but edits to the Area name are not allowed. Area names are stored in many different tables "
        msg = msg & "and this name may have already been used." & Chr(13) & Chr(13)
        msg = msg & "It is possible to archive this as an old area name and add it to the list of Historical area names if you wish. This would "
        msg = msg & " take the format of:" & Chr(13) & Chr(13) & "Old Area name: " & Me![Area Name].OldValue & " now equates to " & Me![Area Name]
        msg = msg & Chr(13) & Chr(13) & "Press Cancel to return to the original Area name"
        msg = msg & Chr(13) & "or "
        msg = msg & "Press OK to change this area name and add the old one to the historical list. "
        
        retval = MsgBox(msg, vbExclamation + vbOKCancel + vbDefaultButton2, "Stop: Area names cannot just be altered")
        If retval = vbCancel Then
            ' Cancel = True 'used in before update
            Me![Area Name] = Me![Area Name].OldValue 'reset to oldval
        ElseIf retval = vbOK Then
            'need to archive this area off, this involves creating a new area in this RS for this new name, getting
            'its new number and then entering the old details along with the new details in the Historical table
            'to allow the 2 to be linked for archival purposes.
            Dim sql, sql2, sql3, newAreaNum
            sql = "INSERT INTO [Exca: Area Sheet] ([Area name], [Mound], [Description]) VALUES ('" & Me![Area Name] & "','" & Me![Mound] & "'," & IIf(IsNull(Me![Description]), "null", "'" & Me![Description] & "'") & ");"
            DoCmd.RunSQL sql
            newAreaNum = DLookup("[Area Number]", "Exca: Area Sheet", "[Area Name] = '" & Me![Area Name] & "'")
            
            sql2 = "INSERT INTO [Exca: Area_Historical_Names] (CurrentAreaNumber, CurrentAreaName, OldAreaNumber, OldAreaName, OldMound, OldDescription)"
            sql2 = sql2 & " VALUES (" & newAreaNum & ", '" & Me![Area Name] & "', " & Me![Area number] & ", '" & Me![Area Name].OldValue & "', '" & Me![Mound] & "', '" & Me![Description] & "');"
            DoCmd.RunSQL sql2
            
            'Cancel = False
            
            'sql3 = "DELETE * FROM [Exca: Area Sheet] WHERE [Area number] = " & Me![Area number]
            ' DoCmd.RunSQL sql3
            'can do delete with screen commands which prevents conflict error being returned to user
            DoCmd.RunCommand acCmdDeleteRecord
            Me.Requery 'get updated RS
            'move to last record as new area name will be the last record now
            DoCmd.GoToRecord acActiveDataObject, , acLast
        End If
    End If

Exit Sub

err_Area_name_afterupdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindArea_AfterUpdate()
'********************************************
'Find the selected building number from the list
'
'SAJ 2009
'********************************************
On Error GoTo err_cboFindArea_AfterUpdate

    If Me![cboFindArea] <> "" Then
        'for existing number the field with be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![Area Name].Enabled = False Then Me![Area Name].Enabled = True
        DoCmd.GoToControl "Area Name"
        DoCmd.FindRecord Me![cboFindArea]
        Me![cboFindArea] = ""
    End If
Exit Sub

err_cboFindArea_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboFindArea_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this area cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'********************************************
'Add a new record
'
'SAJ 2009
'********************************************
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Area name"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdViewHistorical_Click()
'*******************************************************************
' New, to view historical naming of this area, opens form read only
'
' SAJ v9.1
'*******************************************************************
On Error GoTo err_cmdViewHistorical_Click

    DoCmd.OpenForm "Exca: Area Historical", acNormal, , "[CurrentAreaNumber] = " & Me![Area number], acFormReadOnly, acDialog
    

Exit Sub

err_cmdViewHistorical_Click:
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
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Area Sheet"
Exit Sub

err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Form_Current()
'*************************************************
' As mounds are  stored in many different tables,
' associated with the area don't allow them to altered
' after an area record has been entered.
'
' On screen msg tells user to contact DBA who can explain
' what a problem the change could have, unless it really is just
' an error done very recently.
'
' Also check if each area has any historical naming to determine
' whether view button should be enabled
'
' SAJ v9.1
'*************************************************
On Error GoTo err_Form_Current

If IsNull(Me![Area number]) Then
    'this is a new record, show mound combo
    Me![Field24].Visible = True
    Me![txtMound].Visible = False
Else
    'not a new record do not allow mound to be altered
    'so hide mound combo and show mound as a locked txt fld
    Me![Field24].Visible = False
    Me![txtMound].Visible = True
    Me![txtMound].Locked = True
End If

'now check if this area has any historical numbers and enable button if it does
Dim historical
historical = Null
'2009 error trap new record
If Not IsNull(Me![Area number]) Then
    historical = DLookup("[CurrentAreaNumber]", "[Exca: Area_Historical_Names]", "[CurrentAreaNumber] = " & Me![Area number])
End If

If Not IsNull(historical) Then
    Me![cmdViewHistorical].Enabled = True
Else
    Me![cmdViewHistorical].Enabled = False
End If
Exit Sub

err_Form_Current:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Open(Cancel As Integer)
'*************************************************************
' Check permissions of user to define how to display form
' v9.1 SAJ
'*************************************************************
On Error GoTo err_Form_Open
    If GetGeneralPermissions = "ADMIN" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
