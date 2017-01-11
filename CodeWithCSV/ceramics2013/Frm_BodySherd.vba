Option Compare Database
Option Explicit

Private Sub cboFilterUnit_AfterUpdate()
'new 2010 filter for unit
On Error GoTo err_filterunit

If Me![cboFilterUnit] <> "" Then
    Me.Filter = "[Unit] = " & Me![cboFilterUnit]
    Me.FilterOn = True
    Me![cmdRemoveFilter].Visible = True
End If

Exit Sub

err_filterunit:
    Call General_Error_Trap
    Exit Sub
End Sub

'Private Sub UpdateSurfaceTreatment(oldvalue, newvalue)
'Dim newval, count, sql
'
''check for value
'If newvalue = "" Or IsNull(newvalue) Then
'    MsgBox "Surface Treatment must be entered"
'    If oldvalue <> "" Then
'        newvalue = oldvalue
'    Else
'        SendKeys "{ESC}"
'        DoCmd.GoToControl "total"
'        DoCmd.GoToControl "txtSurfTreat"
'    End If
'Else
'    Dim checkexists
'    checkexists = DLookup("[Unit]", "Ceramics_Body_Sherd_SurfaceTreatment", "[Unit] = " & Me![txtUnit] & " AND [Ware Code] ='" & Me![WARE CODE] & "' AND [SurfaceTreatment] = '" & oldvalue & "'")
'    If Not IsNull(checkexists) Then
'        'does exist alter
'        '1.delete whats there
'        If spString <> "" Then
'            Dim mydb As DAO.Database
'            Dim myq1 As QueryDef
'
'            Set mydb = CurrentDb
'            Set myq1 = mydb.CreateQueryDef("")
'
'            myq1.Connect = spString
'
'                myq1.ReturnsRecords = False
 '               myq1.sql = "sp_Ceramics_Delete_BodySherd_SurfaceTreatment " & Me![txtUnit] & ", '" & Me![WARE CODE] & "', " & oldvalue
'                myq1.Execute
'
'            myq1.Close
'            Set myq1 = Nothing
'            mydb.Close
'            Set mydb = Nothing
'
'            '2. Add new value
'            If InStr(newvalue, ",") > 0 Then
'                '>1 treatment to add
'                newval = Split(newvalue, ",")
'                For count = 0 To UBound(newval)
'                    sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & newvalue & "'," & newval(count) & ")"
'                    DoCmd.RunSQL sql
'                Next
'
'            Else
'                'just one value
'                sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & newvalue & "'," & newvalue & ")"
'                DoCmd.RunSQL sql
'            End If
'        Else
'            'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
'            MsgBox "The existing surface treatment record has not been deleted, please contact the administrator.", vbCritical, "Error"
'            Exit Sub
'        End If
'    Else
'        'does not exist - add
'            If InStr(newvalue, ",") > 0 Then
'                '>1 treatment to add
'                newval = Split(newvalue, ",")
'                For count = 0 To UBound(newval)
'                    sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & newvalue & "'," & newval(count) & ")"
'                    DoCmd.RunSQL sql
'                Next
'
'            Else
'                'just one value
'                sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & newvalue & "'," & newvalue & ")"
'                DoCmd.RunSQL sql
'            End If
'    End If
'    Me!Frm_sub_bodysherd_surfacetreatment.Requery
'
'End If
'Exit Sub
'
'err_SurfTreat:
'    Call General_Error_Trap
'    Exit Sub
'
'
'End Sub

Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected unit from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'if a filter is on - turn off
         If Me.FilterOn = True Then
            Me.FilterOn = False
            Me![cmdRemoveFilter].Visible = False
            Me![cboFilterUnit] = ""
         End If
         
         'for existing number the field will be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![txtShowUnit].Enabled = False Then Me![txtShowUnit].Enabled = True
        DoCmd.GoToControl "txtShowUnit"
        DoCmd.FindRecord Me![cboFindUnit], , , , True
        Me![cboFindUnit] = ""
        DoCmd.GoToControl "cboFindUnit"
        Me![txtShowUnit].Enabled = False
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Close_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.Close acForm, Me.Name
    DoCmd.Restore
    
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click

    Dim thisunit
    thisunit = Me![txtUnit]
    
    DoCmd.GoToControl "Period" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    If Not IsNull(thisunit) Then Me![txtUnit] = thisunit
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![WareGroup].Enabled = True
    Me![WareGroup].Locked = False
    Me![WareGroup].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNewUnit_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToControl "Period" 'seems to get focus into tab control and then error as says it can't hide control that has focus
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![WareGroup].Enabled = True
    Me![WareGroup].Locked = False
    Me![WareGroup].BackColor = 16777215
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAlterSurfTreat_Click()
'5th aug 2010
'request that surface treatment change - effect primary key so have implemented cascade update in sql server
On Error GoTo err_alterSF

Dim Response, response2
Response = InputBox("Please enter the Surface Treatment:", "Surface Treatment change")
If Response <> "" Then
    response2 = MsgBox("Are you sure you want to change this records Surface Treatment from " & Me![txtSurfTreat] & " to " & Response & "?", vbExclamation + vbYesNo)
    If response2 = vbYes Then
        'hard work done by after update
        'txtSurfTreat_AfterUpdate 'moved to private sub proc
        Call UpdateSurfaceTreatment(Me![txtSurfTreat], Response, Me![txtUnit], Me![WARE CODE], Null)
        Me![txtSurfTreat] = Response
        Me.Refresh
    End If
Else
    'can't do as cancel returns a "" as well, best ignore it
    'MsgBox "A surface treatment value is required. Action cancelled", vbInformation, "No entry"
End If

Exit Sub
err_alterSF:
    If Err.Number = 3146 Then
        'duplicate key ie unit warecode surface treat already matching rec exists already
        MsgBox "This Unit-Ware code-Surface Treatment combination has already been entered. Change cancelled", vbCritical, "Record Exists"
        Me![txtSurfTreat] = Me![txtSurfTreat].oldvalue
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdAlterWareCode_Click()
'29th june 2009
'request that ware codes change - effect primary key so have implemented cascade update in sql server
On Error GoTo err_alter

Dim Response, response2
Response = InputBox("Please enter the altered Ware code:", "Ware code change")
If Response <> "" Then
    response2 = MsgBox("Are you sure you want to change this records ware code from " & Me![WARE CODE] & " to " & Response & "?", vbExclamation + vbYesNo)
    If response2 = vbYes Then
        Me![WARE CODE] = Response
        Me.Refresh
    End If

End If

Exit Sub
err_alter:
    If Err.Number = 3146 Then
        'duplicate key ie unit warecode surface treat already matching rec exists already
        MsgBox "This Unit-Ware code-Surface Treatment combination has already been entered. Change cancelled", vbCritical, "Record Exists"
        Me![WARE CODE] = Me![WARE CODE].oldvalue
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'allow deletion of entire record
On Error GoTo err_delete

Call DeleteBodySherdRecord(Me![txtUnit], Me![WareGroup], Me![txtSurfTreat])

Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdReNum_Click()
On Error GoTo err_ReNum
Dim val

    If Me![txtUnit] <> "" And Me![WareGroup] <> "" And Me![txtSurfTreat] <> "" Then
        val = ReNumberBodySherd(Me![txtUnit], Me![WareGroup], Me![txtSurfTreat])
        'new number if successful has been fed into find cbo so search to display.
        'if failed to update then cbofind will be blank so nothing happens
        cboFindUnit_AfterUpdate
        'MsgBox val
    Else
        MsgBox "Incomplete GID to process", vbInformation, "Action Cancelled"
    End If
Exit Sub

err_ReNum:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()

On Error GoTo err_current

    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![WareGroup] = "" Or IsNull(Me![WareGroup])) Then
    'don't include find number as defaults to x
    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) Then
        'new record allow GID entry
        Me![txtUnit].Enabled = True
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
        Me![WareGroup].Enabled = True
        Me![WareGroup].Locked = False
        Me![WareGroup].BackColor = 16777215
        Me![txtSurfTreat].Enabled = True
        Me![txtSurfTreat].Locked = False
        Me![txtSurfTreat].BackColor = 16777215
    Else
        'existing entry lock -- removed lock for WareGroup per request CE June 2014
        Me![txtUnit].Enabled = False
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = Me.Section(0).BackColor
        'Me![WareGroup].Enabled = False
        'Me![WareGroup].Locked = True
        'Me![WareGroup].BackColor = Me.Section(0).BackColor
        Me![txtSurfTreat].Enabled = False
        Me![txtSurfTreat].Locked = True
        Me![txtSurfTreat].BackColor = Me.Section(0).BackColor
        
    End If

'set focus to top
If Me![txtUnit].Enabled = True Then DoCmd.GoToControl "txtUnit"

Exit Sub

err_current:
    'when setting a new record the focus can be in the tab element and then the code above fails - trap this
    If Err.Number = 2165 Then 'you can't hide a control that has the focus.
        Resume Next
    Else
        Call General_Error_Trap
    End If
    Exit Sub


End Sub

Private Sub Form_Open(Cancel As Integer)
'new to disable admin features to non admin users
On Error GoTo er_open

If GetGeneralPermissions = "Admin" Then
    'Me!cmdAlterWareCode.Enabled = True
    'Me!cmdAlterSurfTreat.Enabled = True
    'renumber button taken over above
    Me![LabelChange1].Visible = True
    Me![LabelChange2].Visible = True
    Me!cmdDelete.Enabled = True
    Me!cmdReNum.Enabled = True
Else
    'Me!cmdAlterWareCode.Enabled = False
    'Me!cmdAlterSurfTreat.Enabled = False
    Me![LabelChange1].Visible = False
    Me![LabelChange2].Visible = False
    Me!cmdDelete.Enabled = False
    Me!cmdReNum.Enabled = False
End If

If Me.FilterOn = True Then
    Me![cmdRemoveFilter].Visible = True
End If

Exit Sub

er_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:

    If Err.Number = 2105 Then
        MsgBox "Entry for Waregroup in this Unit already exists"
    Else
        MsgBox Err.Description
        Resume Exit_go_next_Click
    End If
End Sub

Private Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    If Err.Number = 2105 Then
        MsgBox "Entry for Waregroup in this Unit already exists"
    Else
        Call General_Error_Trap
        Resume Exit_go_previous2_Click
    End If
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    Call General_Error_Trap
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    Call General_Error_Trap
    Resume Exit_go_last_Click
End Sub

Private Sub txtSurfTreat_AfterUpdate()
'new 2010 - instead of team having to write in the surface treatment into sep table the system will deal with this
Call UpdateSurfaceTreatment(Me![txtSurfTreat].oldvalue, Me![txtSurfTreat], Me![txtUnit], Me![WareGroup], Null)

If Me![txtSurfTreat] <> "" And Not IsNull(Me![txtSurfTreat]) Then
    DoCmd.RunCommand acCmdSaveRecord
End If

Me!Frm_sub_bodysherd_surfacetreatment.Form.Requery
Me!Frm_sub_bodysherd_surfacetreatment.Form.Refresh

'On Error GoTo err_SurfTreat
'Dim newval, count, sql
'
'check for value
''If Me![txtSurfTreat] = "" Or IsNull(Me![txtSurfTreat]) Then
'    MsgBox "Surface Treatment must be entered"
'    If Me![txtSurfTreat].oldvalue <> "" Then
'        Me![txtSurfTreat] = Me![txtSurfTreat].oldvalue
'    Else
'        SendKeys "{ESC}"
'        DoCmd.GoToControl "total"
'        DoCmd.GoToControl "txtSurfTreat"
'    End If
'Else
'    Dim checkexists
'    checkexists = DLookup("[Unit]", "Ceramics_Body_Sherd_SurfaceTreatment", "[Unit] = " & Me![txtUnit] & " AND [Ware Code] ='" & Me![WARE CODE] & "' AND [SurfaceTreatment] = '" & Me![txtSurfTreat] & "'")
'    If Not IsNull(checkexists) Then
'        'does exist alter
'        '1.delete whats there
'        If spString <> "" Then
'            Dim mydb As DAO.Database
'            Dim myq1 As QueryDef
'
'            Set mydb = CurrentDb
'            Set myq1 = mydb.CreateQueryDef("")
'
'            myq1.Connect = spString
'
'                myq1.ReturnsRecords = False
'                myq1.sql = "sp_Ceramics_Delete_BodySherd_SurfaceTreatment " & Me![txtUnit] & ", '" & Me![WARE CODE] & "', " & Me![txtSurfTreat]
'                myq1.Execute
'
'            myq1.Close
'            Set myq1 = Nothing
'            mydb.Close
'            Set mydb = Nothing
'
'            '2. Add new value
'            If InStr(Me![txtSurfTreat], ",") > 0 Then
'                '>1 treatment to add
'                newval = Split(Me![txtSurfTreat], ",")
'                For count = 0 To UBound(newval)
'                    sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & Me![txtSurfTreat] & "'," & newval(count) & ")"
'                    DoCmd.RunSQL sql
'                Next
'
'            Else
'                'just one value
'                sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & Me![txtSurfTreat] & "'," & Me![txtSurfTreat] & ")"
'                DoCmd.RunSQL sql
'            End If
'        Else
'            'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
'            MsgBox "The existing surface treatment record has not been deleted, please contact the administrator.", vbCritical, "Error"
'            Exit Sub
'        End If
'    Else
'        'does not exist - add
'            If InStr(Me![txtSurfTreat], ",") > 0 Then
'                '>1 treatment to add
'                newval = Split(Me![txtSurfTreat], ",")
'                For count = 0 To UBound(newval)
'                    sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & Me![txtSurfTreat] & "'," & newval(count) & ")"
'                    DoCmd.RunSQL sql
'                Next
'
'            Else
'                'just one value
'                sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Me![txtUnit] & ", '" & Me![WARE CODE] & "', '" & Me![txtSurfTreat] & "'," & Me![txtSurfTreat] & ")"
'                DoCmd.RunSQL sql
'            End If
'    End If
'    Me!Frm_sub_bodysherd_surfacetreatment.Requery
'
'End If
'Exit Sub
'
'err_SurfTreat:
'    Call General_Error_Trap
'    Exit Sub
    
End Sub

Private Sub txtSurfTreat_LostFocus()
Me!Frm_sub_bodysherd_surfacetreatment.Form.Requery
End Sub

Private Sub txtUnit_AfterUpdate()
Call CheckUnitDescript(Me![txtUnit])

End Sub

Private Sub WARE_CODE_NotInList(NewData As String, Response As Integer)
On Error GoTo err_warecode_NotInList

Dim retVal, sql
retVal = MsgBox("This Ware Code does not appear in the pre-defined list. Have you checked the list to make sure there is no match?", vbQuestion + vbYesNo, "New ware code")
If retVal = vbYes Then
    MsgBox "Ok this ware code will now be added to the list", vbInformation, "New Ware Code Allowed"
    'allow value,
     Response = acDataErrAdded
    
    Dim desc
    desc = InputBox("Please enter the description for this new code eg: DMS-fine", "Ware Code Description")
    If desc <> "" Then
        sql = "INSERT INTO [Ceramics_Code_Warecode_LOV] ([WareCode], [Description]) VALUES ('" & NewData & "', '" & desc & "');"
    Else
        sql = "INSERT INTO [Ceramics_Code_Warecode_LOV] ([WareCode]) VALUES ('" & NewData & "');"
    End If
    DoCmd.RunSQL sql
    
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub

err_warecode_NotInList:
    Call General_Error_Trap

    Exit Sub

End Sub
Private Sub cmdRemoveFilter_Click()
On Error GoTo Err_cmdRemoveFilter

    Me.Filter = ""
    Me.FilterOn = False
    Me![cboFilterUnit] = ""
    DoCmd.GoToControl "cboFindUnit"
    Me![cmdRemoveFilter].Visible = False

    Exit Sub

Err_cmdRemoveFilter:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub WareGroup_NotInList(NewData As String, Response As Integer)

End Sub
