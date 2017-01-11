Option Compare Database
Option Explicit



Private Sub cboFind_AfterUpdate()
'find skeleton record - SAJ
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
        Me.Filter = "[UnitNumber] = " & Me![cboFind]
        Me.FilterOn = True
    End If
Exit Sub

err_cboFind:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "This unit number has not been entered yet", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
    DoCmd.GoToControl "cmdAddNewUnit"
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'open skeleton description form and add a new record
On Error GoTo err_cmdAddNew
Dim sql
    'to add new individ must get last individ num allocated and increment
    If Me![txtUnit] <> "" Then
        'new may 2009 - table relationships set up so now must save this record to allow insert into basicskeletondata
        DoCmd.RunCommand acCmdSaveRecord
        'check if there are any individuals added for the unit yet
        Dim checknum
        checknum = DLookup("[UnitNumber]", "[HR_BasicSkeletonData]", "[UnitNumber] = " & Me![txtUnit])
        If IsNull(checknum) Then
            'it is not there, add number 1
            sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", 1);"
            DoCmd.RunSQL sql
            Me.Refresh
            MsgBox "Individual number 1 added for Unit " & Me![txtUnit], vbInformation, "Record added"
        Else
            'there is already a skeleton so now must get last individ number and +1
            Dim mydb As Database, myrs As DAO.Recordset, lastnumber, nextnumber
            Set mydb = CurrentDb()
            sql = "SELECT HR_BasicSkeletonData.UnitNumber, HR_BasicSkeletonData.[Individual number] FROM HR_BasicSkeletonData WHERE HR_BasicSkeletonData.UnitNumber = " & Me![txtUnit] & " ORDER BY HR_BasicSkeletonData.UnitNumber, HR_BasicSkeletonData.[Individual number];"
            Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not myrs.BOF And Not myrs.EOF Then
                myrs.MoveLast
                lastnumber = myrs![Individual number]
                nextnumber = lastnumber + 1
                sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", " & nextnumber & ");"
                DoCmd.RunSQL sql
                Me.Refresh
                MsgBox "Individual number " & nextnumber & " added for Unit " & Me![txtUnit], vbInformation, "Record added"
            Else
                'shouldn't get here but just in case, no record exists for this unit so create
                 sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", 1);"
                DoCmd.RunSQL sql
                Me.Refresh
            End If
        End If
        
        'DoCmd.OpenForm "FRM_SkeletonDescription", acNormal, , , acFormAdd
        'Forms![FRM_SkeletonDEscription]![txtUnit] = Me![txtUnit]
        '' DoCmd.GoToControl "txtUnit"
        'DoCmd.Close acForm, Me.Name
        
    Else
        MsgBox "You must enter a unit number first", vbInformation, "Unit Number Missing"
    End If

    
Exit Sub

err_cmdAddNew:
    MsgBox Err.Description
Exit Sub
End Sub

Private Sub cmdAddNewUnit_Click()
'add a new unit here
On Error GoTo err_cmdAddNewUnit

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNewUnit:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdAll_Click()
'take off any filter - saj
On Error GoTo err_all

    Me.FilterOn = False
    Me.Filter = ""
Exit Sub

err_all:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub CmdOpenAgeSexFrm_Click()
'check if record exists in age sext table first, if not create it
On Error GoTo Err_CmdOpenAgeSexFrm_Click

    Call DoRecordCheck("HR_ageing and sexing", Me![txtUnit], Me![txtIndivid], "Unit Number")

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , "[Unit Number] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub

Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
    
End Sub
Private Sub CmdOpenMainMenuFrm_Click()
On Error GoTo Err_CmdOpenMainMenuFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_MAINMENU"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    'saj
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenMainMenuFrm_Click:
    Exit Sub

Err_CmdOpenMainMenuFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMainMenuFrm_Click
    
End Sub

Private Sub Combo28_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub cmdReNumber_Click()
On Error GoTo err_renum
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN") Then
    Dim newnum, check
    newnum = InputBox("Please enter the unit number you would like to change ALL the references of unit " & Me![txtUnit] & " to:", "Re-number unit records")
    If newnum <> "" Then
        'must check number not already exit
        check = DLookup("[UnitNumber]", "[HR_UnitDescription]", "[Unitnumber] = " & newnum)
        If Not IsNull(check) Then
            MsgBox "Sorry but the unit number " & newnum & " already exists. You must delete/renumber that first before you can alter " & Me![txtUnit], vbExclamation, "Unit already exists"
            Exit Sub
        Else
            Me![txtUnit] = newnum
            MsgBox "Re-numbering has been performed", vbInformation, "Complete"
        End If
    End If
Else
    MsgBox "You do not have permissions to renumber units, please talk to your team leader£"
End If
Exit Sub

err_renum:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'new 2009 - lock unit number field where it exists
On Error GoTo err_current
    
    If Me![txtUnit] <> "" Or Not IsNull(Me![txtUnit]) Then
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = 8454143
    Else
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
    End If

    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Me![cmdReNumber].Enabled = True
    Else
        Me![cmdReNumber].Enabled = False
    End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Delete(Cancel As Integer)
'new may 2009, now cascade delete setup this is a dangerous operation - warn user
On Error GoTo err_delete

Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN") Then
    Dim response
    response = MsgBox("Deleting this unit will mean permanent deletion of any skeleton records associated with it in this database." & Chr(13) & Chr(13) & "Do you really want to delete unit " & Me![txtUnit] & " and its skeleton records?", vbCritical + vbYesNo, "Critical Delete")
    If response = vbNo Then
        Cancel = True
    Else
        Cancel = False
    End If
Else
    MsgBox "You do not have permission to delete this record, please contact your team leader"
    Cancel = True
End If

Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'saj season 2007 - Basak suggested its best to open this form ready for a new record
On Error GoTo err_open

   ' DoCmd.RunCommand acCmdRecordsGoToNew
   ' as this form can be opened to see the descript of a unit from else where can't set the new rec
   ' here, instead moved it to the menu button that opens tihs form as its those instances where
   ' they need a new record

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub txtUnit_AfterUpdate()
'if this is being updated it should be a new unit therefore we need to create a new
'record in skeletondescription as obviously this new record will have at least one skeleton
'season 2007, saj
On Error GoTo err_txtUnit
     
'MsgBox Me![txtUnit].OldValue

'If IsNull(Me![txtUnit].OldValue) And (Me![txtUnit] <> "" And Not IsNull(Me![txtUnit])) Then
'    'check there is no entry first for this unit in skeleton description already
'    Dim checknum
'    checknum = DLookup("[UnitNumber]", "[HR_BasicSkeletonData]", "[UnitNumber] = " & Me![txtUnit])
'    If IsNull(checknum) Then
'        'it is not there, add it
'        Dim sql
'        sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", 1);"
'        DoCmd.RunSQL sql
'        Me.Refresh
'    End If
'End If
Exit Sub

err_txtUnit:
    Call General_Error_Trap
    Exit Sub
End Sub
