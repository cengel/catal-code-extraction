Option Compare Database
Option Explicit

Private Sub cboMaterialGroup_AfterUpdate()
On Error GoTo err_cboMat

'Me.Refresh
'Me![cboMaterialSubGroup].Requery

'new 2008 - if change main material group it can make the other groups nonsensical
'so check for oldvalue eg: changing a chipped stone record to botanical
If Not IsNull(Me![cboMaterialGroup].OldValue) Then
    MsgBox "Please update the subgroup and object type values to tally with this change", vbExclamation, "Material Group Change"
End If

If Me![cboMaterialGroup].Column(2) <> "" Then
    Me![txtDB] = Me![cboMaterialGroup].Column(2)
    
    If Me![cboMaterialGroup].Column(3) <> "" Then
        Me![txtForm] = Me![cboMaterialGroup].Column(3)
    End If
    
    If Me![cboMaterialGroup].Column(4) <> "" Then
        Me![txtID] = Me![cboMaterialGroup].Column(4)
    End If
    
    If Me![cboMaterialGroup].Column(5) <> "" Then
        Me![txtTable] = Me![cboMaterialGroup].Column(5)
    End If

End If

If Me![txtDB] = "" Then
    Me![cmdGoDB].Enabled = False
Else
    Me![cmdGoDB].Enabled = True
End If
Exit Sub

err_cboMat:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMaterialSubgroup_AfterUpdate()
On Error GoTo err_cboMaterialSubgroup

'Me![cboObjectType].Requery

If Me![cboMaterialSubgroup].Column(3) <> "" Then
    Me![txtDB] = Me![cboMaterialSubgroup].Column(3)
    
    If Me![cboMaterialSubgroup].Column(4) <> "" Then
        Me![txtForm] = Me![cboMaterialSubgroup].Column(4)
    End If
    
    If Me![cboMaterialSubgroup].Column(5) <> "" Then
        Me![txtID] = Me![cboMaterialSubgroup].Column(5)
    End If
    
    If Me![cboMaterialSubgroup].Column(6) <> "" Then
        Me![txtTable] = Me![cboMaterialSubgroup].Column(6)
    End If
End If

If Me![txtDB] = "" Then
    Me![cmdGoDB].Enabled = False
Else
    Me![cmdGoDB].Enabled = True
End If
Exit Sub

err_cboMaterialSubgroup:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboMaterialSubGroup_GotFocus()
'instead of setting the rowsource in properties set it here to ensure
'previous records fields stay visible
On Error GoTo err_cboMatSubGrp

Dim sql

sql = "SELECT Finds_Code_MaterialGroup_Subgroup.MaterialSubGroupID, Finds_Code_MaterialGroup_Subgroup.MaterialSubgroupText, "
sql = sql & "Finds_Code_MaterialGroup_Subgroup.MaterialGroupID, Finds_Code_MaterialGroup_Subgroup.RelatedDatabase, "
sql = sql & "Finds_Code_MaterialGroup_Subgroup.FormToShow, Finds_Code_MaterialGroup_Subgroup.IDField, Finds_Code_MaterialGroup_Subgroup.TableName FROM Finds_Code_MaterialGroup_Subgroup "
sql = sql & "WHERE (((Finds_Code_MaterialGroup_Subgroup.MaterialGroupID)=" & Forms![Finds: Basic Data]![frm_subform_materialstypes].Form![cboMaterialGroup] & "));"
Me![cboMaterialSubgroup].RowSource = sql

Exit Sub

err_cboMatSubGrp:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboObjectType_AfterUpdate()
On Error GoTo err_cboDescFocus
    'DoCmd.GoToControl Form![Finds: Basic Data]![Description].Name
    'DoCmd.GoToControl "Description"
Exit Sub

err_cboDescFocus:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboObjectType_GotFocus()
'instead of setting the rowsource in properties set it here to ensure
'previous records fields stay visible
On Error GoTo err_cboDescFocus

Dim sql
sql = "SELECT Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID, Finds_Code_ObjectTypes.ObjectTypeText, Finds_Code_MaterialSubGroup_ObjectType.MaterialSubGroupID"
sql = sql & " FROM Finds_Code_MaterialSubGroup_ObjectType INNER JOIN Finds_Code_ObjectTypes ON Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID = Finds_Code_ObjectTypes.ObjectTypeID"
sql = sql & " WHERE (((Finds_Code_MaterialSubGroup_ObjectType.MaterialSubGroupID)= " & Forms![Finds: Basic Data]![frm_subform_materialstypes].Form![cboMaterialSubgroup] & "));"

Me![cboObjectType].RowSource = sql

Exit Sub

err_cboDescFocus:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoDB_Click()
On Error GoTo err_handler

If Me![txtDB] <> "" Then
    
    'first check if the record is avail in the db
    Dim checkrec, fieldname, tablename, crit
    If Me![txtTable] <> "" Then
        fieldname = Me![txtID]
        tablename = Me![txtTable]
        crit = Me![txtID] & " = '" & Me![txtGID] & "'"
        checkrec = DLookup(fieldname, tablename, crit)
        If IsNull(checkrec) Then
            'can't find the record
            MsgBox "This record does not exist in the related database, please contact the relevant team leader for more information", vbInformation, "Link Failed"
            Exit Sub
        End If
    Else
        MsgBox "Sorry but the system does not know which database to check for this record, please contact the Finds Officer", vbInformation, "Link failed"
        Exit Sub
    End If
    
    
    'old code does not work with runtime as not take into account mdw
     'Dim appAccess As Access.Application
     Dim appAccess As New Access.Application
     '' Create new instance of Microsoft Access.
     'Set appAccess = CreateObject("Access.Application")
    
    Dim dbpath
    dbpath = Replace(CurrentDb.Name, "Finds Register Central.mdb", Me![txtDB])
     '' Open database in Microsoft Access window.
     MsgBox "This is a demo of how system linking could work. Please note at present the " & Me![txtDB] & " system will now appear BUT when it does DO NOT remove focus from the " & Me![txtDB] & " database until you close it. Closing it will return you to the finds database. If you do not close " & Me![txtDB] & " first the system will freeze."
     appAccess.OpenCurrentDatabase dbpath, False 'false = shared
     '' Open given form.
     ''as it will be: appAccess.DoCmd.OpenForm SchemeAdminForm, acNormal, , "[QuadNo] = " & Me![Lab Number], , acDialog
     'appAccess.DoCmd.OpenForm "Frm_GS_Main", acNormal, , "GID = '1267.x3'", , acDialog

'***** Current problem if you move focus off form you have to double click on
'*****  mdb file name to get it back
    '--'"G:\saj working\GroundStone-Official.mdb"
    'the schemeDB should now be open so can get it as an object reference
                        Set appAccess = getobject(dbpath)
                        'having it as an object reference allows use of the openform command
                        'to open the scheme admin form - it will hold focus till shut
                      ''  appAccess.DoCmd.OpenForm "Frm_GS_Main", acNormal, , "GID = '1267.x3'", , acDialog
                       
                       appAccess.DoCmd.Maximize
                        appAccess.DoCmd.OpenForm Me![txtForm], acNormal, , Me![txtID] & " = '" & Me![txtGID] & "'", , acDialog
                        'once shut the code will continue - so must close connection to DB
                        appAccess.DoCmd.Maximize
                        appAccess.CloseCurrentDatabase
                        'and quit the shell that it is runningin
                        appAccess.Quit
                        Set appAccess = Nothing
    '--
'     appAccess.CloseCurrentDatabase
'     Set appAccess = Nothing
                
     '           'new code to take into account mdw
                Dim wrkgrpPath, accessPath, cmd As String, shellobj
                Dim ProcessHandle As Long
                Dim ExitCode As Long

                'first get user to enter password
     '           g_temppwd = ""
     '           DoCmd.OpenForm "FPopSchemeLogin", acNormal, , , , acDialog
                
     '           If g_temppwd <> "" Then
                
     '               'get path to msaccess.exe and current workgroup
   '                 wrkgrpPath = SysCmd(acSysCmdGetWorkgroupFile)
  '                  accessPath = SysCmd(acSysCmdAccessDir)
     '               If wrkgrpPath <> "" And accessPath <> "" Then
      '                  'can only proceed if workgroup and msaccess.exe are know
 '                      Dim appAccess As Access.Application
                       'build up string to open DB with workgroup
 '                      cmd = Chr(34) & accessPath & "MSACCESS.EXE" & Chr(34) & " " & Chr(34) & "G:\saj working\GroundStone-Official.mdb" & Chr(34)
 '                       cmd = cmd & " /nostartup '/user catalhoyuk" '& CurrentUser()
                     '   cmd = cmd & " /user catalhoyuk" '& CurrentUser()
                      '  cmd = cmd & " /pwd catalhoyuk " '& g_temppwd
 '                       cmd = cmd & " /wrkgrp " & Chr(34) & wrkgrpPath & Chr(34)
                        'this maybe slow so show processing with hourgladd
 '                       DoCmd.Hourglass True
                    
                        'for security purposes now blank the pswd global
       '                 g_temppwd = ""
       '
                        'to open access with the workgroup must use the shell command
                        'window style is minimused with out focus - DO NOT use hide as will
                        'not be able to see the form and will appear that it has hung
  '                      shellobj = Shell(pathname:=cmd, windowstyle:=6)
                        
                        'the schemeDB should now be open so can get it as an object reference
  '                      Set appAccess = GetObject("G:\saj working\GroundStone-Official.mdb")
                        'having it as an object reference allows use of the openform command
                        'to open the scheme admin form - it will hold focus till shut
  '                      appAccess.DoCmd.OpenForm "Frm_GS_Main", acNormal, , "GID = '1267.x3'", , acDialog
                        'once shut the code will continue - so must close connection to DB
  '                      appAccess.CloseCurrentDatabase
                        'and quit the shell that it is runningin
  '                      appAccess.Quit
  '                      Set appAccess = Nothing
                       'end of processing so reset cursor
  '                      DoCmd.Hourglass False
       '             Else
       '                 'tell user if workgroup cannot be found
       '                 If wrkgrpPath = "" Then MsgBox "Path to Workgroup file cannot be found. Please ask an Administrator to update this value in the WEBPATHS screen", vbInformation, "Unable to proceed action"
       '                 'tell the user if access.exe cannot be found
       '                 If accessPath = "" Then MsgBox "The path to MSAccess cannot be located. Please contact the Administrator of this problem", vbInformation, "Unable to proceed action"
       '             End If
       '         Else
       '             'tell user they cannot proceed
       '             MsgBox "No password supplied. The scheme administration form cannot be opened", vbInformation, "Unable to proceed"
       '         End If
       '     End If
End If
Exit Sub

err_handler:

    If Err.Number = -2147467259 Or Err.Number = 432 Then
       ' the user has entered the incorrect password - must shut the underlying shell
'       If KillProcess(shellobj, 0) Then
'         'MsgBox "App was terminated"
'       End If
       
       'Set shellobj = Null
        If Err.Number = -2147467259 Then MsgBox "The system cannot open the scheme database - you must enter a valid password", vbCritical, "Scheme cannot be opened"
        If Err.Number = 432 Then MsgBox "The system cannot open the scheme database. The path to the database entered into the Scheme Administration screen cannot be found", vbCritical, "Scheme cannot be opened"
        DoCmd.Hourglass False
    Else
        DoCmd.Hourglass False
    
        MsgBox "An error has occurred in General Procedures - OpenASchemeAtGivenForm(). The error is:" & Chr(13) & Chr(13) & Err.Number & " -- " & Err.Description
    End If
    Exit Sub





End Sub

Private Sub Form_Current()
On Error GoTo err_current

If Me.RecordsetClone.RecordCount > 1 Then
'    Me.DefaultView = 1
'    'DoCmd.RunCommand acCmdDatasheetView
'    MsgBox ">1"
    Forms![Finds: Basic Data]![frm_subform_materialstypes].Height = "1400"
Else
'    Me.DefaultView = 0
'    'DoCmd.RunCommand acCmdFormView
'    MsgBox "1"
    'Me.Height = "2000"
    Forms![Finds: Basic Data]![frm_subform_materialstypes].Height = "1000"
End If

Me![txtDB] = ""
Me![txtForm] = ""
Me![txtID] = ""

'Me![cboMaterialGroup].Requery
'Me![cboMaterialSubGroup].Requery
'Me![cboObjectType].Requery
Me![cboMaterialSubgroup].RowSource = "SELECT Finds_Code_MaterialGroup_Subgroup.MaterialSubGroupID, Finds_Code_MaterialGroup_Subgroup.MaterialSubgroupText, Finds_Code_MaterialGroup_Subgroup.MaterialGroupID, Finds_Code_MaterialGroup_Subgroup.RelatedDatabase,  Finds_Code_MaterialGroup_Subgroup.FormToShow, Finds_Code_MaterialGroup_Subgroup.IDField, Finds_Code_MaterialGroup_Subgroup.TableName FROM Finds_Code_MaterialGroup_Subgroup"
Me![cboObjectType].RowSource = "SELECT Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID, Finds_Code_ObjectTypes.ObjectTypeText, Finds_Code_MaterialSubGroup_ObjectType.MaterialSubGroupID  FROM Finds_Code_MaterialSubGroup_ObjectType INNER JOIN Finds_Code_ObjectTypes ON Finds_Code_MaterialSubGroup_ObjectType.ObjectTypeID = Finds_Code_ObjectTypes.ObjectTypeID"

'if subgroup has the DB take that
If Me![cboMaterialSubgroup].Column(3) <> "" Then
    Me![txtDB] = Me![cboMaterialSubgroup].Column(3)
    
    If Me![cboMaterialSubgroup].Column(4) <> "" Then
        Me![txtForm] = Me![cboMaterialSubgroup].Column(4)
    End If
    
    If Me![cboMaterialSubgroup].Column(5) <> "" Then
        Me![txtID] = Me![cboMaterialSubgroup].Column(5)
    End If

    If Me![cboMaterialSubgroup].Column(6) <> "" Then
        Me![txtTable] = Me![cboMaterialSubgroup].Column(6)
    End If
Else
'if sugroup not have db, check if group has and test that
    If Me![cboMaterialGroup].Column(2) <> "" Then
        Me![txtDB] = Me![cboMaterialGroup].Column(2)
    
        If Me![cboMaterialGroup].Column(3) <> "" Then
            Me![txtForm] = Me![cboMaterialGroup].Column(3)
        End If
    
        If Me![cboMaterialGroup].Column(4) <> "" Then
            Me![txtID] = Me![cboMaterialGroup].Column(4)
        End If

        If Me![cboMaterialGroup].Column(5) <> "" Then
            Me![txtTable] = Me![cboMaterialGroup].Column(5)
        End If
    End If
End If

If Me![txtDB] = "" Then
    Me![cmdGoDB].Enabled = False
Else
    Me![cmdGoDB].Enabled = True
End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub

End Sub
