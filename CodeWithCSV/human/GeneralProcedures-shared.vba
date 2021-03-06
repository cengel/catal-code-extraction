Option Compare Database
Option Explicit
'******************************************************************************
' New module to hold general procedures that are shared across the diff db -easy to overwrite
'
' Intro by SAJ v.9 23/11/05 (on)
'******************************************************************************



Sub General_Error_Trap()
'******************************************************************************
' Display general error message
'
' SAJ v.9
'******************************************************************************

    MsgBox "The system has encountered an error. The message is as follows:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Error Code: " & Err.Number, vbOKOnly, "System Error"

End Sub
Function GetCurrentVersion()
'******************************************************************************
' Return current interface version number - if its empty its empty do not put
' a trap to go set it as this is directly called by the main menu that appears
' before the DB links have been checked and validated (therefore if you do this
' the sql server login error will occur)
'
' SAJ v9
'******************************************************************************
On Error GoTo err_GetCurrentVersion

    GetCurrentVersion = VersionNumber

Exit Function

err_GetCurrentVersion:
    Call General_Error_Trap
End Function

Function SetCurrentVersion()
'******************************************************************************
' Return current interface version number stored in DB
'
' SAJ v9
'******************************************************************************
On Error GoTo err_SetCurrentVersion

Dim retVal, centralver
retVal = "v"
If DBName <> "" Then
    Dim mydb As DAO.Database, myrs As DAO.Recordset
    Dim sql, theVersionNumberLocal
    Set mydb = CurrentDb()
    
    sql = "SELECT [Version_Num] FROM [Database_Interface_Version_History] WHERE [MDB_Name] = '" & DBName & "' AND not isnull([DATE_RELEASED]) ORDER BY [Version_Num] DESC;"
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    
     '' Dim myrs As New ADODB.Recordset
   '' myrs.Open sql, CurrentProject.Connection, adOpenKeySet, adLockOptimistic
    
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        
        'new 2010
        centralver = myrs![Version_num]
        retVal = retVal & myrs![Version_num]
        
        'check local constant value held in module Globals-shared to see if this interface
        'matches current version of central copy, if not tell the user
        '2009 was a problem on PC's with where comma showing as decimal so try to capture
        theVersionNumberLocal = VersionNumberLocal
        If InStr(centralver, ",") > 0 Then centralver = Replace(centralver, ",", ".")
        If InStr(theVersionNumberLocal, ",") > 0 Then theVersionNumberLocal = Replace(theVersionNumberLocal, ",", ".")
        
        If CDbl(centralver) <> CDbl(theVersionNumberLocal) Then
            Dim msg
            msg = "There is a new version of the Human Remains database file available. " & Chr(13) & Chr(13) & _
                    "Please close this copy now and run 'Update Databases.bat' on your desktop or " & _
                    "copy the file 'Human Remains Central Database.mdb' from G:\" & Year(Date) & " Central Server Databases " & _
                    " into the 'New Database Files folder' on your desktop." & Chr(13) & Chr(13) & "If you do not do this" & _
                    " you may experience problems using this database and you will not be able to utilise any new functionaility that has been added." & Chr(13) & Chr(13) & _
                    "DO NOT DO THIS IF YOU HAVE SAVED ANY NEW QUERIES INTO YOUR DESKTOP COPY OF THE DATABASE."
                    
            MsgBox msg, vbExclamation + vbOKOnly, "New version available"
        End If
    End If
    
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    
  
Else
    retVal = retVal & "X"
End If

VersionNumber = retVal
SetCurrentVersion = retVal

Exit Function
err_SetCurrentVersion:
    Call General_Error_Trap
End Function

Sub SetGeneralPermissions(username, pwd, connStr)
'******************************************************************************
' To make the link clearer between whether a user has only read-only rights or
' read write access this code assigns a global var to flag what it knows.
' SAJ v9.1
'******************************************************************************
On Error GoTo err_SetGeneralPermissions

Dim tempVal, msg, usr

'If the naming convention for users is maintained then the method below should work ok
' but it is not very safe so replaced below
'   If username = "" Then
'        tempVal = "RO"
'        msg = "Your login name is unknown to the system, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
'   Else
'        usr = UCase(username)
'        If InStr(usr, "RO") <> 0 Or UCase(usr) = "CATALHOYUK" Then
'            tempVal = "RO"
'        ElseIf InStr(usr, "ADMIN") <> 0 Or usr = "RICH" Or usr = "MIA" Or usr = "SHAHINA" Or usr = "SARAH" Then
'            tempVal = "ADMIN"
'        ElseIf InStr(usr, "RW") <> 0 Then
'            tempVal = "RW"
'        Else
'            tempVal = "RO"
'            msg = "The system is unsure of the rights of your login name so you have been assigned " & _
'                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
'                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
'                username & "' does not fall into any of the known types, please update the " & _
'                "SetGeneralPermissions code"
'        End If
'
'    End If


' Alternative way to do this is to check the DB permissions tables for the user
' using a stored procedure to obtain whether the user has select permissions = RO
' if update = RW and if delete = Admin
Dim mydb As DAO.Database
Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
    myq1.ReturnsRecords = True
    ''myq1.sql = "sp_table_list_with_privileges_for_a_user '%', 'dbo', null, '" & username & "'"
    myq1.sql = "sp_table_privilege_overview_for_user '%', 'dbo', null, '" & username & "'"

    Dim myrs As DAO.Recordset
    Set myrs = myq1.OpenRecordset
    ''MsgBox myrs.Fields(0).Value
    If myrs.Fields(0).Value = "" Then
        tempVal = "RO"
        msg = "Your permissions on the database cannot be defined, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
    Else
        usr = UCase(myrs.Fields(0).Value)
        If InStr(usr, "RO") <> 0 Then
            tempVal = "RO"
        ElseIf InStr(usr, "ADMIN") <> 0 Then
            tempVal = "ADMIN"
        ElseIf InStr(usr, "RW") <> 0 Then
            tempVal = "RW"
        Else
            tempVal = "RO"
            msg = "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types, please update the " & _
                "SetGeneralPermissions code"
        End If
    End If
    
    
myrs.Close
Set myrs = Nothing
myq1.Close
Set myq1 = Nothing
mydb.Close
Set mydb = Nothing

If msg <> "" Then
    MsgBox msg, vbInformation, "Permissions setup"
End If
''MsgBox tempVal
GeneralPermissions = tempVal
Exit Sub

err_SetGeneralPermissions:
    GeneralPermissions = "RO"
    msg = "An error has occurred in the procedure: SetGeneralPermissions " & Chr(13) & Chr(13)
    msg = msg & "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types"
                
    MsgBox msg, vbInformation, "Permissions setup"
    Exit Sub
End Sub
Function GetGeneralPermissions()
'******************************************************************************
' Return the current users status - if its empty call set function to reset
' but this will reset to RO
'
' SAJ v9
'******************************************************************************
On Error GoTo err_GetCurrentVersion

    If GeneralPermissions = "" Then
        SetGeneralPermissions "", "", ""
    End If
    
    GetGeneralPermissions = GeneralPermissions

Exit Function

err_GetCurrentVersion:
    Call General_Error_Trap
End Function
Sub ToggleFormReadOnly(frm As Form, readonly, Optional otherarg)
'*****************************************************************************
' To allow a form to toggle between readonly and edit. Changes look of form to
' reflect its status. Sometimes allowing certain fields to be edited only.
'
' Inputs: frm = form name; readonly = T/F; otherarg = additional info on additions, deletions etc
' SAJ v9.1
'*****************************************************************************
Dim ctl As Control, extra
Dim intI As Integer, intCanEdit As Integer
Const conTransparent = 0
Const conWhite = 16777215
    
On Error GoTo err_trap
    
    If Not IsMissing(otherarg) Then extra = otherarg
    
    'if allow edits is set the combo searches dont work - instead locking each control lower down
    If readonly = True Then
        With frm
            If extra <> "Additions" Then .AllowAdditions = False
            .AllowDeletions = False
'            .AllowEdits = False
        End With
    Else
        With frm
            'this next line is added to help with not allowing additions when a filter is set
            If extra = "NoAdditions" Then .AllowAdditions = False
            If extra <> "NoAdditions" Then .AllowAdditions = True
            If extra <> "NoDeletions" Then .AllowDeletions = True
'            .AllowEdits = True
        End With
    End If
    
    For Each ctl In frm.Controls
        With ctl
            'Debug.Print ctl.Name
            Select Case .ControlType
                Case acLabel
                    .SpecialEffect = acEffectNormal
                    .BorderStyle = conTransparent
                Case acTextBox
                    'there maybe some exceptions on some forms - keep main key editable for newrecords;
                    'ignore fields that are always locked like mound
                    ''If ((frm.Name = "Exca: Area Sheet") Or (frm.Name = "Exca: Building Sheet") Or (frm.Name = "Exca: Space Sheet") Or (frm.Name = "Exca: Feature Sheet" And .Name <> "Feature Number") Or (frm.Name = "Exca: Unit Sheet" And .Name <> "Unit Number")) And (.Name <> "Mound") Then
                     If .Name <> "Mound" And (frm.Name <> "Exca: Feature Sheet" Or (frm.Name = "Exca: Feature Sheet" And .Name <> "Feature Number")) And (frm.Name <> "Exca: Unit Sheet" Or (frm.Name = "Exca: Unit Sheet" And .Name <> "Unit Number")) Then
                        
                        If readonly = False Then
                            ''.SpecialEffect = acEffectSunken
                            If frm.DefaultView <> 2 Then 'single or continuous
                                .BackColor = conWhite
                            Else
                                frm.DatasheetBackColor = conWhite 'datasheet
                            End If
                            .Locked = False
                        Else
                            ''.SpecialEffect = acEffectNormal
                            '.BackColor = frm.Detail.BackColor
                            If frm.DefaultView <> 2 Then 'single or continuous
                                .BackColor = frm.Section(0).BackColor
                            Else
                                'frm.DatasheetBackColor = frm.Section(0).BackColor
                                'section color is -2147483633 this will set datasheet to BLACK!!???
                                'how ever this works - you would not believe how long this took
                                frm.DatasheetBackColor = RGB(236, 233, 216)   'datasheet
                            End If
                            .Locked = True
                        End If
                    End If
                Case acComboBox
                    'search combo's must not be affected
                    ''If .Name = "cboCountry" Or InStr(.Name, "Edit") <> 0 Then
                    If InStr(.Name, "Find") = 0 Then
                        If readonly = False Then
                            ''.SpecialEffect = acEffectSunken
                            .BackColor = conWhite
                            .Locked = False
                        Else
                            ''.SpecialEffect = acEffectNormal
                            '.BackColor = frm.Detail.BackColor
                            .BackColor = frm.Section(0).BackColor
                            .Locked = True
                        End If
                    End If
                Case acSubform, acCheckBox
                    If readonly = False Then
                        .Locked = False
                        .Enabled = True
                    Else
                        'put in some extra checks as some subforms are readonly anyway but
                        'can't be disabled as have buttons off to linking forms
                        'amendment - just need to set enabled to true and its ok (?)
                       ' If .Name <> "Exca: subform Features related to Building" And .Name <> "Exca: subform Spaces related to building" Then
                             .Locked = True
                             '.Enabled = False
                             .Enabled = True
                       ' End If
                    End If
                Case acOptionButton
                    If readonly = False Then
                        .Locked = False
                    Else
                         .Locked = True
                    End If
            End Select
        End With
    Next ctl
    
    Exit Sub
    
err_trap:
        MsgBox "An error occurred setting readonly on/off. Code will resume next line" & Chr(13) & "Error: " & Err.Description & " - " & Chr(13), vbInformation, "Error Identified"
        Resume Next
    
End Sub
Sub test(KeyAscii As Integer)
Dim strCharacter As String

    ' Convert ANSI value to character string.
    'strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    'KeyAscii = Asc(UCase(strCharacter))
    MsgBox KeyAscii
End Sub

Sub ListReferences()
Dim refCurr As Reference

  For Each refCurr In Application.References
    Debug.Print refCurr.Name & ": " & refCurr.FullPath
  Next

End Sub

Sub GetRolePermissions()
'**************************************************************************************
' This gets the latest set of permissions form each role and stores them in the table
' Database_Roles_Permissions. Uses the SQL DMO server object
'
' SAJ 25/11/2005
'**************************************************************************************
Dim oServer, oDatabase, oDatabaserole, oRolePermission, currentTable
On Error GoTo err_GetPermissionsForRoles

Set oServer = CreateObject("SQLDMO.SQLServer")
''oServer.LoginSecure = True 'windows auth
''oServer.Connect "catalsql.arch.cam.ac.uk"
oServer.LoginSecure = False
oServer.Connect "catalsql.arch.cam.ac.uk", "catalhoyuk", "catalhoyuk"

' Get and set a database object
Set oDatabase = oServer.Databases("catalhoyuk")

' Get and set a database role object
Set oDatabaserole = oDatabase.DatabaseRoles



' Here is the outer loop for the roles
For Each oDatabaserole In oDatabase.DatabaseRoles
    Debug.Print "Role Name: " & oDatabaserole.Name
    ' Exclude the system roles

        ' Get and set a permissions object - will fail for all db_* roles
        Set oRolePermission = oDatabaserole.ListObjectPermissions(63)
        
        'is only available for user defined roles and will fail on these roles
        'db_owner,db_accessadmin,db_securityadmin,db_ddladmin,db_backupoperator,db_datareader,db_datawriter,db_denydatareader,db_denydatawriter
        If Err.Number <> 20551 Then
            'so no error found for getting this roles permissions
        
            On Error GoTo err_GetPermissionsForRoles
        
            ' Inner loop for the permission, by object and group
            For Each oRolePermission In oDatabaserole.ListObjectPermissions(63)
                'this loop goes through each permission - so check if still referring to same table as last loop
                currentTable = oRolePermission.ObjectName
                 Debug.Print oDatabaserole.Name
                    Debug.Print oRolePermission.ObjectOwner + "." + oRolePermission.ObjectName
             Next
        End If
'    End If 'if dont want the public role this is the IF statement to re-instate

Next 'next database role

' Display if you want
'MsgBox strMsgText, , "Permission, Object, Group"

cleanup:
    On Error Resume Next
    
    Set oRolePermission = Nothing
    Set oDatabase = Nothing
    Set oServer = Nothing
    'MsgBox count
Exit Sub

err_GetPermissionsForRoles:
    'catch ant errors
    MsgBox Err.Description
    
    GoTo cleanup
    'If Err.Number = 20551 Then
    '    Exit For
    'End If

End Sub
