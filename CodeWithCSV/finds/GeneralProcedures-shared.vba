Option Compare Database
Option Explicit
'******************************************************************************
' New module to hold general procedures that are shared across the diff db -easy to overwrite
'
' Intro by SAJ v.1 (on)
'******************************************************************************

Function StartUp()
'*****************************************************************************
' All actions necessary to start the system as smoothly as possible
'
' SAJ v.1 - actions from autoexac macro plus addition of login/attachment check
'*****************************************************************************
On Error GoTo err_startup

''DoCmd.RunCommand acCmdWindowHide 'hide the DB window from prying eyes

DoCmd.OpenForm "Login", acNormal, , , acFormEdit, acDialog

'you can hide the warning messages that Access popups up when
'you do sql tasks in the background - however the negative side to
'this is that you hide all these types of message which you may not
'want to do - the options you have are:
'   DoCmd.SetWarnings False 'turns off macro msgs
'   Application.SetOption "Confirm Record Changes", False
'   Application.SetOption "Confirm Document Deletions", False
'   Application.SetOption "Confirm Action Queries", False  'this will hide behind the scences sql actions
'you could of course turn this on an off around each statement - I'm undecided at present

'now the tables are all ok find out the current version
SetCurrentVersion

''DoCmd.Maximize
'OPEN YOUR MAIN MENU HERE
DoCmd.OpenForm "Finds", acNormal, , , acFormReadOnly 'open main menu

'refresh the main menu so the version number appears - REPLACE WITH YOUR MENU NAME
Forms![finds].Refresh

Exit Function

err_startup:
    Call General_Error_Trap
    'now should the system quit out here?
    'to be decided
End Function


Sub General_Error_Trap()
'******************************************************************************
' Display general error message
'
' SAJ v.1
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
' SAJ v1
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
    Dim mydb As Database, myrs As DAO.Recordset
    Dim sql
    Set mydb = CurrentDb()
    
    sql = "SELECT [Version_Num] FROM [Database_Interface_Version_History] WHERE [MDB_Name] = '" & DBName & "' AND not isnull([DATE_RELEASED]) ORDER BY [Version_Num] DESC;"
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    
     '' Dim myrs As New ADODB.Recordset
   '' myrs.Open sql, CurrentProject.Connection, adOpenKeySet, adLockOptimistic
    
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        centralver = myrs![Version_num]
        retVal = retVal & myrs![Version_num]
        
        'check local constant value held in module Globals-shared to see if this interface
        'matches current version of central copy, if not tell the user
        If centralver <> VersionNumberLocal Then
            Dim msg
            msg = "There is a new version of the Finds database file available. " & Chr(13) & Chr(13) & _
                    "Please close this copy now and run 'Update Databases.bat' on your desktop or " & _
                    "copy the file 'Finds Register Central.mdb' from G:\" & Year(Date) & " Central Server Databases " & _
                    " into the 'New Database Files folder' on your desktop." & Chr(13) & Chr(13) & "If you do not do this" & _
                    " you may experience problems using this database and you will not be able to utilise any new functionaility that has been added."
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
' read write this code assigns a global var to flag what it knows. If the
' naming convention for users is maintained then this should work ok
'
' Alternative way to do this might be to check the DB permissions tables for
' each username
'******************************************************************************
On Error GoTo err_SetGeneralPermissions
Dim tempVal, msg, usr

'If username = "" Then
'    tempVal = "RO"
'    msg = "Your login name is unknown to the system, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
'Else
'    usr = UCase(username)
'    If InStr(usr, "RO") <> 0 Or UCase(usr) = "CATALHOYUK" Then
'        tempVal = "RO"
'    ElseIf InStr(usr, "ADMIN") <> 0 Or usr = "RICH" Or usr = "MIA" Or usr = "SHAHINA" Or usr = "SARAH" Then
'        tempVal = "ADMIN"
'    ElseIf InStr(usr, "RW") <> 0 Then
'        tempVal = "RW"
'    Else
 '       tempVal = "RO"
'        msg = "The system is unsure of the rights of your login name so you have been assigned " & _
'                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
'                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
'                username & "' does not fall into any of the known types, please update the " & _
'                "SetGeneralPermissions code"
'    End If
'
'End If

'If msg <> "" Then
'    MsgBox msg, vbInformation, "Permissions setup"
'End If

' Alternative way to do this is to check the DB permissions tables for the user
' using a stored procedure to obtain whether the user has select permissions = RO
' if update = RW and if delete = Admin
Dim mydb As DAO.Database
Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
    'set the global spString to this conn string for use later in DeleteSampleRecord()
    spString = connStr & ";UID=" & username & ";PWD=" & pwd
    myq1.ReturnsRecords = True
    ''myq1.sql = "sp_table_list_with_privileges_for_a_user '%', 'dbo', null, '" & username & "'"
    myq1.sql = "sp_table_privilege_overview_for_user '%', 'dbo', null, '" & username & "'"

    Dim myrs As DAO.Recordset
    Set myrs = myq1.OpenRecordset
    'MsgBox myrs.Fields(0).Value
    'MsgBox username
    If myrs.Fields(0).Value = "" Then
        tempVal = "RO"
        msg = "Your permissions on the database cannot be defined, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
    Else
        usr = UCase(myrs.Fields(0).Value)
        If InStr(usr, "RO") <> 0 Then
            tempVal = "RO"
        ElseIf InStr(usr, "ADMIN") <> 0 Then
            '2009 currently finds user is an admin but want to restrict certain functions
            'so reset them to RW in terms of interface (they still have their sql delete permiss)
            If LCase(username) = "finds" Then
                tempVal = "RW"
            Else
                tempVal = "ADMIN"
            End If
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
    
    
    'added here to find out specific username which we need later for the
    'locations register to conditionally display fields and buttons
    
    CrateLetterFlag = "*"
    If username = "faunal" Then
        tempVal = "RW"
        CrateLetterFlag = "FB" 'this includes also Depot
    ElseIf username = "conservation" Or username = "conservationleader" Then
        tempVal = "RW"
        CrateLetterFlag = "CONS"
    ElseIf username = "ceramics" Or username = "ceramicsleader" Then
        tempVal = "RW"
        CrateLetterFlag = "P"
    ElseIf username = "clayobjects" Then
        tempVal = "RW"
        CrateLetterFlag = "CO" 'this includes also BE, FG, CB
    ElseIf username = "humanremains" Or username = "humanremainsleader" Then
        tempVal = "RW"
        CrateLetterFlag = "HB"
    ElseIf username = "chippedstone" Or username = "chippedstoneleader" Then
        tempVal = "RW"
        CrateLetterFlag = "OB"
    ElseIf username = "groundstone" Or username = "groundstoneleader" Then
        tempVal = "RW"
        CrateLetterFlag = "GS" 'this includes also NS, Depot
    ElseIf username = "phytoliths" Or username = "phytolithsleader" Then
        tempVal = "RW"
        CrateLetterFlag = "PH"
    ElseIf username = "shell" Or username = "shellleader" Then
        tempVal = "RW"
        CrateLetterFlag = "S"
    ElseIf username = "heavyresidue" Or username = "heavyresidueleader" Then
        tempVal = "RW"
        CrateLetterFlag = "BE"
    ElseIf username = "illustration" Then
        tempVal = "RW"
        CrateLetterFlag = "Illustrate"
    ElseIf username = "photography" Then
        tempVal = "RW"
        CrateLetterFlag = "PHOTO"
     ElseIf username = "archaeobots" Or username = "archaeobotsleader" Then
        tempVal = "RW"
        CrateLetterFlag = "char" 'this includes also "or" crates
     ElseIf username = "catalhoyuk" Then
        tempVal = "RO"
        CrateLetterFlag = ""
    End If
    'MsgBox CrateLetterFlag
    
    
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

