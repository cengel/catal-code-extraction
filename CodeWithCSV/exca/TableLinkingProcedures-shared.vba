Option Compare Database

Option Explicit

'******************************************************************************
' This module was introduced in version 9 - its checks the table links and
' deals with logging the user in so Access connects to SQL Server correctly
'
' This module is also used in all other interfaces. Only shared improvements
' should be held here as it will be imported to other mdbs
'******************************************************************************

Function LogUserIn_OLD()
'******************************************************************************
'-->REDUNDANT _ REPLACED BY LOGUSERIN() KEPT FOR REFERENCE BY SAJ
' When a user first tries to look at one of the tables SQL Server can throw back
' the message "Login Failed for user (null). Reason: not associated with a trusted SQL server connection"
' followed by a login box where the 'Use trusted connection' box must be unchecked before the login details
' can be entered. This is confusing to users.
'
' To overcome this Access will obtain the users login details and by refreshing the link
' on just one table will enable all table links to work successfully. Its doesn't store the login details
' so it will always check on start up to facilitate different users.
'
' If the user doesn't login successfully the system will quit.
'
' If the user provides an invalid login via Access the SQL server login box will appear but without any confusing msg
'
' SAJ v. 9
'******************************************************************************
On Error GoTo err_LogUserIn_OLD

Dim username, pwd, retval

'prompt user for login name, two tries with different msgs before either looping to top or quiting
getuser:
    username = InputBox("Please enter your database LOGIN NAME:", "Login Name")
    If username = "" Then 'either the entered blank or pressed Cancel
        username = InputBox("The system cannot continue without your database login name. " & Chr(13) & Chr(13) & "Please enter your database LOGIN NAME below:", "Login Name")
        If username = "" Then 'again no entry
            retval = MsgBox("Sorry but the system cannot continue without a LOGIN NAME. Do you want to try again?", vbCritical + vbYesNo, "Login required")
            If retval = vbYes Then 'try again, loop back up
                GoTo getuser
            Else 'no, don't try again so quit system
                MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
                DoCmd.Quit acQuitSaveAll
            End If
        End If
    End If
    
'prompt user for password, two tries with different msgs before either looping to top or quiting
getpwd:
    pwd = InputBox("Please enter your database PASSWORD:", "Password")
    If pwd = "" Then 'either the entered blank or pressed Cancel
        pwd = InputBox("The system cannot continue without your database password. " & Chr(13) & Chr(13) & "Please enter your database PASSWORD below:", "Password")
        If pwd = "" Then 'again no entry
            retval = MsgBox("Sorry but the system cannot continue without a PASSWORD. Do you want to try again?", vbCritical + vbYesNo, "Password required")
            If retval = vbYes Then 'try again, loop back up
                GoTo getpwd
            Else 'no, don't try again so quit system
                MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
                DoCmd.Quit acQuitSaveAll
            End If
        End If
    End If


'user login name and password obtained, now find the first linked table in the tables collection (ignoring local tables)
' and refresh the link by adding on the UID and PWD
Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb

For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
             
    If tmptable.Connect <> "" Then
        'only deal with a table that is linked (not local)
        'tmptable.Connect = ";DATABASE=" & g_datapath 'if its were to a file
        'tmptable.Connect = tmptable.Connect & ";UID=" & username 'this will bring up the SQL server login box for pswd - better than previously as at least ready to recieve it
        tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
        tmptable.RefreshLink
        Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
    End If
            
Next I
    
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
    
LogUserIn_OLD = True
    
Exit Function

err_LogUserIn_OLD:
    If Err.Number = 3059 Then
        'operation cancelled by user - probably as login incorrect, sql server has asked for login as well and user pressed CANCEL
        retval = MsgBox("Sorry but the login you have given is incorrect or the database/internet connection is not available. You cannot connect to the database. Do you wish to try logging in again?", vbCritical + vbYesNo, "Login Failure")
        If retval = vbYes Then Resume
    ElseIf Err.Number = 3151 Then
        'ODBC--connection to 'xxxxx' failed.
        'odbc name cannot be found on this machine - send off to function to deal with
        AlterODBC
    Else
        MsgBox Err.Description & Chr(13) & Chr(13) & "The system will now quit", vbCritical, "Login Failure"
    End If
    LogUserIn_OLD = False
    DoCmd.Quit
End Function
Function AlterODBC()
Dim startstr, endstr, namestr
    'simply tell the user and quit
    If Err.Number = 3151 Then
        'odbc connection failed
        
        startstr = InStr(Err.Description, "'")
        endstr = InStr(startstr + 1, Err.Description, "'")
        namestr = Mid(Err.Description, startstr + 1, endstr - startstr)
        MsgBox "This system requires the ODBC connection: " & namestr & Chr(13) & Chr(13) & _
                        "The error returned is: " & Err.Description & Chr(13) & Chr(13) & "Instructions of how " & _
                        "to setup ths DSN can be found on the Web at http://catalsql.arch.cam.ac.uk/database/odbc.html" & _
                        "", vbCritical, "The system cannot start"
    End If
Exit Function

'OR deal with it!!
Dim username, pwd, tblName, rstemp

Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb


For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
             
    If tmptable.Connect <> "" Then
        'only deal with a table that is linked (not local)
        'tmptable.Connect = ";DATABASE=" & g_datapath 'if its were to a file
        'tmptable.Connect = tmptable.Connect & ";UID=" & username 'this will bring up the SQL server login box for pswd - better than previously as at least ready to recieve it
        tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
        tmptable.RefreshLink
        Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
    End If
            
Next I
'For i = 0 To mydb.TableDefs.count - 1 'loop the tables collection
'    Set tmptable = mydb.TableDefs(i)
'
'    If tmptable.Connect <> "" Then
'        tblName = tmptable.Name
'        On Error Resume Next
'        Set rstemp = mydb.OpenRecordset(tblName)
'        If Err <> 0 Then
'            If Err.Number = 3151 Then
'                'odbc connection failed
'                startstr = InStr(Err.Description, "'")
'                endstr = InStr(startstr + 1, Err.Description, "'")
'                namestr = Mid(Err.Description, startstr + 1, endstr - startstr)
'                MsgBox "This system requires the ODBC connection: " & namestr & Chr(13) & Chr(13) & _
'                        "The error returned is: " & Err.Description & Chr(13) & Chr(13) & "Instructions of how " & _
'                        "to setup ths DSN can be found on the Web at http://catalsql.arch.cam.ac.uk/database/odbc.html" & _
'                        "", vbCritical, "The system cannot start"
'                Exit For
'            End If
''            'wasp -> gcnox in upgr
''            'Result = reattachtables(gcnox, "NOXDATA.MDB")
''            ' Added this line as gcnox is no longer used due to the removal of the archive system.
''            Result = reattachtables("NOXDATA.MDB")
''            CheckTableAttachments = Result
'        Else
''            CheckTableAttachments = True
'            rstemp.close
'        End If
'End If
'Next i
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
End Function

Function LogUserIn(username As String, pwd As String)
'******************************************************************************
' When a user first tries to look at one of the tables SQL Server can throw back
' the message "Login Failed for user (null). Reason: not associated with a trusted SQL server connection"
' followed by a login box where the 'Use trusted connection' box must be unchecked before the login details
' can be entered. This is confusing to users.
'
' To overcome this Access will obtain the users login details and by refreshing the link
' on just one table will enable all table links to work successfully. Its doesn't store the login details
' so it will always check on start up to facilitate different users.
'
' If the user doesn't login successfully the system will quit.
'
' SAJ v. 9
'******************************************************************************
On Error GoTo err_LogUserIn

Dim retval

If username <> "" And pwd <> "" Then
    'user login name and password obtained
    Dim mydb As DAO.Database, I, errmsg, connStr
    Dim tmptable As TableDef
    Set mydb = CurrentDb
    
    Dim myq As QueryDef
    Set myq = mydb.CreateQueryDef("")
    connStr = ""
    
    'now find the first linked table in the tables collection (ignoring local tables)
    'to get the connection string to the sql server DB
    For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
         Set tmptable = mydb.TableDefs(I)
             
        If tmptable.Connect <> "" Then
            'only deal with a table that is linked (not local)
            'first check if the user login is valid -
            'this querydef check comes from MSDN KB
            'it will ensure the user details are valid and return a trappable error. This overcomes
            'the problem with the refresh link that would make sql server bring up its own login if
            'the refresh was done with invalid user details. By entering the login into the sql server
            'login box the connection between what the user had entered into the access form
            'and what sql server was using would be broken.
            
            If connStr = "" Then connStr = tmptable.Connect
            
            'new 2010
            Forms![Excavation_Login]![lblMsg] = "System is contacting the server"
            On Error Resume Next
                myq.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                myq.ReturnsRecords = False 'don't waste resources bringing back records
                myq.sql = "select [Unit Number] from [Exca: Unit Sheet] WHERE [Unit Number] = 1" 'this is a shared and core table so should always be avail, the record doesn't have to exist
                myq.Execute
            
            If Err <> 0 Then 'the login deails are incorrect
                GoTo err_LogUserIn
            Else
                'reset error trap
                On Error GoTo err_LogUserIn:
                ' the login is ok, so now try to refresh the link by adding on the UID and PWD
                'tmptable.Connect = ";DATABASE=" & g_datapath 'if its were to a file
                'tmptable.Connect = tmptable.Connect & ";UID=" & username 'this will bring up the SQL server login box for pswd - better than previously as at least ready to recieve it
                
                'new 2010
                Forms![Excavation_Login]![lblMsg] = "System is refreshing the link to the server"
                
                tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                tmptable.RefreshLink
                
                'new 2010
                Forms![Excavation_Login]![lblMsg] = "Refresh Complete"
            End If
            
            Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
        End If
            
    Next I
    
Else
    'somehow got here without username and pwd - catchall
    MsgBox "Both a username and password are required to operate the system correctly. Please quit and restart the application.", vbCritical, "Login problem encountered"
End If

Forms![Excavation_Login]![lblMsg] = "Defing permissions"
SetGeneralPermissions username, pwd, connStr 'requires more thought
'if user enters invalid login sql server will ask for proper one and if its ok connect on that
'and the link between the sql login and the access one is completely lost -****the querydef conn
'intro above should mean they always enter valid logon
LogUserIn = True

cleanup:
    myq.Close
    Set myq = Nothing
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
        
Exit Function

err_LogUserIn:
    If Err.Number = 3059 Or Err.Number = 3151 Then
        '3059 = operation cancelled by user - probably as login incorrect, sql server has asked for login as well and user pressed CANCEL
        '3151 = covers all the errors that could happen if user login incorrect, odbc not exist or internet conn not on
        errmsg = "Sorry but the system cannot log you into the database. There are three reasons this may have occurred:" & Chr(13) & Chr(13)
        errmsg = errmsg & "1. Your login details have been entered incorrectly" & Chr(13) & Chr(13)
        errmsg = errmsg & "2. There is no ODBC connection to the database setup on this computer." & Chr(13) & "    See http://www.catalhoyuk.com/database/odbc.html for details." & Chr(13) & Chr(13)
        errmsg = errmsg & "3. Your computer is not connected to the Internet at this time." & Chr(13) & Chr(13)
        errmsg = errmsg & "Do you wish to try logging in again?"
        retval = MsgBox(errmsg, vbCritical + vbYesNo, "Login Failure")
        If retval = vbYes Then
            GoTo cleanup 'used to be resume before querydef intro, now just cleanup and leave so user can try again
        Else
            'user says they dont want to try logging in again, double check in case they didn't understand so tell them we are quitting!
            retval = MsgBox("Are you really sure you want to quit and close the system?", vbCritical + vbYesNo, "Confirm System Closure")
            If retval = vbNo Then
                GoTo cleanup 'on 2nd thoughts the user doesn't want to quit so now just cleanup and leave so user can try again
            Else
                MsgBox "The system will now quit" & Chr(13) & Chr(13) & "The error reported was: " & Err.Description, vbCritical, "Login Failure"
            End If
        End If
    'ElseIf Err.Number = 3151 Then
    '    'ODBC--connection to 'xxxxx' failed.
    '    'odbc name cannot be found on this machine - send off to function to deal with
    '    AlterODBC
    Else
        MsgBox Err.Description & Chr(13) & Chr(13) & "The system will now quit", vbCritical, "Login Failure"
    End If
    LogUserIn = False
    DoCmd.Quit acQuitSaveAll
End Function
Sub WriteOutTableNames()
'*****************************************************************
' this is an admin bit of code and not related to the functioning
' of the system. It allows all the table names used by this DB to
' be printed to the immediate win
' SAJ
'*****************************************************************
Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb

For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If InStr(tmptable.Name, "MSys") = 0 Then
        Debug.Print tmptable.Name
        If tmptable.Connect <> "" Then
            Debug.Print "Linked"
        Else
            Debug.Print "Local"
        End If
    End If
Next I
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
End Sub

