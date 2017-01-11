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
'SetCurrentVersion
'DoCmd.Maximize
'OPEN YOUR MAIN MENU HERE
DoCmd.OpenForm "Frm_Menu", acNormal, , , acFormPropertySettings 'open main menu
'refresh the main menu so the version number appears - REPLACE WITH YOUR MENU NAME
'Forms![Frm_Main].Refresh

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
    Dim sql, theVersionNumberLocal
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
        '2009 was a problem on PC's with where comma showing as decimal so try to capture
        theVersionNumberLocal = VersionNumberLocal
        If InStr(centralver, ",") > 0 Then centralver = Replace(centralver, ",", ".")
        If InStr(theVersionNumberLocal, ",") > 0 Then theVersionNumberLocal = Replace(theVersionNumberLocal, ",", ".")
        
        'If centralver <> VersionNumberLocal Then
        If CDbl(centralver) <> CDbl(theVersionNumberLocal) Then
            Dim msg
            msg = "There is a new version of the Ceramics database file available. " & Chr(13) & Chr(13) & _
                    "Please close this copy now and run 'Update Databases.bat' on your desktop or " & _
                    "copy the file 'Ceramics_2009.mdb' from G:\" & Year(Date) & " Central Server Databases " & _
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
    spString = connStr & ";UID=" & username & ";PWD=" & pwd
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

Sub DeleteARecord(FromTable, criteria, mydb)
'This is an admin function to delete records, as used within a transaction it must not
'be error trapped. The db workspace is passed in for the transaction
'Inputs -   Fromtable - delete from what table
'           criteria - to select record to delete
'           mydb - currentdb connection sent thro for transaction


Dim sql, myq As QueryDef
Set myq = mydb.CreateQueryDef("")
           
        
        sql = "DELETE FROM [" & FromTable & "] WHERE " & criteria & ";"
        
                
        myq.sql = sql
        myq.Execute dbSeeChanges 'this was added as the elements tables was throwing error
        'that it needed dbSeeChanges as it had identity column
        
                
myq.Close
Set myq = Nothing


End Sub

Function DeleteDiagnosticRecord(Unit, letter, FindNumber)

Dim retVal, msg, msg1, strcriteria

    msg = "Are you quite sure that you want to permanently delete sherd: " & Unit & "." & letter & FindNumber & "?"
    retVal = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retVal = vbYes Then
        strcriteria = "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
       
        Call DeleteARecord("Ceramics_Diagnostic_DrawingNumbers", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_InclusionsDetermined", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_PhotoNumbers", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_ResidueSamples", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_SecondaryUse", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_Slips", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_SurfaceTreatments", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_Technology", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic_Elements", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Diagnostic", strcriteria, mydb)
    
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Deletion has been successful"
            Forms!Frm_Diagnostic.Requery
            Forms!Frm_Diagnostic![cboFindUnit].Requery
        Else
            wrkdefault.Rollback
            'MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description

            msg = "A problem has occured and the deletion has been cancelled. " & Chr(13) & Chr(13) & Err.Description
            MsgBox msg
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    Else
        MsgBox "Deletion cancelled", vbInformation, "Action Cancelled"
    
    End If
    
Exit_cmdDelete_Click:
    Exit Function

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Function
End Function

Function ReNumberDiagnostic(Unit, letter, FindNumber)
'new season 2009 - saj
On Error GoTo err_cmdReNumber

    Dim retUnit, retLetter, retNumber, findGID, sql, Response, msg
    retUnit = InputBox("Please enter the new Unit number for sherd " & Unit & "." & letter & FindNumber & ":", "Enter new unit number")
    If retUnit = "" Then
        MsgBox "No unit number entered, action cancelled"
    Else
        'check valid entry
        If Not IsNumeric(retUnit) Then
            MsgBox "Invalid Unit number, please try again", vbExclamation, "Action Cancelled"
            Exit Function
        End If
        
        retLetter = InputBox("Please enter the new letter code for Unit " & Unit & "." & letter & FindNumber & " (S or X):", "Enter new letter code")
        
        If retLetter = "" Then
            MsgBox "No letter code entered, action cancelled"
        Else
            retLetter = UCase(retLetter) 'always capitals on this form
            retNumber = InputBox("Please enter the new number for Unit " & Unit & "." & letter & FindNumber & " (was number " & FindNumber & "):", "Enter new number")
            If retNumber = "" Then
                MsgBox "No number entered, action cancelled"
            Else
                
                'now check if this sherd already exist already exists
                findGID = DLookup("[GID]", "[Q_Diagnostic_GIDs]", "[GID] = '" & retUnit & "." & retLetter & retNumber & "'")
                If Not IsNull(findGID) Then
                    MsgBox "Sorry but this Sherd GID (" & retUnit & "." & retLetter & retNumber & ") already exists. You must delete it first before you can alter " & Unit & "." & letter & FindNumber, vbExclamation, "Sherd GID already exists"
                    Exit Function
                Else
            
                    'new GID, does not exist so allow alteration
                    msg = "Are you quite sure that you want to renumber sherd " & Unit & "." & letter & FindNumber & " to " & retUnit & "." & retLetter & retNumber & "?"
                    Response = MsgBox(msg, vbExclamation + vbYesNoCancel, "Confirm Sherd Re-Number")
                    If Response = vbYes Then
                        On Error Resume Next
                        Dim mydb As DAO.Database, wrkdefault As Workspace, wrkdefault1 As Workspace
                        Set wrkdefault = DBEngine.Workspaces(0)
                        Set mydb = CurrentDb
        
                        '' Start of transaction.
                        wrkdefault.BeginTrans
        
                        'change unit number
                        Call RenumARecord("Ceramics_Diagnostic", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        'these should all be done by cascade - but do here to
                        Call RenumARecord("Ceramics_Diagnostic_Elements", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_DrawingNumbers", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_InclusionsDetermined", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_PhotoNumbers", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_ResidueSamples", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_SecondaryUse", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_Slips", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_SurfaceTreatments", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        Call RenumARecord("Ceramics_Diagnostic_Technology", "[Unit] = " & retUnit & ", [LetterCode] = '" & retLetter & "', [FindNumber] = " & retNumber, "[Unit] = " & Unit & " AND [LetterCode] = '" & letter & "' AND [FindNumber] = " & FindNumber, mydb)
                        
                        If Err.Number = 0 Then
                            wrkdefault.CommitTrans
                            Forms!Frm_Diagnostic.Requery
                            Forms!Frm_Diagnostic![cboFindUnit].Requery
                            Forms!Frm_Diagnostic![cboFindUnit] = retUnit & "." & retLetter & retNumber
                            MsgBox "Renumbering has been successful. Renumbered record will be displayed."
                        Else
                            wrkdefault.Rollback
                            MsgBox "A problem has occured and the re-numbering has been cancelled. The error message is: " & Err.Description
                        End If
                        mydb.Close
                        Set mydb = Nothing
                        wrkdefault.Close
                        Set wrkdefault = Nothing
                        
                    Else
                        MsgBox "Re-numbering cancelled", vbInformation, "Action Cancelled"
    
                    End If
                End If
            End If
        End If
    End If

Exit Function

err_cmdReNumber:
    Call General_Error_Trap
    Exit Function
End Function

Sub RenumARecord(FromTable, setStatement, whereStatement, mydb)
'This is an admin function to renum records, used within a transaction if must not
'be error trapped. The db workspace is passed in for the transaction
'Inputs -   Fromtable - renum in what table
'           setStatement -
'           whereStatement -
'           mydb - currentdb connection sent thro for transaction


Dim sql, myq As QueryDef
Set myq = mydb.CreateQueryDef("")
           
        
        sql = "UPDATE [" & FromTable & "] SET " & setStatement & " WHERE " & whereStatement & ";"
        
                
        myq.sql = sql
        myq.Execute
                
myq.Close
Set myq = Nothing


End Sub


Function DeleteBodySherdRecord(Unit, WareCode, surftreat)

Dim retVal, msg, msg1, strcriteria, strcriteria2

    msg = "Are you quite sure that you want to permanently delete sherd group: " & Unit & "." & WareCode & "-" & surftreat & "?"
    retVal = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retVal = vbYes Then
        strcriteria = "[Unit] = " & Unit & " AND [Ware Code] = '" & WareCode & "' AND [SurfaceTreatment] = '" & surftreat & "'"
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
       
        Call DeleteARecord("Ceramics_Body_Sherd_Technology", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Body_Sherd_SurfaceTreatment", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Body_Sherd_SecondaryUse", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Body_Sherd_SampleNumbers", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Body_Sherd_Residues", strcriteria, mydb)
        
        'a record can have more than one inclusion group
        Dim myrs2 As DAO.Recordset
        strcriteria2 = "SELECT InclusionGroupID FROM Ceramics_Body_Sherd_Inclusion_Group WHERE " & strcriteria & ";"
        Set myrs2 = mydb.OpenRecordset(strcriteria2, dbOpenDynaset, dbSeeChanges)
        If Not (myrs2.BOF And myrs2.EOF) Then
            myrs2.MoveFirst
            Do Until myrs2.EOF
                strcriteria2 = "[Inclusion_Group_ID] = " & myrs2![InclusionGroupID]
                Call DeleteARecord("Ceramics_Body_Sherd_InclusionsDetermined", strcriteria2, mydb)
                myrs2.MoveNext
            Loop
        End If
        myrs2.Close
        Set myrs2 = Nothing
        Call DeleteARecord("Ceramics_Body_Sherd_Inclusion_Group", strcriteria, mydb)
        Call DeleteARecord("Ceramics_Body_Sherds", strcriteria, mydb)
    
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Deletion has been successful"
            Forms!Frm_BodySherd.Requery
            Forms!Frm_BodySherd![cboFindUnit].Requery
        Else
            wrkdefault.Rollback
            'MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description

            msg = "A problem has occured and the deletion has been cancelled. " & Chr(13) & Chr(13) & Err.Description
            MsgBox msg
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    Else
        MsgBox "Deletion cancelled", vbInformation, "Action Cancelled"
    
    End If
    
Exit_cmdDelete_Click:
    Exit Function

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Function
End Function

Function ReNumberBodySherd(Unit, WareCode, surftreat)
'new season 2009 - saj
On Error GoTo err_cmdReNumber

    Dim retUnit, retWare, retSurf, findGID, sql, Response, msg
    retUnit = InputBox("Please enter the new Unit number for sherd group: " & Unit & "." & WareCode & surftreat & ":", "Enter new unit number", Unit)
    If retUnit = "" Then
        MsgBox "No unit number entered, action cancelled"
    Else
        'check valid entry
        If Not IsNumeric(retUnit) Then
            MsgBox "Invalid Unit number, please try again", vbExclamation, "Action Cancelled"
            Exit Function
        End If
        
        retWare = InputBox("Please change the ware code if necessary, for sherd group:  " & Unit & "." & WareCode & surftreat & ":", "Enter ware code", WareCode)
        
        If retWare = "" Then
            MsgBox "No ware code entered, action cancelled"
        Else
            retSurf = InputBox("Please change the surface treatment if necessary, for sherd group " & Unit & "." & WareCode & surftreat & ":", "Enter surface treatment", surftreat)
            If retSurf = "" Then
                MsgBox "No number entered, action cancelled"
            Else
               
                'now check if this sherd already exist already exists
                findGID = DLookup("[GID]", "[Q_BodySherd_GIDs]", "[GID] = '" & retUnit & "." & retWare & retSurf & "'")
                If Not IsNull(findGID) Then
                    MsgBox "Sorry but this Sherd GID (" & retUnit & "." & retWare & retSurf & ") already exists. You must delete it first before you can alter " & Unit & "." & WareCode & surftreat, vbExclamation, "Sherd GID already exists"
                    Exit Function
                Else
            
                    'new GID, does not exist so allow alteration
                    msg = "Are you quite sure that you want to renumber sherd " & Unit & "." & WareCode & surftreat & " to " & retUnit & "." & retWare & retSurf & "?"
                    Response = MsgBox(msg, vbExclamation + vbYesNoCancel, "Confirm Sherd Re-Number")
                    If Response = vbYes Then
                         
                        MsgBox "The re-numbering will now take place - it can be quite slow. Please just wait for the next message, the system has not hung", vbInformation, "Patience..."
                
                        On Error Resume Next
                        Dim mydb As DAO.Database, wrkdefault As Workspace, wrkdefault1 As Workspace
                        Set wrkdefault = DBEngine.Workspaces(0)
                        Set mydb = CurrentDb
        
                        '' Start of transaction.
                        wrkdefault.BeginTrans
                        
                        'these should all be done by cascade - but do here to
                        Call RenumARecord("Ceramics_Body_Sherd_Residues", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        Call RenumARecord("Ceramics_Body_Sherd_SampleNumbers", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        Call RenumARecord("Ceramics_Body_Sherd_SecondaryUse", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        'more complicated - it wants to do multiple updates to surface treatment table and can't do
                        'those in single transaction so must move out to trans
                        'Call UpdateSurfaceTreatment(surftreat, retSurf, Unit, warecode, mydb)
                        'Call RenumARecord("Ceramics_Body_Sherd_Surfacetreatment", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & warecode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        'Call RenumARecord("Ceramics_Body_Sherd_Surfacetreatment", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & warecode & "' AND [surfacetreatment] = '" & retSurf & "'", mydb)
                        
                        
                        Call RenumARecord("Ceramics_Body_Sherd_Technology", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        Call RenumARecord("Ceramics_Body_Sherd_Inclusion_Group", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        'this must be last has had it first before and it would stall on second change (timeout)
                        Call RenumARecord("Ceramics_Body_Sherds", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & surftreat & "'", mydb)
                        If Err.Number = 0 Then
                            wrkdefault.CommitTrans
                            
                            'cant do two processes on one table within transaction so if succeeds need to finally update surf treat table
                            'as unit or ware code might have altered - had to be done in 2 steps - hope it doesnt fail!!
                            Call UpdateSurfaceTreatment(surftreat, retSurf, Unit, WareCode, mydb)
                            Call RenumARecord("Ceramics_Body_Sherd_Surfacetreatment", "[Unit] = " & retUnit & ", [ware code] = '" & retWare & "', [surfacetreatment] = '" & retSurf & "'", "[Unit] = " & Unit & " AND [ware Code] = '" & WareCode & "' AND [surfacetreatment] = '" & retSurf & "'", mydb)
                            
                            Forms!Frm_BodySherd.Requery
                            Forms!Frm_BodySherd![cboFindUnit].Requery
                            Forms!Frm_BodySherd![cboFindUnit] = retUnit & "." & retWare & "-" & retSurf
                            MsgBox "Renumbering has been successful. Renumbered record will be displayed."
                        Else
                            wrkdefault.Rollback
                            MsgBox "A problem has occured and the re-numbering has been cancelled. The error message is: " & Err.Description
                        End If
                        mydb.Close
                        Set mydb = Nothing
                        wrkdefault.Close
                        Set wrkdefault = Nothing
                        
                    Else
                        MsgBox "Re-numbering cancelled", vbInformation, "Action Cancelled"
    
                    End If
                End If
            End If
        End If
    End If

Exit Function

err_cmdReNumber:
    Call General_Error_Trap
    Exit Function
End Function

Sub UpdateSurfaceTreatment(oldvalue, newvalue, Unit, WareCode, mydb)
Dim newval, Count, sql, wasPassedIn
Dim myq1 As QueryDef

'check for value
If newvalue = "" Or IsNull(newvalue) Then
    MsgBox "Surface Treatment must be entered"
    If oldvalue <> "" Then
        newvalue = oldvalue
    Else
        SendKeys "{ESC}"
        DoCmd.GoToControl "total"
        DoCmd.GoToControl "txtSurfTreat"
    End If
Else
    Dim checkexists
    checkexists = DLookup("[Unit]", "Ceramics_Body_Sherd_SurfaceTreatment", "[Unit] = " & Unit & " AND [Ware Code] ='" & WareCode & "' AND [SurfaceTreatment] = '" & oldvalue & "'")
    If Not IsNull(checkexists) Then
        'does exist alter
        '1.delete whats there
        If spString <> "" Then
            If IsNull(mydb) Then
                wasPassedIn = False
                'Dim mydb As DAO.Database
                Set mydb = CurrentDb
            Else
                wasPassedIn = True
            End If
    
            Set myq1 = mydb.CreateQueryDef("")
     
            If wasPassedIn = False Then
                myq1.Connect = spString
    
                myq1.ReturnsRecords = False
                myq1.sql = "sp_Ceramics_Delete_BodySherd_SurfaceTreatment " & Unit & ", '" & WareCode & "', " & oldvalue
            Else
                myq1.sql = "DELETE FROM [Ceramics_Body_Sherd_SurfaceTreatment] WHERE [Unit] = " & Unit & " AND [Ware Code] = '" & WareCode & "' AND surfacetreatment = '" & oldvalue & "';"
                myq1.Execute
            End If
            'myq1.Close
            'Set myq1 = Nothing
            
            '2. Add new value
            If InStr(newvalue, ",") > 0 Then
                '>1 treatment to add
                newval = Split(newvalue, ",")
                For Count = 0 To UBound(newval)
                    'sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & warecode & "', '" & newvalue & "'," & newval(count) & ")"
                    'DoCmd.RunSQL sql
                    myq1.sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & WareCode & "', '" & newvalue & "'," & newval(Count) & ")"
                    myq1.Execute
                Next

            Else
                'just one value
                'sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & warecode & "', '" & newvalue & "'," & newvalue & ")"
                'DoCmd.RunSQL sql
                myq1.sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & WareCode & "', '" & newvalue & "'," & newvalue & ")"
                myq1.Execute
            End If
            
             myq1.Close
            Set myq1 = Nothing
            
            If wasPassedIn = False Then
                mydb.Close
                Set mydb = Nothing
            End If
        Else
            'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
            MsgBox "The existing surface treatment record has not been deleted, please contact the administrator.", vbCritical, "Error"
            Exit Sub
        End If
    Else
        'does not exist - add
            If IsNull(mydb) Then
                wasPassedIn = False
                'Dim mydb As DAO.Database
                Set mydb = CurrentDb
            Else
                wasPassedIn = True
            End If
    
            Set myq1 = mydb.CreateQueryDef("")
        
            If InStr(newvalue, ",") > 0 Then
                '>1 treatment to add
                newval = Split(newvalue, ",")
                For Count = 0 To UBound(newval)
                    'sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & warecode & "', '" & newvalue & "'," & newval(count) & ")"
                    'DoCmd.RunSQL sql
                    myq1.sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & WareCode & "', '" & newvalue & "'," & newval(Count) & ")"
                    myq1.Execute
                Next

            Else
                'just one value
                'sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & warecode & "', '" & newvalue & "'," & newvalue & ")"
                'DoCmd.RunSQL sql
                myq1.sql = "INSERT INTO [Ceramics_Body_Sherd_SurfaceTreatment] ([Unit], [Ware code], [SurfaceTreatment], [IndividualTreatment]) VALUES (" & Unit & ", '" & WareCode & "', '" & newvalue & "'," & newvalue & ")"
                myq1.Execute
            End If
            
            myq1.Close
            Set myq1 = Nothing
            
            If wasPassedIn = False Then
                mydb.Close
                Set mydb = Nothing
            End If
    End If
    'Me!Frm_sub_bodysherd_surfacetreatment.Requery
    
End If
Exit Sub

err_SurfTreat:
    Call General_Error_Trap
    Exit Sub
    

End Sub

