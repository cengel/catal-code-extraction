Option Compare Database
Option Explicit


Function StartUp()
'*****************************************************************************
' All actions necessary to start the system as smoothly as possible
'
' SAJ v.9 - actions from autoexac macro plus addition of login/attachment check
'*****************************************************************************
On Error GoTo err_startup

'DoCmd.RunCommand acCmdWindowHide 'hide the DB window from prying eyes
'**** opening main menu moved from here lower down code

'moved to being called from form Excavation login below, opened as dialog
'If LogUserIn = True Then 'function in module TableLinkingProcedures - get user to login
'    'if true then login ok and tables accessible - ready to go
'Else
'    'if loginuser = false then the system should have quit by now, this is a catchall
'    MsgBox "The system may not of logged into the database correctly. If you encounter any problems please restart the application"
'End If

DoCmd.OpenForm "FRM_Login", acNormal, , , acFormEdit, acDialog

'you can hide the warning messages that Access popups up when
'you do sql tasks in the background - however the negative side to
'this is that you hide all these types of message which you may not
'want to do - the options you have are:
'   DoCmd.SetWarnings False 'turns off macro msgs
'   Application.SetOption "Confirm Record Changes", False
'   Application.SetOption "Confirm Document Deletions", False
    Application.SetOption "Confirm Action Queries", False  'this will hide behind the scences sql actions
'you could of course turn this on an off around each statement - I'm undecided at present

'now the tables are all ok find out the current version
SetCurrentVersion

'**** open move from marked place above
DoCmd.OpenForm "FRM_MAINMENU", acNormal, , , acFormReadOnly 'open main menu
'DoCmd.Maximize 'really, should we? 'no I don't think so
'refresh the main menu so the version number appears
Forms![FRM_MAINMENU].Refresh

Exit Function

err_startup:
    Call General_Error_Trap
    'now should the system quit out here?
    'to be decided
End Function

Function CheckIfLOVValueUsed(LOVName, LOVField, LOVValue, CheckTable, CheckTableKeyField, CheckTableField, task, Optional extracrit)
'******************************************************************************
' This function is used by the Administration area of the site - it checks if
' a LOV value can be edited or deleted by checking dependant tables
' Inputs:   LOVName = lov table name
'           LOVField = lov field name being checked
'           LOVVAlue = LOV value being checked out
'           CheckTable = dependant table name to check if value exists in
'           CheckTableKeyField = key of dependant table
'           CheckTAbleField = field name where LOV value stored in dependant table
'           task = edit or delete
'           extracrit = any extra criteria for record search, optional
' Outputs:  msg back to user or OK
' v9.2 SAJ
'*****************************************************************************
On Error GoTo err_CheckIFLOVValueUsed

'only proceed if all inputs are present
If LOVName <> "" And LOVField <> "" And LOVValue <> "" And CheckTable <> "" And CheckTableKeyField <> "" And CheckTableField <> "" And task <> "" Then

    Dim mydb As Database, myrs As Recordset, sql As String, msg As String, msg1 As String, keyfld As Field, Count As Integer
    Set mydb = CurrentDb
    
    If Not IsMissing(extracrit) Then
        sql = "SELECT [" & CheckTableKeyField & "], [" & CheckTableField & "] FROM [" & CheckTable & "] WHERE [" & CheckTableField & "] = '" & LOVValue & "' " & extracrit & " ORDER BY [" & CheckTableKeyField & "];"
    Else
        sql = "SELECT [" & CheckTableKeyField & "], [" & CheckTableField & "] FROM [" & CheckTable & "] WHERE [" & CheckTableField & "] = '" & LOVValue & "' ORDER BY [" & CheckTableKeyField & "];"
    End If
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)

    If myrs.BOF And myrs.EOF Then
        msg = "ok"
    Else
        myrs.MoveFirst
        Count = 0
        msg = "You cannot " & task & " this " & LOVField & " because the following records in the table " & CheckTable & " use it: "
        msg1 = ""
        Do Until myrs.EOF
            Set keyfld = myrs.Fields(CheckTableKeyField)
            If msg1 <> "" Then msg1 = msg1 & ", "
            msg1 = msg1 & keyfld
            Count = Count + 1
            'there is a limit to amount of text in msgbox so may not be able to show more
            If Count > 50 Then
                msg1 = msg1 & ".....etc"
                Exit Do
            End If
        myrs.MoveNext
        Loop
        
        msg = msg & Chr(13) & Chr(13) & CheckTableKeyField & ": " & msg1
        If task = "edit" Then
            msg = msg & Chr(13) & Chr(13) & "It is suggested you add a new " & LOVField & " to the list and then change all records that refer to '"
            msg = msg & LOVValue & "' to your new " & LOVField & ". You will then be able to delete it from the list."
        ElseIf task = "delete" Then
             msg = msg & Chr(13) & Chr(13) & "You must change all records that refer to this " & LOVField
            msg = msg & " '" & LOVValue & "' before you will be able to delete it from the list."
        End If
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    
    CheckIfLOVValueUsed = msg
Else
    CheckIfLOVValueUsed = "fail"
End If
Exit Function

err_CheckIFLOVValueUsed:
    Call General_Error_Trap
    Exit Function
End Function


Function AdminDeletionCheck(CheckTable, CheckField, CheckVal, Term, retField)
'******************************************************************************
' This function is used by the Administration area of the site - it checks if
' a LOV value can be edited or deleted by checking dependant tables
' Inputs:   CheckTable = table to check if val is used in
'           CheckField = field to check value against
'           CheckVal = value to be checked for existent in the field Checkfield within CheckTable
'           Term = user friendly term for object being checked
'           retField = field value to return eg: if looking for units in features return the feature number
' Outputs:  msg back to user or ""
' v9.2 SAJ
'*****************************************************************************
On Error GoTo err_AdminDeletionCheck

'only proceed if all inputs are present
If CheckTable <> "" And CheckField <> "" And CheckVal <> "" And Term <> "" Then

    Dim mydb As Database, myrs As Recordset, sql As String, msg As String, msg1 As String, keyfld As Field, Count As Integer
    Set mydb = CurrentDb
    
    If CheckTable = "Exca: stratigraphy" And CheckField = "To_units" Then
        sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [" & CheckField & "] = '" & CheckVal & "';"
    
    ElseIf CheckTable = "Exca: graphics list" Then 'graphics needs to define if feature num or unit
        If CheckField = "Unit" Then
           sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [Unit/feature number] = " & CheckVal & " AND lcase([Feature/Unit]) =  'u';"
        ElseIf CheckField = "Feature" Then
            sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [Unit/feature number] = " & CheckVal & " AND lcase([Feature/Unit]) =  'f';"
        End If
    Else
        sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [" & CheckField & "] = " & CheckVal & ";"
    End If
    
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)

    If myrs.BOF And myrs.EOF Then
        msg = ""
    Else
        myrs.MoveFirst
        Count = 0
        msg = Term & ": "
        msg1 = ""
        Do Until myrs.EOF
            Set keyfld = myrs.Fields(retField)
            If msg1 <> "" Then msg1 = msg1 & ", "
            msg1 = msg1 & keyfld
            Count = Count + 1
            'there is a limit to amount of text in msgbox so may not be able to show more
            If Count > 50 Then
                msg1 = msg1 & ".....etc"
                Exit Do
            End If
        myrs.MoveNext
        Loop
        
        msg = msg & msg1
        
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    
    AdminDeletionCheck = msg
Else
    AdminDeletionCheck = ""
End If
Exit Function

err_AdminDeletionCheck:
    Call General_Error_Trap
    Exit Function
End Function

Sub DeleteARecord(FromTable, FieldName, FieldValue, Text, mydb)
'This is an admin function to delete records, used within a transaction if must not
'be error trapped. The db workspace is passed in for the transaction
'Inputs -   Fromtable - delete from what table
'           Fieldname - field to id records to delete
'           fieldvalue - criteria to delete
'           Text - true = text field that requires '' around it
'           mydb - currentdb connection sent thro for transaction


Dim sql, myq As QueryDef
Set myq = mydb.CreateQueryDef("")
           
        If Text = False Then
            If FromTable = "Exca: graphics list" Then 'graphics needs to define if feature num or unit
                If FieldName = "Unit" Then
                    sql = "DELETE FROM [" & FromTable & "] WHERE [Unit/feature number] = " & FieldValue & " AND lcase([Feature/Unit]) =  'u';"
                ElseIf FieldName = "Feature" Then
                    sql = "DELETE FROM [" & FromTable & "] WHERE [Unit/feature number] = " & FieldValue & " AND lcase([Feature/Unit]) =  'f';"
                End If
        
            Else
                sql = "DELETE FROM [" & FromTable & "] WHERE [" & FieldName & "] = " & FieldValue & ";"
            End If
        Else
            sql = "DELETE FROM [" & FromTable & "] WHERE [" & FieldName & "] = '" & FieldValue & "';"
        End If
                
        myq.sql = sql
        myq.Execute
                
myq.Close
Set myq = Nothing


End Sub
Sub RenameLinks()
'when new tables are linked in from sql server they come with the owner
'name prefixed to it - remove this. SAJ
On Error GoTo err_rename
Dim mydb As DAO.Database, I, newName
Dim tmptable As TableDef
Set mydb = CurrentDb
    

For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
         Set tmptable = mydb.TableDefs(I)
             
        If tmptable.Connect <> "" Then
            Debug.Print tmptable.Name
            newName = Replace(tmptable.Name, "dbo_", "")
            tmptable.Name = newName
            Debug.Print tmptable.Name
        End If
Next

Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Sub

err_rename:
    MsgBox Err.Description
  '  Resume
    Exit Sub
End Sub

