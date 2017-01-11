Option Compare Database
Option Explicit

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

    Dim mydb As Database, myrs As Recordset, sql As String, msg As String, msg1 As String, keyfld As Field, count As Integer
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
        count = 0
        msg = "You cannot " & task & " this " & LOVField & " because the following records in the table " & CheckTable & " use it: "
        msg1 = ""
        Do Until myrs.EOF
            Set keyfld = myrs.Fields(CheckTableKeyField)
            If msg1 <> "" Then msg1 = msg1 & ", "
            msg1 = msg1 & keyfld
            count = count + 1
            'there is a limit to amount of text in msgbox so may not be able to show more
            If count > 50 Then
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
