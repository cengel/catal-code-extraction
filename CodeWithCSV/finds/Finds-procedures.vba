Option Compare Database
Option Explicit

Function DeleteCrateRecord(num, mydb) As Boolean
'when something is moved from crate to crate it must be inserted then deleted but RW users
'don't have permissions to delete so need to use SP to do so
'On Error GoTo err_delrec

If spString <> "" Then
    Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = spString
    myq1.ReturnsRecords = False
    myq1.sql = "sp_Store_Delete_CrateEntry " & num
    myq1.Execute
    myq1.Close
    Set myq1 = Nothing
    
    DeleteCrateRecord = True

Else
    MsgBox "Sorry but the record cannot be deleted out of the this crate, restart the database and try again", vbCritical, "Error"
    DeleteCrateRecord = False
End If
Exit Function

'err_delrec:
'    Call General_Error_Trap
'    Exit Function
End Function

