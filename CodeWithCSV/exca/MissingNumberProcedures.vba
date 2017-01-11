Option Compare Database
Option Explicit

Function FindMissingNumbers(tbl, fld) As Boolean
'*****************************************************************************************
' New procedure Feb 09 to aid data cleaning - SAJ
' Identifies numbers that are missing in a particular table - for use with units, features etc
' It can only know a number is missing at present, it has no master table to compare with
' yet so will come up with runs of numbers not used at present.
' The code is called from the form: Exca: Admin_Subform_MissingNumbers
' Inputs: table to check, field name that contains the number
'*****************************************************************************************
On Error GoTo err_nums

Dim mydb As DAO.Database, myrs As DAO.Recordset
Dim sql As String, sql1 As String, val As Field, holdval1 As Long, holdval2 As Long
Dim response As Integer
MsgBox "The first thing this code must do is retrieve the whole dataset. If your connection is slow it may time out but it will give you a message if this happens. Starting now......", vbInformation, "Start Procedure"
Set mydb = CurrentDb
'first get the last number in the table - often the end numbers in these tables are mistakes
'so the idea here is to show the user the last few numbers so they can decide where the genuine end point it.
' eg: in units there is 999999 (the Mellaart number) and all other 900000 should not be listed
sql = "SELECT [" & fld & "] FROM [" & tbl & "] ORDER BY [" & fld & "];"
Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
myrs.MoveLast
Set val = myrs.Fields(fld)
holdval1 = val
'now get the number 2nd from last
myrs.MovePrevious
holdval2 = val

'show the user the last two numbers in the sequence and ask if the last number is the end of the range
response = MsgBox("The last two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
If response = vbNo Then
    'if user does not want to use last number as end of range then move back through the recordset asking the same question for
    ' four more iterations before giving up - this covers the incorrect numbers identified in the unit sheet whilst writing the code
    holdval1 = holdval2
    myrs.MovePrevious
    holdval2 = val
    response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
    If response = vbNo Then
        holdval1 = holdval2
        myrs.MovePrevious
        holdval2 = val
        response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
        If response = vbNo Then
            holdval1 = holdval2
            myrs.MovePrevious
            holdval2 = val
            response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
            If response = vbNo Then
                holdval1 = holdval2
                myrs.MovePrevious
                holdval2 = val
                response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
                If response = vbNo Then
                    MsgBox "Please clean up the last " & fld & " values and run this procedure again"
                    FindMissingNumbers = False
                    Exit Function
                Else
                    GoTo cont
                End If
            Else
                GoTo cont
            End If

        Else
            GoTo cont
        End If
    Else
        GoTo cont
    End If


Else
    GoTo cont
End If

cont:

    MsgBox "The code will now run to compile the list of missinig numbers up to: " & holdval1 & ". It may be quite slow , a report will appear when complete so you know it has finished"
    sql1 = "DELETE * FROM LocalMissingNumbers;"
    DoCmd.RunSQL sql1
    sql1 = ""
    myrs.MoveFirst
    Dim counter As Long, checknum
    counter = 0
    Do Until counter = holdval1
        checknum = DLookup("[" & fld & "]", "[" & tbl & "]", "[" & fld & "] = " & counter)
        If IsNull(checknum) Then
            sql1 = "INSERT INTO [LocalMissingNumbers] (MissingNumber) VALUES (" & counter & ");"
            DoCmd.RunSQL sql1
        End If
        ''myrs.MoveNext
        counter = counter + 1
    Loop
'sql1 =

myrs.Close
Set myrs = Nothing
mydb.Close
Set mydb = Nothing
FindMissingNumbers = True
Exit Function

err_nums:
    Call General_Error_Trap
    Exit Function
End Function

