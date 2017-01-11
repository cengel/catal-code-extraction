Option Compare Database
Option Explicit

'****************************************************************
' Archaeobots specific procedures
'
' SAJ March 2006
'****************************************************************

Function AddRecordToPriorityTable(frm As Form) As Boolean
'*******************************************
' Add the flot record to the priority table
' SAJ
'*******************************************
On Error GoTo err_AddRecordToPriorityTable
Dim sql, msg, retVal

If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
    'enough info to add record to Priority table
    'new season 2006 - new records if entered in as a range also go auto into scanning table so
    'now need to check if the record is in scanning before moving to priority
    If frm![chkInScanning] = True Then
        'new request from AB, check if 1mm random split field filled out if so don't delete
        Dim checkSplit
        checkSplit = DLookup("[1 mm random split]", "[Bot: Sample Scanning]", "[Flot Number] = " & frm![Flot Number])
        If Not IsNull(checkSplit) Then
            retVal = MsgBox("A sample scanning record exists for this Flot that contains data, do you want to delete it from the Scanning table now?", vbInformation + vbYesNo, "Scanning Record Exists")
           
            If retVal = vbYes Then
                'call sp to delete this record as rw users no permission to delete
                If DeleteSampleRecord(frm![Flot Number]) = False Then
                    MsgBox "The deletion failed", vbCritical, "Error"
                End If
            End If
            'retVal = MsgBox("This operation will remove the Scanning record for this Flot, are you sure you want to continue?", vbYesNo + vbQuestion, "Remove from Scanning")
        
            'If retVal = vbNo Then
            '    AddRecordToPriorityTable = False
            '    Exit Function
            'Else
            'call sp to delete this record as rw users no permission to delete
            '    If DeleteSampleRecord(frm![Flot Number]) = False Then
            '    AddRecordToPriorityTable = False
            '    Exit Function
            'End If
        Else
            DeleteSampleRecord (frm![Flot Number])
        End If
    End If
    'all is well carry on
    Application.SetOption "Confirm Action Queries", False 'turn flag off as if you get append msg and press No it seems to get in a mess
    sql = "INSERT INTO [Bot: Priority Sample] ([Flot Number], [Year]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
    DoCmd.RunSQL sql
    Application.SetOption "Confirm Action Queries", True
    'throwing error that form is not bound so remove where clause - not always throwing error!
    'DoCmd.OpenForm "FrmPriority", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmPriority", acNormal, , , , , frm![Flot Number]
    AddRecordToPriorityTable = True
        
        
        
Else
    'not enough info to add record
    msg = "This record cannot be entered into the Priority Sample table until the following values are present:" & Chr(13) & Chr(13)
    If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
        msg = msg & " Flot number " & Chr(13)
    ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
        msg = msg & " Float Date "
    End If
    MsgBox msg, vbExclamation, "Insufficient Information"
    AddRecordToPriorityTable = False
End If

Exit Function
err_AddRecordToPriorityTable:
    Call General_Error_Trap
    Exit Function
End Function

Function AddRecordToScanTable(frm As Form) As Boolean
'*******************************************
' Add the flot record to the scanning table
' SAJ
'*******************************************
On Error GoTo err_AddRecordToScanTable
Dim sql, msg
If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
    'enough info to add record to Priority table
    sql = "INSERT INTO [Bot: Sample Scanning] ([Flot Number], [YearScanned]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
    DoCmd.RunSQL sql
    'throwing error that form is not bound so remove where clause - not always throwing error!
    'DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmSampleScan", acNormal, , , , , frm![Flot Number]
    AddRecordToScanTable = True
Else
    'not enough info to add record
    msg = "This record cannot be entered into the Sample Scanning table until the following values are present:" & Chr(13) & Chr(13)
    If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
        msg = msg & " Flot number " & Chr(13)
    ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
        msg = msg & " Float Date "
    End If
    MsgBox msg, vbExclamation, "Insufficient Information"
    AddRecordToScanTable = False
End If

Exit Function
err_AddRecordToScanTable:
    Call General_Error_Trap
    Exit Function

End Function


Function AddRecordToSieveScanTable(frm As Form) As Boolean
'*******************************************
' Add the flot record to the scanning table
' DL 2015 (taken from SAJ)
'*******************************************
On Error GoTo err_AddRecordToSieveScanTable
Dim sql, msg
If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
    'enough info to add record to Priority table
    sql = "INSERT INTO [Bot: SieveScanning] ([Flot Number], [Year_scanned]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
    DoCmd.RunSQL sql
    'throwing error that form is not bound so remove where clause - not always throwing error!
    'DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmSieveScan", acNormal, , , , , frm![Flot Number]
    AddRecordToSieveScanTable = True
Else
    'not enough info to add record
    msg = "This record cannot be entered into the Sieve Scanning table until the following values are present:" & Chr(13) & Chr(13)
    If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
        msg = msg & " Flot number " & Chr(13)
    ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
        msg = msg & " Float Date "
    End If
    MsgBox msg, vbExclamation, "Insufficient Information"
    AddRecordToSieveScanTable = False
End If

Exit Function
err_AddRecordToSieveScanTable:
    Call General_Error_Trap
    Exit Function

End Function

Function ViewPriorityRecord(frm As Form) As Boolean
'*******************************************
' Open the priority record for this the flot num
' SAJ
'*******************************************
On Error GoTo err_GoToPriority_click
Dim sql, msg

If frm![chkInPriority] = True Then
    'go to Priority form
    'form is not bound where clause not needed
    'DoCmd.OpenForm "FrmPriority", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmPriority", acNormal, , , , , frm![Flot Number]
    ViewPriorityRecord = True
Else
    'error, is not flagged in priority record here so better check
    Dim there
    there = DLookup(frm![Flot Number], "[Bot: Priority Sample]", "[Flot Number] = " & frm![Flot Number])
    If IsNull(there) Then
        'number not in table insert it
    
        If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
            sql = "INSERT INTO [Bot: Priority Sample] ([Flot Number], [Year]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
            DoCmd.RunSQL sql
            DoCmd.OpenForm "FrmPriority", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
            ViewPriorityRecord = True
        Else
            msg = "The record is not actually in the Priority table and cannot be entered into it until the following values are present:" & Chr(13) & Chr(13)
            If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
                msg = msg & " Flot number " & Chr(13)
            ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
                msg = msg & " Float Date "
            End If
            MsgBox msg, vbExclamation, "Insufficient Information"
            ViewPriorityRecord = False
        End If
    Else
        'number is there so open form
        DoCmd.OpenForm "FrmPriority", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
        ViewPriorityRecord = True
    End If
End If

Exit Function
err_GoToPriority_click:
    Call General_Error_Trap
    Exit Function

End Function

Function ViewScanRecord(frm As Form) As Boolean
'*******************************************
' Open the sample scan record for this the flot num
' SAJ
'*******************************************
On Error GoTo err_GoToScanning_click
Dim sql, msg

If frm![chkInScanning] = True Then
    'go to Priority form
    'form is not bound, where clause not used
    'DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmSampleScan", acNormal, , , , , frm![Flot Number]
    ViewScanRecord = True
Else
    'error, is not flagged in priority record here so better check
    Dim there
    there = DLookup(frm![Flot Number], "[Bot: Sample Scanning]", "[Flot Number] = " & frm![Flot Number])
    If IsNull(there) Then
        'number not in table insert it
    
        If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
            sql = "INSERT INTO [Bot: Sample Scanning] ([Flot Number], [YearScanned]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
            DoCmd.RunSQL sql
            DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
            ViewScanRecord = True
        Else
            msg = "The record is not actually in the Scanning table and cannot be entered into it until the following values are present:" & Chr(13) & Chr(13)
            If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
                msg = msg & " Flot number " & Chr(13)
            ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
                msg = msg & " Float Date "
            End If
            MsgBox msg, vbExclamation, "Insufficient Information"
            ViewScanRecord = False
        End If
    Else
        'number is there so open form
        DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
        ViewScanRecord = True
    End If
End If

Exit Function
err_GoToScanning_click:
    Call General_Error_Trap
    Exit Function


End Function

Function ViewSieveScanRecord(frm As Form) As Boolean
'*******************************************
' Open the Sieve scan record for this the flot num
' DL 2015 (taken from SAJ)
'*******************************************
On Error GoTo err_GoToSieveScanning_click
Dim sql, msg

If frm![chkinsievescanning] = True Then
    'go to Priority form
    'form is not bound, where clause not used
    'DoCmd.OpenForm "FrmSampleScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    DoCmd.OpenForm "FrmSieveScan", acNormal, , , , , frm![Flot Number]
    ViewSieveScanRecord = True
Else
    'error, is not flagged in priority record here so better check
    Dim there
    there = DLookup(frm![Flot Number], "[Bot: SieveScanning]", "[Flot Number] = " & frm![Flot Number])
    If IsNull(there) Then
        'number not in table insert it
    
        If frm![Float Date] <> "" And frm![Flot Number] <> "" Then
            sql = "INSERT INTO [Bot: SieveScanning] ([Flot Number], [Year_scanned]) VALUES (" & frm![Flot Number] & ", " & Year(frm![Float Date]) & ");"
            DoCmd.RunSQL sql
            DoCmd.OpenForm "FrmSieveScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
            ViewSieveScanRecord = True
        Else
            msg = "The record is not actually in the Sieve Scanning table and cannot be entered into it until the following values are present:" & Chr(13) & Chr(13)
            If frm![Flot Number] = "" Or IsNull(frm![Flot Number]) Then
                msg = msg & " Flot number " & Chr(13)
            ElseIf frm![Float Date] = "" Or IsNull(frm![Float Date]) Then
                msg = msg & " Float Date "
            End If
            MsgBox msg, vbExclamation, "Insufficient Information"
            ViewSieveScanRecord = False
        End If
    Else
        'number is there so open form
        DoCmd.OpenForm "FrmSieveScan", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
        ViewSieveScanRecord = True
    End If
End If

Exit Function
err_GoToSieveScanning_click:
    Call General_Error_Trap
    Exit Function


End Function


Function DeleteSampleRecord(num) As Boolean
'all basic records are auto put into sample scanning but if they require to go in
'priority need to delete out of sample. RW users don't have permissions to delete so
'need to use SP to do so
On Error GoTo err_delrec

If spString <> "" Then
    Dim mydb As DAO.Database
    Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = spString
    myq1.ReturnsRecords = False
    myq1.sql = "sp_Bot_Delete_SampleScanRecord " & num
    myq1.Execute
    myq1.Close
    Set myq1 = Nothing
    mydb.Close
    Set mydb = Nothing
    
    DeleteSampleRecord = True

Else
    MsgBox "Sorry but the record cannot be deleted out of the sample scanning table, restart the database and try again", vbCritical, "Error"
    DeleteSampleRecord = False
End If
Exit Function

err_delrec:
    Call General_Error_Trap
    Exit Function
End Function

Function AddRecordToPriorityReport(frm As Form) As Boolean
'*******************************************
' Add the flot record to the priority report table
' SAJ
'*******************************************
On Error GoTo err_AddRecordToPriorityReport
Dim sql, msg
If frm![Flot Number] <> "" Then
    'enough info to add record to Priority table
    sql = "INSERT INTO [Bot: Priority Report] ([Flot Number]) VALUES (" & frm![Flot Number] & ");"
    DoCmd.RunSQL sql
    DoCmd.OpenForm "FrmPriorityReport", acNormal, , "[Flot Number] = " & frm![Flot Number], , , frm![Flot Number]
    AddRecordToPriorityReport = True
Else
    'not enough info to add record
    msg = "This record cannot be entered into the Priority Report table until a Flot number has been entered"
    MsgBox msg, vbExclamation, "Insufficient Information"
    AddRecordToPriorityReport = False
End If

Exit Function
err_AddRecordToPriorityReport:
    Call General_Error_Trap
    Exit Function

End Function
