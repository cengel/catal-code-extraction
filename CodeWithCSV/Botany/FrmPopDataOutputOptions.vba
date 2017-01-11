Option Compare Database
Option Explicit

Private Sub cmdAction_Click()
'************************************************************************
' Process output depending on values selected on screen
' SAJ
'************************************************************************
On Error GoTo err_cmdAction
Dim which, where, sql, rpt

    If Me![optgrpWhich] = 1 Then
        'report
        which = "report"
    ElseIf Me![optgrpWhich] = 2 Then
        'excel
        which = "excel"
    Else
        MsgBox "Please select to either report or export data to excel", vbInformation, "No action to carry out"
        which = ""
        Exit Sub
    End If

   If which <> "" Then
        If Me![optgrpData] = 1 Then
            'current only
            where = "[Flot Number] = " & Me![txtRec]
        ElseIf Me![optgrpData] = 2 Then
            'range
            If Me![cboStart] = "" Or Me![cboEnd] = "" Then
                MsgBox "Please enter both a Start and End value for the range of records to output.", vbInformation, "Insufficient Data"
                Exit Sub
            Else
                where = "[Flot Number] BETWEEN " & Me![cboStart] & " AND " & Me![cboEnd]
            End If
        ElseIf Me![optgrpData] = 3 Then
            'all
            where = ""
        End If
   
        sql = "SELECT * FROM [" & Me![txtData] & "]"
        If where <> "" Then sql = sql & " where " & where
        sql = sql & ";"
   
   
        If which = "report" Then
            If Me![txtData] = "Bot: Basic Data" Or Me![txtData] = "Q_ExportBasicData_AllRecs_withUnit" Then
                rpt = "R_BasicData"
                DoCmd.OpenReport rpt, acViewPreview, , where
                'DoCmd.SelectObject acReport, "R_BasicData", True
            ElseIf Me![txtData] = "Bot: Priority Sample" Then
                rpt = "R_PrioritySample"
                DoCmd.OpenReport rpt, acViewPreview, , where
            ElseIf Me![txtData] = "Bot: SieveScanning" Then
                rpt = "R_SieveScanning"
                DoCmd.OpenReport rpt, acViewPreview, , where
            ElseIf Me![txtData] = "Bot: Sample Scanning" Then
                rpt = "R_SampleScanning"
                DoCmd.OpenReport rpt, acViewPreview, , where
            ElseIf Me![txtData] = "Bot: Priority Report" Then
                rpt = "R_PriorityReport"
                DoCmd.OpenReport rpt, acViewPreview, , where
            Else
                MsgBox "Sorry but the table name passed into this form cannot be matched with a report. Please contact the database administrator", vbCritical, "Report cannot be found"
            End If
            DoCmd.Close acForm, Me.Name
            'rpt.SetFocus
        ElseIf which = "excel" Then
            Dim mydb As Database, myq As QueryDef
            Set mydb = CurrentDb
            Set myq = mydb.CreateQueryDef("ArchBotExcelExport")
            
            myq.sql = sql
            myq.ReturnsRecords = False
            
            DoCmd.OutputTo acOutputQuery, "ArchBotExcelExport", acFormatXLS, , True
            
            mydb.QueryDefs.Delete ("ArchBotExcelExport")
            
            myq.Close
            Set myq = Nothing
            mydb.Close
            Set mydb = Nothing
            
            DoCmd.Close acForm, Me.Name
        End If
   End If
Exit Sub

err_cmdAction:
    If Err.Number = 3012 Then
        'query already exists
        mydb.QueryDefs.Delete ("ArchBotExcelExport")
        Resume
    Else
        Call General_Error_Trap
    End If
    Exit Sub

End Sub

Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' This form allows the user to select what data they wish to output and in
' what format. OpenArgs are required to specify where the call to the form
' was made from (ie: what table was viewed) and what the current record was.
' This must take the format: table;record
' SAJ
'**********************************************************************
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    Dim data, rec
    data = Left(Me.OpenArgs, InStr(Me.OpenArgs, ";") - 1)
    rec = Right(Me.OpenArgs, Len(Me.OpenArgs) - InStr(Me.OpenArgs, ";"))
    
    Me![txtRec] = rec
    Me![txtData] = data
    
    Me![cboStart].RowSource = "SELECT DISTINCT [Flot Number] FROM [" & data & "] ORDER BY [Flot Number];"
    Me![cboEnd].RowSource = "SELECT DISTINCT [Flot Number] FROM [" & data & "] ORDER BY [Flot Number];"

Else
    MsgBox "This form has been called without the necessary parameters, it will now close", vbCritical, "Insufficient Parameters"
    DoCmd.Close acForm, Me.Name
End If


Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub optgrpData_Click()
'***********************************************************************
' enable/disable range combos depending on value selected
' SAJ
'***********************************************************************
On Error GoTo err_optgrpData

If Me![optgrpData] = 2 Then
    Me![cboStart].Enabled = True
    Me![cboEnd].Enabled = True
Else
    Me![cboStart].Enabled = False
    Me![cboEnd].Enabled = False
End If

Exit Sub

err_optgrpData:
    Call General_Error_Trap
    Exit Sub
End Sub
