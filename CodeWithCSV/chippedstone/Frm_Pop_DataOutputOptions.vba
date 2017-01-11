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
            If Me![txttable] = "basic" Then
                where = "[Q_BasicData].[BagNo] = '" & Me![txtRec] & "'"
            Else
                where = "[Q_StageTwo].[GID] = '" & Me![txtRec] & "'"
            End If
        ElseIf Me![optgrpData] = 2 Then
            'range
            If Me![cboStart] = "" Or Me![cboEnd] = "" Then
                MsgBox "Please enter both a Start and End value for the range of records to output.", vbInformation, "Insufficient Data"
                Exit Sub
            Else
                If Me![txttable] = "basic" Then
                    where = "[Q_BasicData].[Bagno] BETWEEN '" & Me![cboStart] & "' AND '" & Me![cboEnd] & "'"
                Else
                    where = "[Q_StageTwo].[GID] BETWEEN '" & Me![cboStart] & "' AND '" & Me![cboEnd] & "'"
                End If
            End If
        ElseIf Me![optgrpData] = 3 Then
            'all
            where = ""
        ElseIf Me![optgrpData] = 4 Then
            'current only
            If Me![txttable] = "basic" Then
                where = "[Q_BasicData].[Unit] = " & Forms![frm_CS_basicdata]![Unit]
            Else
                where = "[Q_StageTwo].[Unit] = '" & Forms![frm_CS_stagetwo]![Unit]
            End If
        End If
   
        If Me![txttable] = "basic" Then
            sql = "SELECT * FROM [Q_BasicData]"
        Else
            sql = "SELECT * FROM [Q_StageTwo]"
        End If
        If where <> "" Then sql = sql & " where " & where
        sql = sql & ";"
   
   
        If which = "report" Then
            If Me![txttable] = "basic" Then
                rpt = "R_BasicData"
                DoCmd.OpenReport rpt, acViewPreview, , where
                'DoCmd.SelectObject acReport, "R_BasicData", True
            Else
                rpt = "R_StageTwo"
                DoCmd.OpenReport rpt, acViewPreview, , where
            'Else
            '    MsgBox "Sorry but the table name passed into this form cannot be matched with a report. Please contact the database administrator", vbCritical, "Report cannot be found"
            End If
            MsgBox "Sorry the report will have appeared behind the form, click on it to bring it to the front" & Chr(13) & Chr(13) & "This is an outstanding problem.", vbInformation, "Report Location"
            DoCmd.Close acForm, Me.Name
            'rpt.SetFocus
        ElseIf which = "excel" Then
            Dim mydb As Database, myq As QueryDef
            Set mydb = CurrentDb
            Set myq = mydb.CreateQueryDef("CSExcelExport")
            
            myq.sql = sql
            myq.ReturnsRecords = False
            
            DoCmd.OutputTo acOutputQuery, "CSExcelExport", acFormatXLS, , True
            
            mydb.QueryDefs.Delete ("CSExcelExport")
            
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
        mydb.QueryDefs.Delete ("CSExcelExport")
        Resume
    Else
        Call General_Error_Trap
    End If
    Exit Sub

End Sub

Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' This form allows the user to select what data they wish to output and in
' what format. OpenArgs are required to specify what the current record was
' and which table to output
' This must take the format: record;table
' SAJ
'**********************************************************************
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    Dim tbl, rec
    rec = Left(Me.OpenArgs, InStr(Me.OpenArgs, ";") - 1)
    tbl = Right(Me.OpenArgs, Len(Me.OpenArgs) - InStr(Me.OpenArgs, ";"))
    
    Me![txtRec] = rec
    Me![txttable] = tbl
    
    If tbl = "basic" Then
        Me![cboStart].RowSource = "SELECT DISTINCT [BagNo] FROM [ChippedStone_Basic_Data] ORDER BY [BagNo];"
        Me![cboEnd].RowSource = "SELECT DISTINCT [BagNo] FROM [ChippedStone_Basic_Data] ORDER BY [BagNo];"
    Else
        'stage two
        Me![cboStart].RowSource = "SELECT DISTINCT [GID] FROM [ChippedStone_StageTwo_Data] ORDER BY [GID];"
        Me![cboEnd].RowSource = "SELECT DISTINCT [GID] FROM [ChippedStone_StageTwo_Data] ORDER BY [GID];"
    End If
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
