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
            where = "[dbo_Groundstone: Basic_Data_Pre2013_WithReRecordingFlag].[GID] = '" & Me![txtRec] & "'"
        ElseIf Me![optgrpData] = 2 Then
            'range
            If Me![cboStart] = "" Or Me![cboEnd] = "" Then
                MsgBox "Please enter both a Start and End value for the range of records to output.", vbInformation, "Insufficient Data"
                Exit Sub
            Else
                where = "[dbo_Groundstone: Basic_Data_Pre2013_WithReRecordingFlag].[GID] BETWEEN '" & Me![cboStart] & "' AND '" & Me![cboEnd] & "'"
            End If
        ElseIf Me![optgrpData] = 3 Then
            'all
            where = ""
        End If
   
        
        sql = "SELECT * FROM [Q_GS_Basic_with_Excavation]"
        
        If where <> "" Then sql = sql & " where " & where
        sql = sql & ";"
   
   
        If which = "report" Then
            rpt = "R_Basic"
            DoCmd.OpenReport rpt, acViewPreview, , where
         
            'MsgBox "Sorry the report will have appeared behind the form, click on it to bring it to the front" & Chr(13) & Chr(13) & "This is an outstanding problem.", vbInformation, "Report Location"
            Reports![R_Basic].SetFocus
            
            'DoCmd.Close acForm, Me.Name
            DoCmd.Close acForm, "Frm_Pop_DataOutputOptions"
            
            'rpt.SetFocus
        ElseIf which = "excel" Then
            Dim mydb As Database, myq As QueryDef
            Set mydb = CurrentDb
            Set myq = mydb.CreateQueryDef("GSExcelExport")
            
            myq.sql = sql
            myq.ReturnsRecords = False
            
            DoCmd.OutputTo acOutputQuery, "GSExcelExport", acFormatXLS, , True
            
            mydb.QueryDefs.Delete ("GSExcelExport")
            
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
        mydb.QueryDefs.Delete ("GSExcelExport")
        Resume
    ElseIf Err.Number = 2465 Then
        'application error when try to give form the focus
        ''Resume
        DoCmd.Close acForm, Me.Name
    Else
        Call General_Error_Trap
    End If
    Exit Sub

End Sub

Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' This form allows the user to select what data they wish to output and in
' what format. OpenArgs are required to specify what the current record was
' and whether it was worked
' This must take the format: record
' SAJ
'**********************************************************************
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    Dim rec
    rec = Me.OpenArgs
    
    Me![txtRec] = rec
    
    Me![cboStart].RowSource = "SELECT DISTINCT [GID] FROM [dbo_Groundstone: Basic_Data_Pre2013_WithReRecordingFlag] ORDER BY [GID];"
    Me![cboEnd].RowSource = "SELECT DISTINCT [GID] FROM [dbo_Groundstone: Basic_Data_Pre2013_WithReRecordingFlag] ORDER BY [GID];"
    
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
