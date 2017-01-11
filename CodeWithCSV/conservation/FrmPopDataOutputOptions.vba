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
            where = "[FullConservation_Ref] = '" & Me![txtRec] & "'"
        ElseIf Me![optgrpData] = 2 Then
            'range
            If Me![cboStart] = "" Or Me![cboEnd] = "" Then
                MsgBox "Please enter both a Start and End value for the range of records to output.", vbInformation, "Insufficient Data"
                Exit Sub
            Else
                where = "[FullConservation_Ref] BETWEEN '" & Me![cboStart] & "' AND '" & Me![cboEnd] & "'"
            End If
        ElseIf Me![optgrpData] = 3 Then
            'all
            where = ""
        ElseIf Me![optgrpData] = 4 Then
            'current only
            where = "[NameID] = " & Me![cboName]
        End If
   
        If Me![optgrpData] = 4 Then
            sql = "SELECT * FROM [Q_BasicDataWithConservatorName]"
        Else
            sql = "SELECT * FROM [" & Me![txtData] & "]"
        End If
        If where <> "" Then sql = sql & " where " & where
        sql = sql & ";"
   
   
        If which = "report" Then
            If Me![optgrpData] = 4 Then
                rpt = "Conserv: Full Printout for Conservator"
                DoCmd.OpenReport rpt, acViewPreview, , where
                'DoCmd.SelectObject acReport, "R_BasicData", True
            Else
                 rpt = "Conserv: Full Printout"
                DoCmd.OpenReport rpt, acViewPreview, , where
            End If
            MsgBox "If you can't see the report it has appeared behind the form, go to the Window menu and select it from there", vbInformation, "Report Location"
            DoCmd.Close acForm, Me.Name
            'rpt.SetFocus
        ElseIf which = "excel" Then
            Dim mydb As Database, myq As QueryDef
            Set mydb = CurrentDb
            Set myq = mydb.CreateQueryDef("ConservationExcelExport")
            
            myq.sql = sql
            myq.ReturnsRecords = False
            
            DoCmd.OutputTo acOutputQuery, "ConservationExcelExport", acFormatXLS, , True
            
            mydb.QueryDefs.Delete ("ConservationExcelExport")
            
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
    
    Me![cboStart].RowSource = "SELECT DISTINCT [FullConservation_Ref] FROM [" & data & "] ORDER BY [FullConservation_Ref];"
    Me![cboEnd].RowSource = "SELECT DISTINCT [FullConservation_Ref] FROM [" & data & "] ORDER BY [FullConservation_Ref];"

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
    Me![cboName].Enabled = False
ElseIf Me![optgrpData] = 4 Then
    Me![cboName].Enabled = True
    Me![cboStart].Enabled = False
    Me![cboEnd].Enabled = False
Else
    Me![cboName].Enabled = False
    Me![cboStart].Enabled = False
    Me![cboEnd].Enabled = False
End If

Exit Sub

err_optgrpData:
    Call General_Error_Trap
    Exit Sub
End Sub
