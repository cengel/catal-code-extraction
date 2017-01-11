Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************







Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind
    
    If Me![cboFind] <> "" Then
        Me.RecordSource = "SELECT *  FROM [Exca: Report_Problem] WHERE [ReportedOn] = #" & Format(Me![cboFind], "dd-mmm-yyyy") & "#;"
        Me![tglShowAll] = False
        Me![tglShowAll].Caption = "Show All Records"
    End If
Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub








Private Sub cmdPrint_Click()
'may 2009 - add a print report function
On Error GoTo err_cmdPrint

    If Me![tglShowAll] = True Then
        DoCmd.OpenReport "R_Problem_Reports", acViewPreview
    Else
        DoCmd.OpenReport "R_Problem_Reports", acViewPreview, , "[Resolved] = false"
    End If

Exit Sub

err_cmdPrint:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
    
    Me![tglShowAll] = False
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub tglShowAll_Click()
'filter list depending on button state
On Error GoTo err_tglShowAll

    If Me![tglShowAll] = True Then
        Me.RecordSource = "SELECT *  FROM [Exca: Report_Problem];"
        Me![tglShowAll].Caption = "Unresolved Only"
        Me![cboFind] = ""
    Else
        Me.RecordSource = "SELECT *  FROM [Exca: Report_Problem] WHERE ((([Exca: Report_Problem].Resolved)=False))"
        Me![tglShowAll].Caption = "Show All Reports"
        Me![cboFind] = ""
    End If
Exit Sub

err_tglShowAll:
    Call General_Error_Trap
    Exit Sub
End Sub
