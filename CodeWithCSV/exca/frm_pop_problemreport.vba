Option Compare Database
Option Explicit
Dim toShow, entitynum






Private Sub cmdCancel_Click()
On Error GoTo err_cmdCancel
    DoCmd.Close acForm, "frm_pop_problemreport"
Exit Sub

err_cmdCancel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClear_Click()
On Error GoTo err_cmdClear
    Me![txtToFind] = ""
    Me![cboSelect] = ""
Exit Sub

err_cmdClear:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOK_Click()
On Error GoTo err_cmdOK
    
    If (Me![Comment] = "" Or Me![ReportersName] = "") Or (IsNull(Me![Comment]) Or IsNull(Me![ReportersName])) Then
        MsgBox "Please enter both your comment and your name, otherwise cancel the report", vbInformation, "Insufficient Info"
        Exit Sub
    Else
        ''2010 allow a comment by anyone even catalhoyuk read only login - as just as valid and don't want to miss them
        'insert this info into the table
        Dim sql, strcomment
        strcomment = Replace(Me![Comment], "'", "''") 'bug fix july 2009 on site
        ''sql = "INSERT INTO [Exca: Report_Problem] ([EntityNumber], [EntityType], [Comment], [ReportersName], [ReportedOn]) VALUES (" & entitynum & ", '" & toShow & "', '" & strcomment & "', '" & Me![ReportersName] & "', Format(Date(), 'dd/mm/yyyy'));"
        '''MsgBox sql
        ''DoCmd.RunSQL sql
        
        If spString <> "" Then
            Dim mydb As DAO.Database
            Dim myq1 As QueryDef
    
            Set mydb = CurrentDb
            Set myq1 = mydb.CreateQueryDef("")
            myq1.Connect = spString
    
            myq1.ReturnsRecords = False
            'myq1.sql = "sp_Excavation_Add_Problem_Report_Entry " & entitynum & ", '" & toShow & "','" & strcomment & "','" & Me![ReportersName] & "','" & Format(Date, "mm/dd/yyyy") & "'"
            '24/07/2011 - the above line was failing on this date as it was reading date as "24/07/2011" and when this was run to stored proc it would fail
            'changing to long date solves this even though 24/07/2011 is exactly what is written into the database field. Another US/UK date format issue that only appears in the latter
            'part of month.
            myq1.sql = "sp_Excavation_Add_Problem_Report_Entry " & entitynum & ", '" & toShow & "','" & strcomment & "','" & Me![ReportersName] & "','" & Format(Date, "Long Date") & "'"
            ''dbo.sp_Excavation_Add_Problem_Report_Entry (@entitynum int, @toShow nvarchar(50), @comment nvarchar(1000), @by nvarchar(100), @when datetime) AS
            myq1.Execute
            
            myq1.Close
            Set myq1 = Nothing
            mydb.Close
            Set mydb = Nothing
        
            MsgBox "Thank you, your report has been saved for the Administrator to check", vbInformation, "Done"
        Else
            'it failed and this should be handled better than this but I have no time, please fix this - SAJ sept 08
            MsgBox "Sorry but this comment cannot be inserted at this time, please restart the database and try again", vbCritical, "Error"
        End If

        DoCmd.Close acForm, "frm_pop_problemreport"
    End If
Exit Sub

err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'when the form opens it should bring in the entity info: eg whether its a unit, feature, building etc
'also it should bring in the number. However in case this is needed where the user needs to enter this info
'its left flexible
'inputs: entitytype;entitynumber eg: unit;1056
On Error GoTo err_open

Dim colonpos
     If Not IsNull(Me.OpenArgs) Then
        Me![cboSelect].Visible = False
        
        toShow = LCase(Me.OpenArgs)
        colonpos = InStr(toShow, ";")
        
        If colonpos > 0 Then
            'if there is a ; then this means there is some existing criteria to extract from openargs
            entitynum = right(toShow, Len(toShow) - colonpos)
            'MsgBox existing
            toShow = Left(toShow, colonpos - 1)
        End If
        
        Select Case toShow
        Case "building"
            Me![lblTitle].Caption = "Report a Building Record Problem"
            Me![cboSelect].RowSource = "Select [Number] from [Exca: Building Details];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Building Number: " & entitynum
        Case "space"
            Me![lblTitle].Caption = "Report a Space Record Problem"
            Me![cboSelect].RowSource = "Select [Space Number] from [Exca: Space Sheet];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Space Number: " & entitynum
        Case "feature number"
            Me![lblTitle].Caption = "Report a Feature Record Problem"
            Me![cboSelect].RowSource = "Select [Feature Number] from [Exca: Features];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Feature Number: " & entitynum
        Case "unit number"
            Me![lblTitle].Caption = "Report a Unit Record Problem"
            Me![cboSelect].RowSource = "Select [unit number] from [Exca: Unit Sheet] ORDER BY [unit number];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Unit Number: " & entitynum
        End Select
        
        Me.refresh
   
Else
    Me![lblTitle].Visible = False
    Me![lblEntity].Visible = False
    Me![cboSelect].Visible = True
End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
