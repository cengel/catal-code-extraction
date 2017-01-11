Option Compare Database
Option Explicit

'module level variable to hold report source
Dim g_reportfilter

Private Sub Close_Click()
'***************************************************
' Standard close
'***************************************************
On Error GoTo err_close_Click

    DoCmd.Close acForm, Me.Name
    
    Exit Sub

err_close_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdBuildSQL_Click()
'This takes the criteria entered at the top of the screen and builds the sql string that
'will be used as the recordsource for the sub form that displays the results.

On Error GoTo err_buildsql
'remember to replace * with %
Dim selectsql, wheresql, orderbysql, fullsql

selectsql = "SELECT * FROM [Exca: Unit Sheet with Relationships] "

wheresql = ""

If Me![txtBuildingNumbers] <> "" Then
    'wheresql = "[Building] like '%" & Me![txtBuildingNumbers] & "%'"
    wheresql = wheresql & "(" & Me![txtBuildingNumbers] & ")"
End If

If Me![txtSpaceNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    'wheresql = wheresql & "[Space] like '%" & Me![txtSpaceNumbers] & "%'"
    wheresql = wheresql & "(" & Me![txtSpaceNumbers] & ")"
End If

If Me![txtFeatureNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    'wheresql = wheresql & "[Feature] like '%" & Me![txtFeatureNumbers] & "%'"
    wheresql = wheresql & "(" & Me![txtFeatureNumbers] & ")"
End If

If Me![txtLevels] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    'wheresql = wheresql & "[Levels] like '%" & Me![txtLevels] & "%'"
    wheresql = wheresql & "(" & Me![txtLevels] & ")"
End If

'new 2010
If Me![txtHodderLevel] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    'wheresql = wheresql & "[Levels] like '%" & Me![txtLevels] & "%'"
    wheresql = wheresql & "(" & Me![txtHodderLevel] & ")"
End If

If Me![txtCategory] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Category] like '%" & Me![txtCategory] & "%'"
End If

If Me![cboArea] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Area] = '" & Me![cboArea] & "'"
End If

If Me![cboYear] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Year] = " & Me![cboYear]
End If

If Me![txtUnitNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtUnitNumbers] & ")"
End If

If Me![txtText] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    'actually description is not in unit sheet with relationships
    'wheresql = wheresql & "([Description] like '*" & Me![txtText] & "*' OR [Discussion] like '*" & Me![txtText] & "*')"
    wheresql = wheresql & "([Discussion] like '%" & Me![txtText] & "%' OR [Exca: Unit Sheet with Relationships].[Description] like '%" & Me![txtText] & "%')"
End If

If Me![cboDataCategory] <> "" Then
    'change select statement when a data category has been entered
    selectsql = "SELECT [Exca: Unit Sheet with Relationships].[Unit Number], [Exca: Unit Sheet with Relationships].Year, " & _
                "[Exca: Unit Sheet with Relationships].Area, [Exca: Unit Sheet with Relationships].Category, " & _
                "[Exca: Unit Sheet with Relationships].[Grid X], [Exca: Unit Sheet with Relationships].[Grid Y], " & _
                "[Exca: Unit Sheet with Relationships].Description, [Exca: Unit Sheet with Relationships].Discussion, [Exca: Unit Sheet with Relationships].[Priority Unit], " & _
                "[Exca: Unit Sheet with Relationships].ExcavationStatus, [Exca: Unit Sheet with Relationships].HodderLevel, [Exca: Unit Sheet with Relationships].MellaartLevels," & _
                "[Exca: Unit Sheet with Relationships].Building, [Exca: Unit Sheet with Relationships].Space, [Exca: Unit Sheet with Relationships].Feature, " & _
                "[Exca: Unit Sheet with Relationships].TimePeriod, [Exca: Unit Data Categories].[Data Category]" & _
                " FROM [Exca: Unit Sheet with Relationships] INNER JOIN [Exca: Unit Data Categories] ON [Exca: Unit Sheet with Relationships].[Unit Number] = [Exca: Unit Data Categories].[Unit Number]"
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Data Category] = '" & Me![cboDataCategory] & "'"
End If

'end of where clause if it contains info add the where
If wheresql <> "" Then selectsql = selectsql & " WHERE "

orderbysql = " ORDER BY [Exca: Unit Sheet with Relationships].[Unit Number];"

'create full sql string
fullsql = selectsql & wheresql & orderbysql
'the where clause will be used as the filter if a report is produced
g_reportfilter = wheresql
Me!txtSQL = fullsql
'make the sql the recordsource for the subform of results
Me![frm_subSearch].Form.RecordSource = fullsql
If Me![frm_subSearch].Form.RecordsetClone.RecordCount = 0 Then
    'if no records returned then tell the user
    MsgBox "No records match the criteria you entered.", 48, "No Records Found"
    Me![cmdClearSQL].SetFocus
End If

Exit Sub

err_buildsql:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClear_Click()
On Error GoTo err_clear

Me![txtBuildingNumbers] = ""

Exit Sub
err_clear:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearFeature_Click()
On Error GoTo err_feature

Me![txtFeatureNumbers] = ""
Exit Sub
err_feature:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearHodder_Click()
On Error GoTo err_cmdClearHodder
Me![txtHodderLevel] = ""
Exit Sub
err_cmdClearHodder:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCLearLevel_Click()
On Error GoTo err_level
Me![txtLevels] = ""
Exit Sub
err_level:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearSpace_Click()
On Error GoTo err_space
Me![txtSpaceNumbers] = ""
Exit Sub
err_space:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearSQL_Click()
On Error GoTo err_clearsql
'clear all criteria and reset sql
Dim sql

Me![txtBuildingNumbers] = ""
Me![txtSpaceNumbers] = ""
Me![txtFeatureNumbers] = ""
Me![txtLevels] = ""
Me![txtHodderLevel] = ""
Me![txtCategory] = ""
Me![cboArea] = ""
Me![cboYear] = ""
Me![txtUnitNumbers] = ""
Me![txtText] = ""
Me![cboDataCategory] = ""
sql = "SELECT * FROM [Exca: Unit Sheet with Relationships] ORDER BY [Unit Number];"
Me!txtSQL = sql
Me![frm_subSearch].Form.RecordSource = sql
Exit Sub
err_clearsql:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearUnit_Click()
On Error GoTo err_unit
Me![txtUnitNumbers] = ""
Exit Sub
err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterBuilding_Click()
On Error GoTo err_building
Dim openarg
openarg = "Building"

If Me![txtBuildingNumbers] <> "" Then openarg = "Building;" & Me![txtBuildingNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_building:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterFeature_Click()
On Error GoTo err_enterfeature
Dim openarg
openarg = "Feature"

If Me![txtFeatureNumbers] <> "" Then openarg = "Feature;" & Me![txtFeatureNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterfeature:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterHodder_Click()
On Error GoTo err_enterHlevel
Dim openarg
openarg = "HodderLevel"

If Me![txtHodderLevel] <> "" Then openarg = "HodderLevel;" & Me![txtHodderLevel]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterHlevel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterLevel_Click()
On Error GoTo err_enterlevel
Dim openarg
openarg = "MellaartLevels"

If Me![txtLevels] <> "" Then openarg = "MellaartLevels;" & Me![txtLevels]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterlevel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterSpace_Click()
On Error GoTo err_enterspace
Dim openarg
openarg = "Space"

If Me![txtSpaceNumbers] <> "" Then openarg = "Space;" & Me![txtSpaceNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterspace:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err_cmdPrint
    Call cmdBuildSQL_Click
    
    If Me![frm_subSearch].Form.RecordsetClone.RecordCount = 0 Then
        'MsgBox "No records match the criteria you entered.", 48, "No Records Found"
        Me![cmdClearSQL].SetFocus
        Exit Sub
    Else
        DoCmd.OpenReport "R_unit_search_report", acViewPreview
        If Not IsNull(g_reportfilter) Then
            ''MsgBox g_reportfilter
            
            Reports![R_unit_search_report].FilterOn = True
            Reports![R_unit_search_report].Filter = g_reportfilter
        End If
    
    End If

    'g_reportsource = Me![frm_subSearch].Form.RecordSource
    'DoCmd.OpenReport "rpt_unit_search_report", acViewPreview
    'Reports![rpt_unit_search_report].RecordSource = Me![frm_subSearch].Form.RecordSource

Exit Sub

err_cmdPrint:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnit_Click()
On Error GoTo err_unitclick
Dim openarg
openarg = "unit number"

If Me![txtUnitNumbers] <> "" Then openarg = "unit number;" & Me![txtUnitNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_unitclick:
    Call General_Error_Trap
    Exit Sub
End Sub


