Option Compare Database
Option Explicit

'module level variable to hold report source
Dim g_reportfilter

Private Sub Close_Click()
'***************************************************
' Standard close
'***************************************************
On Error GoTo err_close_Click
     DoCmd.OpenForm "Frm_Menu", , , , acFormPropertySettings
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

selectsql = "SELECT * FROM [view_Fig_MainData_Collated] "

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
    ''wheresql = wheresql & "([Discussion] like '%" & Me![txtText] & "%' OR [Exca: Unit Sheet with Relationships].[Description] like '%" & Me![txtText] & "%')"
    wheresql = wheresql & "([Description] like '%" & Me![txtText] & "%')"
End If

If Me![cboDataCategory] <> "" Then
    'change select statement when a data category has been entered
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Data Category] = '" & Me![cboDataCategory] & "'"
End If

If Me![txtFigurineID] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtFigurineID] & ")"
End If

If Me![txtMellID] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtMellID] & ")"
End If

If Me![txtObjectType] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtObjectType] & ")"
End If
       
If Me![txtObjectForm] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtObjectForm] & ")"
End If
        
If Me![txtFormType] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtFormType] & ")"
End If
        
If Me![txtForm] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtForm] & ")"
End If
        
If Me![txtQuadruped] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtQuadruped] & ")"
End If

'end of where clause if it contains info add the where
If wheresql <> "" Then selectsql = selectsql & " WHERE "

orderbysql = " ORDER BY [view_Fig_MainData_Collated].[UnitNumber];"

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

Private Sub cmdClearFigID_Click()
On Error GoTo err_figid
Me![txtFigurineID] = ""
Exit Sub
err_figid:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearForm_Click()
On Error GoTo err_cmdClearForm
Me![txtForm] = ""
Exit Sub
err_cmdClearForm:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearFormType_Click()
On Error GoTo err_cmdClearFormType
Me![txtFormType] = ""
Exit Sub
err_cmdClearFormType:
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

Private Sub cmdClearMellID_Click()
On Error GoTo err_mellid
Me![txtMellID] = ""
Exit Sub
err_mellid:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearObjectForm_Click()
On Error GoTo err_ObjectForm
Me![txtObjectForm] = ""
Exit Sub
err_ObjectForm:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearObjectType_Click()
On Error GoTo err_ObjectType
Me![txtObjectType] = ""
Exit Sub
err_ObjectType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClearQuadruped_Click()
On Error GoTo err_Quadruped
Me![txtQuadruped] = ""
Exit Sub
err_Quadruped:
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
Me![txtFigurineID] = ""
Me![txtMellID] = ""
Me![txtObjectType] = ""
Me![txtObjectForm] = ""
Me![txtForm] = ""
Me![txtFormType] = ""
Me![txtQuadruped] = ""
        

sql = "SELECT * FROM [view_Fig_MainData_Collated] ORDER BY [UnitNumber];"
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
openarg = "Features"

If Me![txtFeatureNumbers] <> "" Then openarg = "Features;" & Me![txtFeatureNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterfeature:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterFigID_Click()
On Error GoTo err_figclick
Dim openarg
openarg = "ID number"

If Me![txtFigurineID] <> "" Then openarg = "id number;" & Me![txtFigurineID]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_figclick:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterForm_Click()
On Error GoTo err_f
Dim openarg
openarg = "FigForms"

If Me![txtForm] <> "" Then openarg = "FigForms;" & Me![txtForm]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_f:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterFormType_Click()
On Error GoTo err_ft
Dim openarg
openarg = "FormTypes"

If Me![txtFormType] <> "" Then openarg = "FormTypes;" & Me![txtFormType]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_ft:
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
openarg = "MellaartLevelCertain"

If Me![txtLevels] <> "" Then openarg = "MellaartLevelCertain;" & Me![txtLevels]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterlevel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterMellID_Click()
On Error GoTo err_figclick
Dim openarg
openarg = "MellaartID"

If Me![txtMellID] <> "" Then openarg = "MellaartID;" & Me![txtMellID]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_figclick:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdEnterObjectForm_Click()
On Error GoTo err_of
Dim openarg
openarg = "Object Form"

If Me![txtObjectForm] <> "" Then openarg = "Object Form;" & Me![txtObjectForm]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_of:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterObjectType_Click()
On Error GoTo err_ot
Dim openarg
openarg = "ObjectTypes"

If Me![txtObjectType] <> "" Then openarg = "ObjectTypes;" & Me![txtObjectType]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_ot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterQuadruped_Click()
On Error GoTo err_q
Dim openarg
openarg = "Quadruped"

If Me![txtQuadruped] <> "" Then openarg = "Quadruped;" & Me![txtQuadruped]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_q:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdEnterSpace_Click()
On Error GoTo err_enterspace
Dim openarg
openarg = "Spaces"

If Me![txtSpaceNumbers] <> "" Then openarg = "Spaces;" & Me![txtSpaceNumbers]
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
        DoCmd.OpenReport "R_search_report", acViewPreview
        If Not IsNull(g_reportfilter) Then
            ''MsgBox g_reportfilter
            
            Reports![R_search_report].FilterOn = True
            Reports![R_search_report].Filter = g_reportfilter
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


Private Sub Form_Open(Cancel As Integer)
''MsgBox IsNull(Forms![frm_search]![txtUnitNumbers])
End Sub

