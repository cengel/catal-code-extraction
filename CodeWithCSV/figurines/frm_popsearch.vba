Option Compare Database
Option Explicit
Dim toShow


Private Sub cboSelect_AfterUpdate()
On Error GoTo err_cboSelect

If toShow = "unit number" Then
    If Me![txtToFind] <> "" And Not IsNull(Me![txtToFind]) Then
        Me![txtToFind] = Me![txtToFind] & " OR "
    End If

    'Me![txtToFind] = Me![txtToFind] & "[Exca: Unit sheet with Relationships].[Unit Number] = " & Me!cboSelect
    Me![txtToFind] = Me![txtToFind] & "[unitnumber] = " & Me!cboSelect
ElseIf toShow = "MellaartID" Or toShow = "ID Number" Or toShow = "Object Form" Then
    If Me![txtToFind] <> "" And Not IsNull(Me![txtToFind]) Then
        Me![txtToFind] = Me![txtToFind] & " OR "
    End If

    'Me![txtToFind] = Me![txtToFind] & "[Exca: Unit sheet with Relationships].[Unit Number] = " & Me!cboSelect
    Me![txtToFind] = Me![txtToFind] & "[" & toShow & "] = '" & Me!cboSelect & "'"

Else
    'If Me!cboSelect <> "" Then
    '    If Me![txtToFind] = "" Or IsNull(Me![txtToFind]) Then Me![txtToFind] = ","
    '    Me![txtToFind] = Me![txtToFind] & Me!cboSelect & ","
    'End If
    If Me!cboSelect <> "" Then
        If Me![txtToFind] <> "" And Not IsNull(Me![txtToFind]) Then Me![txtToFind] = Me![txtToFind] & " OR"
        Me![txtToFind] = Me![txtToFind] & "[" & toShow & "] LIKE '%," & Me!cboSelect & ",%'"
    End If
End If
Me![cboSelect] = ""
DoCmd.GoToControl "cmdOK"
Exit Sub

err_cboSelect:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdCancel_Click()
On Error GoTo err_cmdCancel
    DoCmd.Close acForm, "frm_popsearch"
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
    Select Case toShow 'toshow is a module level variable that is set in on Open depending on the openargs
        Case "building"
            Forms![Frm_Search]![txtBuildingNumbers] = Me![txtToFind]
        Case "spaces"
            Forms![Frm_Search]![txtSpaceNumbers] = Me![txtToFind]
        Case "features"
            Forms![Frm_Search]![txtFeatureNumbers] = Me![txtToFind]
        Case "MellaartLevelCertain"
            Forms![Frm_Search]![txtLevels] = Me![txtToFind]
        Case "HodderLevel"
            Forms![Frm_Search]![txtHodderLevel] = Me![txtToFind]
        Case "unit number"
            Forms![Frm_Search]![txtUnitNumbers] = Me![txtToFind]
         Case "ID number"
            Forms![Frm_Search]![txtFigurineID] = Me![txtToFind]
        Case "MellaartID"
            Forms![Frm_Search]![txtMellID] = Me![txtToFind]
        Case "ObjectTypes"
             Forms![Frm_Search]![txtObjectType] = Me![txtToFind]
        Case "Object Form"
             Forms![Frm_Search]![txtObjectForm] = Me![txtToFind]
        Case "FigForms"
             Forms![Frm_Search]![txtForm] = Me![txtToFind]
        Case "FormTypes"
             Forms![Frm_Search]![txtFormType] = Me![txtToFind]
        Case "Quadruped"
             Forms![Frm_Search]![txtQuadruped] = Me![txtToFind]
        End Select

DoCmd.Close acForm, "frm_popsearch"
Exit Sub

err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
Dim existing, colonpos
    If Not IsNull(Me.OpenArgs) Then
        'MsgBox Me.OpenArgs
        toShow = LCase(Me.OpenArgs)
        colonpos = InStr(toShow, ";")
        
        If colonpos > 0 Then
            'if there is a ; then this means there is some existing criteria to extract from openargs
            existing = Right(toShow, Len(toShow) - colonpos)
            'MsgBox existing
            toShow = Left(toShow, colonpos - 1)
        End If
        
        Select Case toShow
        Case "building"
            Me![lblTitle].Caption = "Select Building Number"
            Me![cboSelect].RowSource = "Select [Number] from [Exca: Building Details];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "spaces"
            Me![lblTitle].Caption = "Select Space Number"
            Me![cboSelect].RowSource = "Select [Space Number] from [Exca: Space Sheet];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "features"
            Me![lblTitle].Caption = "Select Feature Number"
            Me![cboSelect].RowSource = "Select [Feature Number] from [Exca: Features];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "MellaartlevelCertain"
            Me![lblTitle].Caption = "Select Mellaart Level"
            Me![cboSelect].RowSource = "Select [Level] from [Exca:LevelLOV];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "Hodderlevel"
            Me![lblTitle].Caption = "Select Hodder Level"
            Me![cboSelect].RowSource = "SELECT DISTINCT [Exca: Space Sheet].HodderLevel FROM [Exca: Space Sheet] WHERE ((([Exca: Space Sheet].HodderLevel) <> '')) ORDER BY [Exca: Space Sheet].HodderLevel;"
            If existing <> "" Then Me![txtToFind] = existing
        Case "unit number"
            Me![lblTitle].Caption = "Select Unit Number"
            Me![cboSelect].RowSource = "Select DISTINCT [unitnumber] from [fig_maindata] WHERE [unitnumber] <> null ORDER BY [unitnumber];"
            If existing <> "" Then Me![txtToFind] = existing
         Case "ID number"
            Me![lblTitle].Caption = "Select Figurine ID"
            Me![cboSelect].RowSource = "Select DISTINCT [id number] from [fig_maindata] WHERE [id number] <> null ORDER BY [id number];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "MellaartID"
            Me![lblTitle].Caption = "Select Mellaart ID"
            Me![cboSelect].RowSource = "Select DISTINCT [Mellaartid] from [fig_maindata] WHERE [Mellaartid] <> null ORDER BY [mellaartid];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "ObjectTypes"
            Me![lblTitle].Caption = "Select Object Type"
            Me![cboSelect].RowSource = "Select DISTINCT [ObjectType] from [fig_objecttypes] ORDER BY [ObjectType];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "Object Form"
            Me![lblTitle].Caption = "Select Object Form"
            Me![cboSelect].RowSource = "Select DISTINCT [Object Form] from [fig_maindata] ORDER BY [Object form];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "FigForms"
            Me![lblTitle].Caption = "Select Form"
            Me![cboSelect].RowSource = "Select DISTINCT [Form] from [fig_forms] ORDER BY [Form];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "FormTypes"
            Me![lblTitle].Caption = "Select Form Type"
            Me![cboSelect].RowSource = "Select DISTINCT [FormType] from [fig_formtypes] ORDER BY [FormType];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "Quadruped"
            Me![lblTitle].Caption = "Select Quadruped"
            Me![cboSelect].RowSource = "Select DISTINCT [Quadruped] from [fig_quadruped] ORDER BY [Quadruped];"
            If existing <> "" Then Me![txtToFind] = existing
        End Select
        
        Me.Refresh
    End If



Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub


