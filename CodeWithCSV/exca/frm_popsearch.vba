Option Compare Database
Option Explicit
Dim toShow


Private Sub cboSelect_AfterUpdate()
On Error GoTo err_cboSelect

If toShow = "unit number" Then
    If Me![txtToFind] <> "" And Not IsNull(Me![txtToFind]) Then
        Me![txtToFind] = Me![txtToFind] & " OR "
    End If
    
    Me![txtToFind] = Me![txtToFind] & "[Exca: Unit sheet with Relationships].[Unit Number] = " & Me!cboSelect
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
            Forms![frm_search]![txtBuildingNumbers] = Me![txtToFind]
        Case "space"
            Forms![frm_search]![txtSpaceNumbers] = Me![txtToFind]
        Case "feature"
            Forms![frm_search]![txtFeatureNumbers] = Me![txtToFind]
        Case "MellaartLevels"
            Forms![frm_search]![txtLevels] = Me![txtToFind]
        Case "HodderLevel"
            Forms![frm_search]![txtHodderLevel] = Me![txtToFind]
        Case "unit number"
            Forms![frm_search]![txtUnitNumbers] = Me![txtToFind]
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
            existing = right(toShow, Len(toShow) - colonpos)
            'MsgBox existing
            toShow = Left(toShow, colonpos - 1)
        End If
        
        Select Case toShow
        Case "building"
            Me![lblTitle].Caption = "Select Building Number"
            Me![cboSelect].RowSource = "Select [Number] from [Exca: Building Details];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "space"
            Me![lblTitle].Caption = "Select Space Number"
            Me![cboSelect].RowSource = "Select [Space Number] from [Exca: Space Sheet];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "feature"
            Me![lblTitle].Caption = "Select Feature Number"
            Me![cboSelect].RowSource = "Select [Feature Number] from [Exca: Features];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "Mellaartlevels"
            Me![lblTitle].Caption = "Select Mellaart Level"
            Me![cboSelect].RowSource = "Select [Level] from [Exca:LevelLOV];"
            If existing <> "" Then Me![txtToFind] = existing
        Case "Hodderlevel"
            Me![lblTitle].Caption = "Select Hodder Level"
            Me![cboSelect].RowSource = "SELECT DISTINCT [Exca: Space Sheet].HodderLevel FROM [Exca: Space Sheet] WHERE ((([Exca: Space Sheet].HodderLevel) <> '')) ORDER BY [Exca: Space Sheet].HodderLevel;"
            If existing <> "" Then Me![txtToFind] = existing
        Case "unit number"
            Me![lblTitle].Caption = "Select Unit Number"
            Me![cboSelect].RowSource = "Select [unit number] from [Exca: Unit Sheet] ORDER BY [unit number];"
            If existing <> "" Then Me![txtToFind] = existing
        End Select
        
        Me.refresh
    End If



Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
