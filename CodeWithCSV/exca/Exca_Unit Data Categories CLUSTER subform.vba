Option Compare Database
Option Explicit



Private Sub Form_Current()
On Error GoTo err_curr
'data category is cluster
'location rowsource is defined
'location may already be present

Select Case Me.Location
            Case "cut"
            Me.Description.RowSource = " ; burial; ditch; foundation cut; gully; pit; posthole; scoop; stakehole"
            Me.Description.Enabled = True
            Case "feature"
            Me.Description.RowSource = " ; basin; bin; hearth; niche; oven"
            Me.Description.Enabled = True
            Case Else
            Me.Description.RowSource = ""
            Me.Description.Enabled = False
End Select
Exit Sub

err_curr:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' Set up form view depending on permissions
' SAJ v9.1
'**********************************************************************
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ToggleFormReadOnly Me, False
    Else
        'set read only form here, just once
        ToggleFormReadOnly Me, True
        'see subform Skeleton Sheet on open for reason for this line
        If Me.AllowAdditions = False Then Me.AllowAdditions = True
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Location_Change()
'description blank again, others stay
    Me.Description = ""
    
    Select Case Me.Location
        Case "cut"
        Me.Description.RowSource = " ; burial; ditch; foundation cut; gully; pit; posthole; scoop; stakehole"
        Me.Description.Enabled = True
        Case "feature"
        Me.Description.RowSource = " ; basin; bin; hearth; niche; oven"
        Me.Description.Enabled = True
        Case Else
        Me.Description.RowSource = ""
        Me.Description.Enabled = False
    End Select
End Sub
