Option Compare Database
Option Explicit



Private Sub Form_Current()

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
