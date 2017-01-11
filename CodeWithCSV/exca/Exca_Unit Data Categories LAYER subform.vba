Option Compare Database
Option Explicit



Private Sub Data_Category_AfterUpdate()

'all values should be blank again if you change

    Me.In_Situ = ""
    Me.Location = ""
    Me.Description = ""
    Me.Material = ""
    Me.Deposition = ""
    Me.basal_spit = ""
    
End Sub

Private Sub Data_Category_Change()

Select Case Me.Data_Category
    Case "fill"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True
    'set values
    Me.Location.RowSource = " ; between walls; building; cut; feature"
    Me.Description.RowSource = "" 'depends on location
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous"
    Me.basal_spit.RowSource = " ; basal deposit"
           
    Case "floors (use)"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = True
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False
    'set values
    Me.Location.RowSource = " ;building; external; feature"
    Me.Description.RowSource = "" 'depends on location
    Me.Material.RowSource = " ;dark grey clay; mix (dark grey&white); occupation; white clay"
    Me.Deposition.RowSource = " ; composite (floors/bedding/plaster/packing/occupation); multiple; single"
    
    Case "construction/make-up/packing"
    'set fields
    Me.In_Situ.Enabled = True
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = True
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False
    'set values
    Me.Location.RowSource = " ; between walls; building; external; feature; floor (packing only); roof (building); wall/blocking"
    Me.Description.RowSource = "" 'depends on location
    Me.Material.RowSource = "" 'depends on location
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous; layered (wall plaster)"
    'Me.In_Situ.SetFocus
        
    Case "midden"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True

    'values
    Me.Location.RowSource = " ;external; in abandoned building"
    Me.Deposition.RowSource = " ; alluviated dumps; coarsely bedded (dumps); finely bedded"
    Me.basal_spit.RowSource = " ; basal deposit"
    
    Case "activity"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True

    'values
    Me.Description.RowSource = " ;fire spots (non-structured); lime burning; penning"
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous"
    Me.basal_spit.RowSource = " ; basal deposit"
    
    Case "natural"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False

    'values
    Me.Deposition.RowSource = " ; alluvium; backswamp; buried soil; colluvium; marl"
    
    Case "arbitrary"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = False
    Me.basal_spit.Enabled = False

    'values
    Me.Description.RowSource = " ; 60's; animal burrow; arbitrary allocation for samples; baulks; cleaning; not excavated; unstratified; very mixed; void (unused unit no)"
    
    Case Else
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.basal_spit.Enabled = False
    
End Select

End Sub


Private Sub Data_Category_Exit(Cancel As Integer)
Me.refresh

End Sub

Private Sub Description_AfterUpdate()
'all should be blank again

    Me.Material = ""
    Me.Deposition = ""
    Me.basal_spit = ""

Select Case Me.Data_Category
Case "floors (use)"
    If Me.Description = "oven" Or Me.Description = "hearth" Then
        Me.Material.RowSource = " ; baked; dark grey clay; mix (dark grey&white); occupation; white clay"
        Else
        Me.Material.RowSource = " ; dark grey clay; mix (dark grey&white); occupation; white clay"
    End If
End Select

End Sub


Private Sub Form_Current()
On Error GoTo err_curr

Select Case Me.Data_Category
'------------------------------------------------
    Case "fill"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True
    'set values
    Me.Location.RowSource = " ; between walls; building; cut; feature"
    'Me.Description.RowSource = "" 'depends on location
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
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous"
    Me.basal_spit.RowSource = " ; basal deposit"
    
    
    
'-------------------------------------------------
    Case "floors (use)"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = True
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False
    'set values
    Me.Location.RowSource = " ; building; external; feature"
    'Me.Description.RowSource = "" 'depends on location
        Select Case Me.Location
            Case "building"
            Me.Description.RowSource = " ; general; raised area (platform); roof (use)"
            Me.Description.Enabled = True
            Case "feature"
            Me.Description.RowSource = " ; basin; bin; burial; hearth; niche; oven; pedestal/podium/plinth; ridge"
            Me.Description.Enabled = True
            Case Else
            Me.Description.RowSource = ""
            Me.Description.Enabled = False
        End Select
    Me.Material.RowSource = " ; dark grey clay; mix (dark grey&white); occupation; white clay"
    Me.Deposition.RowSource = " ; composite (floors/bedding/plaster/packing/occupation); multiple; single"
        
    'set material for burnt  fetures.
    If Me.Description = "oven" Or Me.Description = "hearth" Then
        Me.Material.RowSource = " ; baked; dark grey clay; mix (dark grey&white); occupation; white clay"
        Else
        Me.Material.RowSource = " ; dark grey clay; mix (dark grey&white); occupation; white clay"
    End If
    
 '------------------------------------------------
    Case "construction/make-up/packing"
    'set fields
    Me.In_Situ.Enabled = True
    Me.Location.Enabled = True
    Me.Description.Enabled = True
    Me.Material.Enabled = True
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False
    Me.In_Situ.SetFocus
    
    'set values
    Me.Location.RowSource = " ; between walls; building; external; feature; floor (packing only); roof (building); wall/blocking"
    'Me.Description.RowSource = "" 'depends on location
        Select Case Me.Location
            Case "feature"
            Me.Description.RowSource = " ; basin; bin; hearth; moulding; niche; oven; pedestal/podium/plinth; post; raised area (platform); ridge "
            Me.Description.Enabled = True
            Case Else
            Me.Description.RowSource = ""
            Me.Description.Enabled = False
        End Select
    Me.Material.RowSource = " ; brick; brick&mortar; mortar; pise-like; plaster; re-used brick&mortar; re-used superstructure"
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous; layered (wall plaster)"
    
   
'------------------------------------------------
    Case "midden"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = True
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True

    'values
    Me.Location.RowSource = " ;external; in abandoned building"
    Me.Deposition.RowSource = " ; alluviated dumps; coarsely bedded (dumps); finely bedded"
    Me.basal_spit.RowSource = " ; basal deposit"
'------------------------------------------------
    Case "activity"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = True

    'values
    Me.Description.RowSource = " ;fire spots (non-structured); lime burning; penning"
    Me.Deposition.RowSource = " ; heterogeneous; homogeneous"
    Me.basal_spit.RowSource = " ; basal deposit"
'------------------------------------------------
    Case "natural"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.Deposition.Enabled = True
    Me.basal_spit.Enabled = False

    'values
    Me.Deposition.RowSource = " ; alluvium; backswamp; buried soil; colluvium; marl"
'------------------------------------------------
    Case "arbitrary"
    'set fields
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = True
    Me.Material.Enabled = False
    Me.Deposition.Enabled = False
    Me.basal_spit.Enabled = False

    'values
    Me.Description.RowSource = " ; 60's; animal burrow; arbitrary allocation for samples; baulks; cleaning; not excavated; unstratified; very mixed; void (unused unit no);"
'------------------------------------------------
    Case Else
    Me.In_Situ.Enabled = False
    Me.Location.Enabled = False
    Me.Description.Enabled = False
    Me.Material.Enabled = False
    Me.basal_spit.Enabled = False
    
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

Private Sub Location_AfterUpdate()

'all values blank again
    Me.Description = ""
    Me.Material = ""
    Me.Deposition = ""
    Me.basal_spit = ""
    
Select Case Me.Data_Category

    Case "fill"
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
    
    Case "floors (use)"
    Select Case Me.Location
        Case "building"
        Me.Description.RowSource = " ; general; raised area (platform); roof (use)"
        Me.Description.Enabled = True
        Case "feature"
        Me.Description.RowSource = " ; basin; bin; burial; hearth; niche; oven; pedestal/podium/plinth; ridge"
        Me.Description.Enabled = True
        Case Else
        Me.Description.RowSource = ""
        Me.Description.Enabled = False
    End Select
    
    Case "construction/make-up/packing"
    Select Case Me.Location
        Case "feature"
        Me.Description.RowSource = " ; basin; bin; hearth; moulding; niche; oven; pedestal/podium/plinth; post; raised area (platform); ridge"
        Me.Description.Enabled = True
        Case Else
        Me.Description.RowSource = ""
        Me.Description.Enabled = False
    End Select
    
    Me.Material.RowSource = " ; brick; brick&mortar; mortar; pise-like; plaster; re-used brick&mortar; re-used superstructure"

End Select

End Sub



Private Sub Material_AfterUpdate()
'all should be blank again
Me.Deposition = ""
Me.basal_spit = ""
Me.basal_spit.Enabled = False

'basal spit field (additional info) varies for floors

Select Case Me.Data_Category
Case "construction/make-up/packing"
    If Me.Material = "plaster" Then
    Me.basal_spit.Enabled = True
    Me.basal_spit.RowSource = " ; painted; unpainted"
    End If
    
Case "floors (use)"
    If Me.Material = "dark grey clay" Or Me.Material = "mix (dark grey&white)" Or Me.Material = "white clay" Then
    Me.basal_spit.Enabled = True
    Me.basal_spit.RowSource = " ; painted; unpainted"
    End If
End Select
    
End Sub

