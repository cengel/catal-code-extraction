Option Compare Database
Option Explicit



Private Sub cmdCompleteLHand_Click()
'new season 2007 - fill out all hands bones
On Error GoTo err_completeLHand

    'check all the hand bone fields
    
    Me!Metacarpal_1_left = True
    Me!Metacarpal_2_left = True
    Me!Metacarpal_3_left = True
    Me!Metacarpal_4_left = True
    Me!Metacarpal_5_left = True

    Me!Proximal_phalanx_1_left = True
    Me!Distal_phalanx_1_left = True
    
    Me![Proximal_phalanges_2-5_left] = 4
    Me![Middle_phalanges_2-5_left] = 4
    Me![Distal_phalanges_2-5_left] = 4
    
    '16/06/2008 SAJ SH
    
    'Me!Proximal_phalanx_2_left = True
    'Me!Proximal_phalanx_3_left = True
    'Me!Proximal_phalanx_4_left = True
    'Me!Proximal_phalanx_5_left = True

    'Me!Middle_phalanx_2_left = True
    'Me!Middle_phalanx_3_left = True
    'Me!Middle_phalanx_4_left = True
    'Me!Middle_phalanx_5_left = True

    
    'Me!Distal_phalanx_2_left = True
    'Me!Distal_phalanx_3_left = True
    'Me!Distal_phalanx_4_left = True
    'Me!Distal_phalanx_5_left = True


Exit Sub

err_completeLHand:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdCompleteRHand_Click()
'new season 2007 - fill out all hands bones
On Error GoTo err_completeRHand

    'check all the hand bone fields
    
    Me!Metacarpal_1_right = True
    Me!Metacarpal_2_right = True
    Me!Metacarpal_3_right = True
    Me!Metacarpal_4_right = True
    Me!Metacarpal_5_right = True

    Me!Proximal_phalanx_1_right = True
    Me!Distal_phalanx_1_right = True
    
    Me![Proximal_phalanges_2-5_right] = 4
    Me![Middle_phalanges_2-5_right] = 4
    Me![Distal_phalanges_2-5_right] = 4
    
    '16/06/2008 SAJ SH
    'Me!Proximal_phalanx_2_right = True
    'Me!Proximal_phalanx_3_right = True
    'Me!Proximal_phalanx_4_right = True
    'Me!Proximal_phalanx_5_right = True

    
    'Me!Middle_phalanx_2_right = True
    'Me!Middle_phalanx_3_right = True
    'Me!Middle_phalanx_4_right = True
    'Me!Middle_phalanx_5_right = True

    'Me!Distal_phalanx_1_right = True
    'Me!Distal_phalanx_2_right = True
    'Me!Distal_phalanx_3_right = True
    'Me!Distal_phalanx_4_right = True
    'Me!Distal_phalanx_5_right = True


Exit Sub

err_completeRHand:
    Call General_Error_Trap
    Exit Sub
End Sub
