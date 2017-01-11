Option Compare Database
Option Explicit

Private Sub cmdCompleteLFoot_Click()
'new season 2007 - fill out all hands bones
On Error GoTo err_completeLFoot

    'check all the hand bone fields
    
    Me!Metatarsal_1_left = True
    Me!Metatarsal_2_left = True
    Me!Metatarsal_3_left = True
    Me!Metatarsal_4_left = True
    Me!Metatarsal_5_left = True

    Me!Proximal_phalanx_1_left = True
    Me!Distal_phalanx_1_left = True
    
    Me![Proximal_phalanges_2-5_left] = 4
    Me![Middle_phalanges_2-5_left] = 4
    Me![Distal_phalanges_2-5_left] = 4
    
    
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

err_completeLFoot:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdCompleteRFoot_Click()
'new season 2007 - fill out all hands bones
On Error GoTo err_completeRFoot

    'check all the hand bone fields
    
    Me!Metatarsal_1_right = True
    Me!Metatarsal_2_right = True
    Me!Metatarsal_3_right = True
    Me!Metatarsal_4_right = True
    Me!Metatarsal_5_right = True

    Me!Proximal_phalanx_1_right = True
    Me!Distal_phalanx_1_right = True
    
    Me![Proximal_phalanges_2-5_right] = 4
    Me![Middle_phalanges_2-5_right] = 4
    Me![Distal_phalanges_2-5_right] = 4
    
    'Me!Proximal_phalanx_2_right = True
    'Me!Proximal_phalanx_3_right = True
    'Me!Proximal_phalanx_4_right = True
    'Me!Proximal_phalanx_5_right = True

    'Me!Middle_phalanx_2_right = True
    'Me!Middle_phalanx_3_right = True
    'Me!Middle_phalanx_4_right = True
    'Me!Middle_phalanx_5_right = True

    
    'Me!Distal_phalanx_2_right = True
    'Me!Distal_phalanx_3_right = True
    'Me!Distal_phalanx_4_right = True
    'Me!Distal_phalanx_5_right = True


Exit Sub

err_completeRFoot:
    Call General_Error_Trap
    Exit Sub
End Sub
