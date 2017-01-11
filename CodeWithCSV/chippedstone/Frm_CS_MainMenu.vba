Option Compare Database
Option Explicit



Private Sub cmdAdmin_Click()
'open the pub data screen
On Error GoTo err_Admin

    DoCmd.OpenForm "Frm_GS_AdminMenu", acNormal

Exit Sub

err_Admin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOpenData_Click()
'open the main data screen
On Error GoTo err_OpenData

    DoCmd.OpenForm "Frm_CS_BasicData", acNormal

Exit Sub

err_OpenData:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cmdQuit_Click()
'quit system
On Error GoTo err_Quit

    DoCmd.Quit acQuitSaveAll
    

Exit Sub

err_Quit:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdStageTwo_Click()
'open the sample data screen
On Error GoTo err_Sample

    DoCmd.OpenForm "Frm_CS_StageTwo", acNormal

Exit Sub

err_Sample:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub stageone_Click()
'open the main data screen
On Error GoTo err_stageone_Click

    DoCmd.OpenForm "Frm_CS_StageOne2016", acNormal

Exit Sub

err_stageone_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPub_Click()
'open the pub data screen
On Error GoTo err_pub

    DoCmd.OpenForm "LithicForm:BagAndUnitDescription", acNormal
Exit Sub

err_pub:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub stagetwo_Click()
'open the main data screen
On Error GoTo err_stageone_Click

    DoCmd.OpenForm "Frm_CS_StageTwo2016", acNormal

Exit Sub

err_stageone_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
