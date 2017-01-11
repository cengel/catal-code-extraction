Option Compare Database
Option Explicit

Private Sub cmdAdmin_Click()
'open the pub data screen
On Error GoTo err_Admin

    DoCmd.OpenForm "Frm_Photo_AdminMenu", acNormal
    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_Admin:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOpenData_Click()
'open the main data screen
On Error GoTo err_OpenData

    Dim stDocName
    stDocName = "photo_sheet"
    DoCmd.OpenForm stDocName, acNormal
    DoCmd.Close acForm, Me.Name
    
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




