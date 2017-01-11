Option Compare Database
Option Explicit

Private Sub Bulb_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_BulbNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Bulb].Undo
Exit Sub

err_BulbNot:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub BulbScar_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_BulbScarNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![BulbScar].Undo
Exit Sub

err_BulbScarNot:
    Call General_Error_Trap
    Exit Sub

End Sub



Private Sub cboOverhang_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_OverhangNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![OverHang].Undo
Exit Sub

err_OverhangNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboDistal_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cboDistalNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboDistal].Undo
Exit Sub

err_cboDistalNot:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Edges_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_EdgesNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Edges].Undo
Exit Sub

err_EdgesNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Profile_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_ProfileNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Profile].Undo
Exit Sub

err_ProfileNot:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Ridges_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_RidgesNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Ridges].Undo
Exit Sub

err_RidgesNot:
    Call General_Error_Trap
    Exit Sub
End Sub
