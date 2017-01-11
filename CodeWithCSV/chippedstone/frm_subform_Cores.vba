Option Compare Database
Option Explicit
'code new 2010

Private Sub REjuv_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_RejuvNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![REjuv].Undo
Exit Sub

err_RejuvNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub SecondaryUse_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_SecUseNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![SecondaryUse].Undo
Exit Sub

err_SecUseNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Stage_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_StageNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Stage].Undo
Exit Sub

err_StageNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Type_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_TypeNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![Type].Undo
Exit Sub

err_TypeNot:
    Call General_Error_Trap
    Exit Sub
End Sub
