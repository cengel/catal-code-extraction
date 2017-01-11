Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




Private Sub cmdAddNew_Click()
'v9.2 SAJ - add a new record
On Error GoTo err_cmdAddNew_Click

    DoCmd.RunCommand acCmdRecordsGoToNew

Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
'SAJ - close the form
    DoCmd.Close acForm, Me.Name
End Sub




Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
       MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
   End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
