Option Compare Database
Option Explicit


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
