Option Compare Database

Private Sub Update_GID()
Me![GID number] = Me![Unit Number] & "." & Me![X-Find Number]
End Sub

Private Sub Find_Number_AfterUpdate()
Update_GID
End Sub

Private Sub Find_Number_Change()
Update_GID
End Sub


Private Sub Find_Number_Enter()
Update_GID
End Sub

Private Sub Find_Number_Exit(Cancel As Integer)
Update_GID
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
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
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
