Option Compare Database

Private Sub Update_GID()
Me![GID number] = Me![Unit number] & "." & Me![X-Find Number]
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
Me![Date Changed] = Now()
End Sub


