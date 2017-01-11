Option Compare Database

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date Changed] = Now()
End Sub


