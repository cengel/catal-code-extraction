Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)

Forms![AdminCrateRegister]![Date Changed] = Now()

End Sub
