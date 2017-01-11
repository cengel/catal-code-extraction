Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)

Forms![Store: Crate Register]![Date Changed] = Now()

End Sub
