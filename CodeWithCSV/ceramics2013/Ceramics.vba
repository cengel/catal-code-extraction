Option Compare Database
Option Explicit

Sub CheckUnitDescript(thisunit)
'when a unit is entered in the diagnostic, unid diagnostic tables etc we must check
'it is also entered into the unit description table
On Error GoTo err_CheckUnitDescript

Dim checknum
checknum = DCount("[Unit]", "Ceramics_Unit_Description", "[Unit] = " & thisunit)
If checknum = 0 Then
    DoCmd.RunSQL "INSERT INTO Ceramics_Unit_Description ([Unit]) VALUES (" & thisunit & ");"
End If


Exit Sub

err_CheckUnitDescript:
    Call General_Error_Trap
    Exit Sub
End Sub
