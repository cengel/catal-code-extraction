Option Compare Database
Option Explicit
Private Sub Update_Family_GID()
Me![Family_GID] = Me![Family] & "." & Me![Genus] & "." & Me![Species] & "." & Me![SubSpecies] & "." & Me![Variety] & "." & Me![Part/Type]
End Sub


Private Sub Family_AfterUpdate()
Update_Family_GID
End Sub

Private Sub Family_Change()
Update_Family_GID
End Sub


Private Sub Genus_AfterUpdate()
Update_Family_GID
End Sub


Private Sub Genus_Change()
Update_Family_GID
End Sub


Private Sub Part_Type_AfterUpdate()
Update_Family_GID
End Sub

Private Sub Part_Type_Change()
Update_Family_GID
End Sub


Private Sub Species_AfterUpdate()
Update_Family_GID
End Sub


Private Sub Species_Change()
Update_Family_GID
End Sub


Private Sub SubSpecies_AfterUpdate()
Update_Family_GID
End Sub


Private Sub SubSpecies_Change()
Update_Family_GID
End Sub


Private Sub Variety_AfterUpdate()
Update_Family_GID
End Sub


Private Sub Variety_Change()
Update_Family_GID
End Sub


