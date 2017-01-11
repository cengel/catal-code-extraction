Option Compare Database
Option Explicit

Private Sub copy_data_Click()
On Error GoTo Err_copy_data_Click

Dim msg, Style, Title, response
msg = "This action will replace the unit sheet contents, and cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.

' Display message.
response = MsgBox(msg, Style, Title)
If response = vbYes Then    ' User chose Yes.
    ' overwrite records
    Forms![Exca: Unit Sheet]![Bedding] = Me![Bedding]
    Forms![Exca: Unit Sheet]![Inclusions] = Me![Inclusions]
    Forms![Exca: Unit Sheet]![Porosity] = Me![Porosity]
    Forms![Exca: Unit Sheet]![Moisture] = Me![Moisture]
    Forms![Exca: Unit Sheet]![Munsell Colour] = Me![Munsell Colour]
    Forms![Exca: Unit Sheet]![Consistency: Strength] = Me![Consistency: Strength]
    Forms![Exca: Unit Sheet]![Structure] = Me![Structure]
    Forms![Exca: Unit Sheet]![Consistencey: Plasticity and Stickiness] = Me![Consistencey: Plasticity and Stickiness]
    Forms![Exca: Unit Sheet]![Texture] = Me![Texture]
    Forms![Exca: Unit Sheet]![Post-depositional Features] = Me![Post-depositional Features]
    Forms![Exca: Unit Sheet]![Basal Boundary] = Me![Basal Boundary]

Else    ' User chose No.

    ' do nothing.
End If


   

Exit_copy_data_Click:
    Exit Sub

Err_copy_data_Click:
    MsgBox Err.Description
    Resume Exit_copy_data_Click
End Sub


Sub find_unit_Click()
On Error GoTo Err_find_unit_Click


    Screen.PreviousControl.SetFocus
     Unit_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_unit_Click:
    Exit Sub

Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
    
End Sub
Sub Command13_Click()
On Error GoTo Err_Command13_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command13_Click:
    Exit Sub

Err_Command13_Click:
    MsgBox Err.Description
    Resume Exit_Command13_Click
    
End Sub
Sub Command14_Click()
On Error GoTo Err_Command14_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_Command14_Click:
    Exit Sub

Err_Command14_Click:
    MsgBox Err.Description
    Resume Exit_Command14_Click
    
End Sub
Sub Close_Click()
On Error GoTo err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub
