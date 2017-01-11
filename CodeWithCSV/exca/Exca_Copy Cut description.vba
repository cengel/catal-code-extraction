Option Compare Database
Option Explicit

Private Sub cboFindUnitToCopy_AfterUpdate()
'**************************************************************
' This combo has replaced the old 'Find unit' button that simply
' brought up the Find/Replace dialog. This allows the user to select
' the unit to copy from it then sets the control source of the
' fields below (the form opens blank so there is nothing the user can do
' until they select a Unit
' SAJ v.91
'**************************************************************
On Error GoTo err_cboFindUnitToCopy_AfterUpdate

If Me![cboFindUnitToCopy] <> "" Then
    Me.RecordSource = "SELECT * FROM [Exca: Descriptions Cut] WHERE [Unit Number] = " & Me![cboFindUnitToCopy]
    Me![Unit Number].ControlSource = "Unit Number"
    Me![Shape].ControlSource = "Shape"
    Me![Corners].ControlSource = "Corners"
    Me![Top Break].ControlSource = "Top Break"
    Me![Sides].ControlSource = "Sides"
    Me![Base Break].ControlSource = "Base Break"
    Me![Base].ControlSource = "Base"
    Me![Orientation].ControlSource = "Orientation"
    Me![All Layers within].ControlSource = "All Layers within"
    Me![copy data].Enabled = True
End If
Exit Sub

err_cboFindUnitToCopy_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub copy_data_Click()
'**************************************************************
' Mainly original code, intro more into message and error trap
' SAJ v.91
'**************************************************************
On Error GoTo Err_copy_data_Click

Dim msg, Style, Title, response
msg = "This action will replace the unit sheet (" & Me![Text17] & ") "
msg = msg & "data with with that of unit " & Me![Unit Number] & " shown here." & Chr(13) & Chr(13)
msg = msg & "This action cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.

' Display message.
response = MsgBox(msg, Style, Title)
If response = vbYes Then    ' User chose Yes.
    ' overwrite records
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Shape] = Me![Shape]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Corners] = Me![Corners]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Top Break] = Me![Top Break]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Sides] = Me![Sides]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Base Break] = Me![Base Break]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Base] = Me![Base]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Orientation] = Me![Orientation]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![All Layers within] = Me![All Layers within]

Else    ' User chose No.

    ' do nothing.
End If


   

Exit_copy_data_Click:
    Exit Sub

Err_copy_data_Click:
    Call General_Error_Trap
    Resume Exit_copy_data_Click
End Sub


Sub find_unit_Click()
'replaced by cboFindUnitToCopy_AfterUpdate
'On Error GoTo Err_find_unit_Click
'
'
'    Screen.PreviousControl.SetFocus
'     Unit_Number.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
'
'Exit_find_unit_Click:
'    Exit Sub
'
'Err_find_unit_Click:
'    MsgBox Err.Description
'    Resume Exit_find_unit_Click
'
End Sub

Sub Close_Click()
'**************************************************************
' Mainly original code, intro name of form and error trap
' SAJ v.91
'**************************************************************
On Error GoTo err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

err_close_Click:
   Call General_Error_Trap
    Resume Exit_close_Click
    
End Sub
