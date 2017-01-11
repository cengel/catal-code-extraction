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
    Me.RecordSource = "SELECT * FROM [Exca: Descriptions Layer] WHERE [Unit Number] = " & Me![cboFindUnitToCopy]
    Me![Unit Number].ControlSource = "Unit Number"
    Me![Consistency].ControlSource = "Consistency"
    Me![Colour].ControlSource = "Colour"
    Me![Texture].ControlSource = "Texture"
    Me![Bedding].ControlSource = "Bedding"
    Me![Inclusions].ControlSource = "Inclusions"
    Me![Post-depositional Features].ControlSource = "Post-depositional Features"
    Me![Basal Boundary].ControlSource = "Basal Boundary"
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
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Consistency] = Me![Consistency]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Colour] = Me![Colour]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Texture] = Me![Texture]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Bedding] = Me![Bedding]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Inclusions] = Me![Inclusions]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Post-depositional Features] = Me![Post-depositional Features]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Basal Boundary] = Me![Basal Boundary]
 
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
    
End Sub


Sub Close_Click()
'**************************************************************
' Mainly original code, intro name of form and error trap
' SAJ v.91
'**************************************************************
On Error GoTo err_close_Click


    DoCmd.Close acForm, "Exca: Copy layer description"

Exit_close_Click:
    Exit Sub

err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
    
End Sub
