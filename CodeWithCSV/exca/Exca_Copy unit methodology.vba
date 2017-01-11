Option Compare Database
Option Explicit
'**************************************************************
' Functionality tightened in v9.1
'**************************************************************
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
    Me.RecordSource = "SELECT * FROM [Exca: Unit Sheet] WHERE [Unit Number] = " & Me![cboFindUnitToCopy]
    Me![Unit Number].ControlSource = "Unit Number"
    Me![Recognition].ControlSource = "Recognition"
    Me![Definition].ControlSource = "Definition"
    Me![Execution].ControlSource = "Execution"
    Me![Condition].ControlSource = "Condition"
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
msg = msg & "fields: Recognition, Definition, Execution and Definition with those of unit " & Me![Unit Number] & " shown here." & Chr(13) & Chr(13)
msg = msg & "This action cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.

' Display message.
response = MsgBox(msg, Style, Title)
If response = vbYes Then    ' User chose Yes.
    ' overwrite records
    Forms![Exca: Unit Sheet]![Recognition] = Me![Recognition]
    Forms![Exca: Unit Sheet]![Definition] = Me![Definition]
    Forms![Exca: Unit Sheet]![Execution] = Me![Execution]
    Forms![Exca: Unit Sheet]![Condition] = Me![Condition]
  
Else    ' User chose No.

    ' do nothing.
End If


   

Exit_copy_data_Click:
    Exit Sub

Err_copy_data_Click:
    Call General_Error_Trap
    Resume Exit_copy_data_Click
End Sub





Sub Close_Click()
'**************************************************************
' Mainly original code, intro name of form and error trap
' SAJ v.91
'**************************************************************
On Error GoTo err_close_Click


    DoCmd.Close acForm, "Exca: copy unit methodology"

Exit_close_Click:
    Exit Sub

err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
    
End Sub
