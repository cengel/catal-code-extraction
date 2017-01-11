Option Compare Database

Private Sub cboName_NotInList(NewData As String, Response As Integer)
'On Error GoTo err_cboName_NotInList

'Dim retVal, sql, getfirst, getsurname
'retVal = MsgBox("This Name does not appear in the pre-defined list. Have you checked the list to make sure there is no match?", vbQuestion + vbYesNo, "New Conservator Name")
'If retVal = vbYes Then
'    'allow value,
'     Response = acDataErrAdded
'
'    getfirst = InputBox("Please enter the first name of this conservator", "First Name")
'    getsurname = InputBox("Plese enter the surname of this conservator", "Surname")
'
'    'sql = "INSERT INTO [Conservation_Code_ConservatorNames] ([ConservatorFirstName], [ConservatorSurname]) VALUES ('" & getfirst & "', '" & getsurname & "');"
'    sql = "INSERT INTO [Conservation_Code_ConservatorNames] ([ConservatorFirstName], [ConservatorSurname]) VALUES ('" & NewData & "', '" & getsurname & "');"
'    DoCmd.RunSQL sql
'
'Else
'    'no leave it so they can edit it
'    Response = acDataErrContinue
'End If
'Exit Sub
'
'err_cboName_NotInList:
'    Call General_Error_Trap
'    Exit Sub

End Sub
