Option Compare Database

Private Sub GSAnalyst_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_GSAnalyst_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    inputname = InputBox("Please enter the full name of the analyst to go along with these initials:", "Analyst Name")
    If inputname <> "" Then
        sql = "INSERT INTO [ChippedStoneLOV_Analyst]([CSAnalystInitials], [CSAnalystName]) VALUES ('" & NewData & "', '" & inputname & "');"
        DoCmd.RunSQL sql
    Else
        Response = acDataErrContinue
    End If
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_GSAnalyst_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub
