Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo err_current

    If Me!species <> "" Then
        Dim sql
        sql = "SELECT Shell_Species_LOV.[type number], Shell_Species_LOV.genus, Shell_Species_LOV.DESCRIPTION FROM Shell_Species_LOV WHERE Shell_Species_LOV.genus like  '" & Me![species] & "%';"
    
    Else
        sql = "SELECT Shell_Species_LOV.[type number], Shell_Species_LOV.genus, Shell_Species_LOV.DESCRIPTION FROM Shell_Species_LOV;"
    End If
    Me![type].RowSource = sql
    Me.Refresh

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub species_AfterUpdate()
'when genus selected filter down type list to given numbers
On Error GoTo err_species

    If Me!species <> "" Then
        Dim sql
        sql = "SELECT Shell_Species_LOV.[type number], Shell_Species_LOV.genus, Shell_Species_LOV.DESCRIPTION FROM Shell_Species_LOV WHERE Shell_Species_LOV.genus Like '" & Me![species] & "%';"
    
    Else
        sql = "SELECT Shell_Species_LOV.[type number], Shell_Species_LOV.genus, Shell_Species_LOV.DESCRIPTION FROM Shell_Species_LOV;"
    End If
    Me![type].RowSource = sql
    Me.Refresh

Exit Sub

err_species:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub species_NotInList(NewData As String, response As Integer)
'add new species
On Error GoTo err_species
    Dim sql, respon
    respon = MsgBox("This species does not exist in the list. Do you wish to add it?" & _
                Chr(13) & Chr(13) & "Yes - will add it for use in future " & Chr(13) & _
                "No - enters the value here in this field but does not add it to the general list" & Chr(13) & _
                "Cancel - allows re-selection from the list", _
         vbYesNoCancel)
    'If MsgBox("This species does not exist in the list. Do you wish to add it?", _
    '     vbOKCancel) = vbOK Then
        ' Set Response argument to indicate that datais being added.
   If respon = vbYes Then
        response = acDataErrAdded
        ' Add string in NewData argument to row source.
        Dim retstr
        retstr = InputBox("Please enter the type number to match this species:", "Type Number")
        If retstr <> "" Then
            sql = "INSERT INTO Shell_Species_LOV ([genus], [type number]) VALUES ('" & NewData & "', " & retstr & ");"
            DoCmd.RunSQL sql
        Else
            response = acDataErrContinue
            Me![species].Undo
        End If
    ElseIf respon = vbNo Then
        Me![species].LimitToList = False
        response = acDataErrContinue
        Me![species] = NewData
        Me![species].LimitToList = True
        DoCmd.GoToControl "type letter"
    Else
        ' Cancel = suppress error message and undo changes.
        response = acDataErrContinue
        Me![species].Undo
    End If

Exit Sub

err_species:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub type_NotInList(NewData As String, response As Integer)
'add new number
On Error GoTo err_type
    Dim sql, respon
    respon = MsgBox("This type number does not exist in the list for the genus '" & Me![species] & "'. Do you wish to add it?" & _
                Chr(13) & Chr(13) & "Yes - will add it for use with this genus in future " & Chr(13) & _
                "No - enters the value here in this field but does not add it to the general list" & Chr(13) & _
                "Cancel - allows re-selection from the list", _
         vbYesNoCancel)
     If respon = vbYes Then
        ' Set Response argument to indicate that datais being added.
        response = acDataErrAdded
        ' Add string in NewData argument to row source.
        sql = "INSERT INTO Shell_Species_LOV ([genus], [type number]) VALUES ('" & Me![species] & "', '" & NewData & "');"
        DoCmd.RunSQL sql
    ElseIf respon = vbNo Then
        Me![type].LimitToList = False
        response = acDataErrContinue
        Me![type] = NewData
        Me![type].LimitToList = True
        DoCmd.GoToControl "type letter"
    Else
        ' Cancel = suppress error message and undo changes.
        response = acDataErrContinue
        Me![type].Undo
    End If

Exit Sub

err_type:
    Call General_Error_Trap
    Exit Sub
End Sub
