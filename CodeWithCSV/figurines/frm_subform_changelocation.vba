Option Compare Database
Option Explicit

Dim Unit, Letter, findnum, currentCrate, idnum 'these vars will come in as openargs

Private Sub cmdCancel_Click()
'cancel operation - do nothing
On Error GoTo err_cancel

DoCmd.Close acForm, Me.Name

Exit Sub

err_cancel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOK_Click()
'if user has selected a crate ask for confirmation and movethe object in the Units in Crates register
'adding a record for the Finds Officer that this has been done
On Error GoTo err_cmdOK

    If Me![cboCrateNumber] <> "" Then
        Dim Response
        Response = MsgBox("Please confirm that the crate register should be updated with the following details: " & Chr(13) & Chr(13) & "Findnumber: " & Me![GID] & " moves from " & Me![txtOldCrate] & " to crate FG" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
        If Response = vbOK Then
            'write to crate register
            Dim sql
             sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
             sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, 'FG" & Me![cboCrateNumber] & "' as MovedToCrate, 'Figurines Team', #" & Now & "# "
             sql = sql & " FROM [Store: Units in Crates] "
             sql = sql & " WHERE [Unit Number] = " & Unit & " AND [FindSampleLetter] = '" & Letter & "' AND [FindNumber] = " & findnum & ";"
            DoCmd.RunSQL sql
            
            sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & Me![cboCrateNumber] & " WHERE [CrateLetter] = 'FG' AND [CrateNumber] = " & Replace(currentCrate, "FG", "") & " AND [Unit number] = " & Unit & " AND [FindSampleLetter] = '" & Letter & "' AND [FindNumber] = " & findnum & ";"
            DoCmd.RunSQL sql
            
            Forms![Frm_MainData]![frm_subform_location].Requery
            Forms![Frm_MainData]![frm_subform_location].Form.Refresh
            
            DoCmd.Close acForm, Me.Name
            
        End If
    Else
        MsgBox "No crate number selected", vbInformation, "Invalid Selection"
        
    End If

Exit Sub

err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'get open args. They will take this format
'idnumber:currentcrate
On Error GoTo err_open

    If Not IsNull(Me.OpenArgs) Then
        Dim getrest, arg
        arg = Me.OpenArgs
        Unit = Left(arg, InStr(arg, ".") - 1)
        
        Letter = Mid(arg, InStr(arg, ".") + 1, 1)
        
        findnum = Mid(arg, InStr(arg, ".") + 2, InStr(arg, ":") - (InStr(arg, ".") + 2))
        
        '''findnum = CInt(Right([ID number], Len([ID number]) - (InStr([ID number], ".") + 1)))
        'getrest = Right(Me.OpenArgs, Len(Me.OpenArgs) - (InStr(Me.OpenArgs, ":")))
        
        'Letter = Left(getrest, InStr(getrest, ":") - 1)
        
        'getrest = Right(getrest, Len(getrest) - (InStr(getrest, ":")))
        'findnum = Left(getrest, InStr(getrest, ":") - 1)
        'currentCrate = Mid(getrest, InStr(getrest, ":") + 1)
        currentCrate = Right(arg, Len(arg) - InStr(arg, ":"))
        
        Me![GID] = Unit & "." & Letter & findnum
        Me![txtOldCrate] = currentCrate
    Else
        MsgBox "Form opened with incorrect parameters, it will now shut", vbInformation, "Invalid Call"
        DoCmd.Close acForm, Me.Name
        
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
