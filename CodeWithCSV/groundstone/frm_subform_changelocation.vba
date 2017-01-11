Option Compare Database
Option Explicit

Dim Unit, Letter, findnum, currentCrate 'these vars will come in as openargs

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
        Response = MsgBox("Please confirm that the crate register should be updated with the following details: " & Chr(13) & Chr(13) & "Findnumber: " & Me![GID] & " moves from " & Me![txtOldCrate] & " to crate GS" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
        If Response = vbOK Then
            'write to crate register
            Dim sql
             sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
             sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, 'GS" & Me![cboCrateNumber] & "' as MovedToCrate, 'Groundstone Team', #" & Now & "# "
             sql = sql & " FROM [Store: Units in Crates] "
             sql = sql & " WHERE [Unit Number] = " & Unit & " AND [FindSampleLetter] = '" & Letter & "' AND [FindNumber] = " & findnum & ";"
            DoCmd.RunSQL sql
            
            sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & Me![cboCrateNumber] & " WHERE [CrateLetter] = 'GS' AND [CrateNumber] = " & Replace(currentCrate, "GS", "") & " AND [Unit number] = " & Unit & " AND [FindSampleLetter] = '" & Letter & "' AND [FindNumber] = " & findnum & ";"
            DoCmd.RunSQL sql
            
           
            
            Forms![Frm_Basic_Data]![frm_subform_basic].Requery
            
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
'unit number:lettercode:findnumber:currentcratelocation
On Error GoTo err_open

    If Not IsNull(Me.OpenArgs) Then
        Dim getrest
        Unit = Left(Me.OpenArgs, InStr(Me.OpenArgs, ":") - 1)
        getrest = Right(Me.OpenArgs, Len(Me.OpenArgs) - (InStr(Me.OpenArgs, ":")))
        
        Letter = Left(getrest, InStr(getrest, ":") - 1)
        
        getrest = Right(getrest, Len(getrest) - (InStr(getrest, ":")))
        findnum = Left(getrest, InStr(getrest, ":") - 1)
        currentCrate = Mid(getrest, InStr(getrest, ":") + 1)
        
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
