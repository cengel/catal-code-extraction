Option Compare Database
Option Explicit

Dim Unit, Letter, findnum 'these vars will come in as openargs

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
'if user has selected a crate ask for confirmation and then place the object in the Units in Crates register
'adding a record for the Finds Officer that this has been done
On Error GoTo err_cmdOK

    If Me![cboCrateNumber] <> "" Then
        Dim Response
        Response = MsgBox("Please confirm that the crate register should be updated with the following details:" & Chr(13) & Chr(13) & "Findnumber: " & Me![GID] & " in  crate GS" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
        If Response = vbOK Then
            'write to crate register
            Dim sql
            sql = "INSERT INTO [Store: Units in Crates] ([CrateLetter], [CrateNumber], [Unit Number], [Year], [Area], [FindSampleLetter], [FindNumber], [Material], [LastUpdated]) "
            sql = sql & " VALUES ('GS', " & Me![cboCrateNumber] & ", " & Unit & ","
            If IsNull(Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Year]) Or Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Year] = "" Then
                sql = sql & "null,"
            Else
                sql = sql & Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Year] & ", "
            End If
            If IsNull(Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Area]) Or Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Area] = "" Then
                sql = sql & " null,"
            Else
                sql = sql & "'" & Forms![Frm_Basic_Data]![frm_subform_basic_2013].Form![Area] & "',"
            End If
            
            sql = sql & "'" & Letter & "', " & findnum & ",'Stone', #" & Now() & "#);"
            DoCmd.RunSQL sql
            
            sql = "INSERT INTO [Store: Crate Movement by Teams] ([Unit Number], [FindSampleLetter], [FindNumber], [MovedToCrate], [MovedBy], [MovedOn]) VALUES (" & Unit & ",'" & Letter & "'," & findnum & ",'GS" & Me![cboCrateNumber] & "', 'Groundstone team', #" & Now() & "#);"
            DoCmd.RunSQL sql
            
            Forms![Frm_Basic_Data]![frm_subform_basic_2013].Requery
            
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
'unit number:lettercode:findnumber
On Error GoTo err_open

    If Not IsNull(Me.OpenArgs) Then
        Dim getrest
        Unit = Left(Me.OpenArgs, InStr(Me.OpenArgs, ":") - 1)
        getrest = Right(Me.OpenArgs, Len(Me.OpenArgs) - (InStr(Me.OpenArgs, ":")))
        
        Letter = Left(getrest, InStr(getrest, ":") - 1)
        findnum = Right(getrest, Len(getrest) - (InStr(getrest, ":")))
        
        Me![GID] = Unit & "." & Letter & findnum
        
    Else
        MsgBox "Form opened with incorrect parameters, it will now shut", vbInformation, "Invalid Call"
        DoCmd.Close acForm, Me.Name
        
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
