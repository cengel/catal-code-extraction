Option Compare Database
Option Explicit

Dim Unit, Letter, findnum, idnum 'these vars will come in as openargs

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
        Response = MsgBox("Please confirm that the crate register should be updated with the following details:" & Chr(13) & Chr(13) & "Findnumber: " & Me![GID] & " in  crate FG" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
        If Response = vbOK Then
            'write to crate register
            Dim sql
            sql = "INSERT INTO [Store: Units in Crates] ([CrateLetter], [CrateNumber], [Unit Number], [Year], [Area], [FindSampleLetter], [FindNumber], [Material], [Description],[LastUpdated]) "
            sql = sql & " VALUES ('FG', " & Me![cboCrateNumber] & ", " & Unit & ","
            If IsNull(Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Year]) Or Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Year] = "" Then
                sql = sql & "null,"
            Else
                sql = sql & Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Year] & ", "
            End If
            If IsNull(Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Area]) Or Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Area] = "" Then
                sql = sql & " null,"
            Else
                sql = sql & "'" & Forms![Frm_MainData]![Frm_Sub_UnitDetails].Form![Area] & "',"
            End If
            
            sql = sql & "'" & Letter & "', " & findnum & ",'Clay',"
            
            If IsNull(Forms![Frm_MainData]![Frm_Sub_ObjectType].Form![ObjectType]) Or Forms![Frm_MainData]![Frm_Sub_ObjectType].Form![ObjectType] = "" Then
                sql = sql & " null,"
            Else
                sql = sql & "'" & Forms![Frm_MainData]![Frm_Sub_ObjectType].Form![ObjectType] & "',"
            End If
            
            
            sql = sql & "#" & Now() & "#);"
            DoCmd.RunSQL sql
            
            sql = "INSERT INTO [Store: Crate Movement by Teams] ([Unit Number], [FindSampleLetter], [FindNumber], [MovedToCrate], [MovedBy], [MovedOn]) VALUES (" & Unit & ",'" & Letter & "'," & findnum & ",'FG" & Me![cboCrateNumber] & "', 'Figurines team', #" & Now() & "#);"
            DoCmd.RunSQL sql
            
            Forms![Frm_MainData]![frm_subform_location].Requery
            Forms![Frm_MainData]![frm_subform_location].Visible = True
            Forms![Frm_MainData]![lblCrateRegMsg].Visible = False
            'DoCmd.GoToControl Forms!Frm_MainData.Name
            'DoCmd.GoToControl Forms![Frm_MainData]![frm_subform_location].Name
            'it has focus so cant be hidden so simply masked by being underneath location subform
            'Forms![Frm_MainData]![cmdLocate].Visible = False
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
'idnumber
On Error GoTo err_open

    If Not IsNull(Me.OpenArgs) Then
        Dim arg
        arg = Me.OpenArgs
        Unit = Left(arg, InStr(arg, ".") - 1)
        Letter = Mid(arg, InStr(arg, ".") + 1, 1)
        findnum = CInt(Right(arg, Len(arg) - (InStr(arg, ".") + 1)))
        
        Me![GID] = arg
        
    Else
        MsgBox "Form opened with incorrect parameters, it will now shut", vbInformation, "Invalid Call"
        DoCmd.Close acForm, Me.Name
        
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
