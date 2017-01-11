Option Compare Database
Option Explicit

Dim Unit, Letter, findnum, Bag 'these vars will come in as openargs

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
Dim Response, sql
    If Me![cboCrateNumber] <> "" Then
        If Me![GID] <> "" Then
            Response = MsgBox("Please confirm that the crate register should be updated with the following details:" & Chr(13) & Chr(13) & "Findnumber: " & Me![GID] & " in  crate OB" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
            If Response = vbOK Then
                'write to crate register
                sql = "INSERT INTO [Store: Units in Crates] ([CrateLetter], [CrateNumber], [Unit Number], [Year], [Area], [FindSampleLetter], [FindNumber], [Material], [Bag], [LastUpdated], [Description]) "
                sql = sql & " VALUES ('OB', " & Me![cboCrateNumber] & ", " & Unit & ","
                
                Dim getyear, getArea
                getyear = DLookup("[Year]", "[Exca: Unit Sheet with Relationships]", "[Unit number] = " & Unit)
                If IsNull(getyear) Or getyear = "" Then
                    sql = sql & "null,"
                Else
                    sql = sql & getyear & ", "
                End If
                getArea = DLookup("[area]", "[Exca: Unit Sheet with Relationships]", "[Unit number] = " & Unit)
                If IsNull(getArea) Or getArea = "" Then
                    sql = sql & " null,"
                Else
                    sql = sql & "'" & getArea & "',"
                End If
                
                sql = sql & "'" & Letter & "', " & findnum & ",'" & Forms![frm_CS_stagetwo]![RawMaterial] & "', '" & Forms![frm_CS_stagetwo]![Bag] & "', #" & Now() & "#"
                sql = sql & ",'" & Forms![frm_CS_stagetwo]![cboCategory] & "');"
                DoCmd.RunSQL sql
                
                sql = "INSERT INTO [Store: Crate Movement by Teams] ([Unit Number], [FindSampleLetter], [FindNumber], [MovedToCrate], [MovedBy], [MovedOn]) VALUES (" & Unit & ",'" & Letter & "'," & findnum & ",'OB" & Me![cboCrateNumber] & "', 'ChippedStone team', #" & Now() & "#);"
                DoCmd.RunSQL sql
                
                Forms![frm_CS_stagetwo]![frm_subform_location_object].Requery
                
                DoCmd.Close acForm, Me.Name
             End If
        ElseIf Me![txtUnit] <> "" Then
            Response = MsgBox("Please confirm that the crate register should be updated with the following details:" & Chr(13) & Chr(13) & "Unit: " & Me![txtUnit] & ", Bag: " & Me![txtBag] & " in  crate OB" & Me![cboCrateNumber], vbQuestion + vbOKCancel, "Confirm Action")
            If Response = vbOK Then
                'write to crate register
                sql = "INSERT INTO [Store: Units in Crates] ([CrateLetter], [CrateNumber], [Unit Number], [Year], [Area], [Bag], [Material], [LastUpdated]) "
                sql = sql & " VALUES ('OB', " & Me![cboCrateNumber] & ", " & Me!txtUnit & ","
                If IsNull(Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Year]) Or Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Year] = "" Then
                    sql = sql & "null,"
                Else
                    sql = sql & Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Year] & ", "
                End If
                If IsNull(Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Area]) Or Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Area] = "" Then
                    sql = sql & " null,"
                Else
                    sql = sql & "'" & Forms![frm_CS_basicdata]![Frm_subformUnitDetails].Form![Area] & "',"
                End If
                
                sql = sql & "'" & Me![txtBag] & "', '" & Forms![frm_CS_basicdata]![RawMaterial] & "',#" & Now() & "#);"
                DoCmd.RunSQL sql
                
                sql = "INSERT INTO [Store: Crate Movement by Teams] ([Unit Number], [MovedToCrate], [MovedBy], [MovedOn]) VALUES (" & Unit & ",'OB" & Me![cboCrateNumber] & "', 'ChippedStone team', #" & Now() & "#);"
                DoCmd.RunSQL sql
                
                Forms![frm_CS_basicdata]!frm_subform_location.Requery
                
                DoCmd.Close acForm, Me.Name
             End If
        
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
        If InStr(Me.OpenArgs, "BAG") > 0 Then
            Me!GID.Visible = False
            Me!txtBag.Visible = True
            Me!txtUnit.Visible = True
            'comes from basic data
            Bag = Mid(Me.OpenArgs, 4, InStr(Me.OpenArgs, ":") - 4)
            getrest = Right(Me.OpenArgs, Len(Me.OpenArgs) - (InStr(Me.OpenArgs, ":")))
        
            Unit = Left(getrest, InStr(getrest, ":") - 1)

            Me![txtUnit] = Unit
            Me![txtBag] = Bag
            
        Else
            Me!GID.Visible = True
            Me!txtBag.Visible = False
            Me!txtUnit.Visible = False
            
            Unit = Left(Me.OpenArgs, InStr(Me.OpenArgs, ":") - 1)
            getrest = Right(Me.OpenArgs, Len(Me.OpenArgs) - (InStr(Me.OpenArgs, ":")))
            
            Letter = Left(getrest, InStr(getrest, ":") - 1)
            findnum = Right(getrest, Len(getrest) - (InStr(getrest, ":")))
            
            Me![GID] = Unit & "." & Letter & findnum
        End If
    Else
        MsgBox "Form opened with incorrect parameters, it will now shut", vbInformation, "Invalid Call"
        DoCmd.Close acForm, Me.Name
        
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
