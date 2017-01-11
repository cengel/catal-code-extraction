Option Compare Database
Option Explicit
Private Sub FindFacility(what)
'original code moved from Find Unit button - kept very basic
'saj season 2006
On Error GoTo Err_find_unit_Click


    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim message As String, title As String, Unit As String, default As String
    Dim Material As String, descrip As String
    If what = "unit" Then
        message = "Enter a unit number"   ' Set prompt.
        title = "Searching Crate Register" ' Set title.
        default = "1000"   ' Set default.
        ' Display message, title, and default value.
        Unit = InputBox(message, title, default)
        If Unit = "" Then Exit Sub 'saj catch no entry
        'stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
        'saj - jules wants to just find numbers directly
        stLinkCriteria = "[Unit Number] =" & Unit
    ElseIf what = "material" Then
        message = "Enter a material"   ' Set prompt.
        title = "Searching Crate Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        Material = InputBox(message, title, default)
        If Material = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Material] like '*" & Material & "*'"
    ElseIf what = "descrip" Then
        message = "Enter a description"   ' Set prompt.
        title = "Searching Crate Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        descrip = InputBox(message, title, default)
        If descrip = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[TempDescription] like '*" & descrip & "*'"
    ElseIf what = "find" Then
        message = "Enter a description"   ' Set prompt.
        title = "Searching Crate Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        Dim un, lett, num
        message = "Enter a Unit"   ' Set prompt.
        un = InputBox(message, title, default)
        If un = "" Then Exit Sub 'saj catch no entry
        message = "Enter a letter code"   ' Set prompt.
        lett = InputBox(message, title, default)
        If lett = "" Then Exit Sub 'saj catch no entry
        message = "Enter a number"   ' Set prompt.
        num = InputBox(message, title, default)
        If num = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Unit number] =" & un & " AND [FindLetter] ='" & lett & "' AND [FindNumber] = " & num
        
    Else
        Exit Sub
    End If
    stDocName = "Store: Find Unit in Crate2"
    'stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
    'DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria
    
Exit_find_unit_Click:
    Exit Sub

Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
End Sub

Private Sub Update_GID()
'sub used by crate fields written
On Error GoTo err_updategid

Me![txtFullCrateName] = Me![cboCrateLetter] & Me![txtCrateNumber]
If Me![cboCrateLetter] <> "" And Me![txtCrateNumber] <> "" Then
    Me.Refresh
End If
Exit Sub

err_updategid:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboCrateLetter_AfterUpdate()
'update field that holds crate number and letter together
On Error GoTo err_cboCrate

    Update_GID
    
Exit Sub

err_cboCrate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cboCrateLetter_NotInList(NewData As String, Response As Integer)
'only admin can add new codes
On Error GoTo err_new

    If GetGeneralPermissions = "Admin" Then
        Response = acDataErrContinue
        Dim retVal
        retVal = MsgBox("Are you sure you want to add this completely new crate code prefix", vbQuestion + vbYesNo, "Confirm New Code")
        If retVal = vbYes Then
            Me![cboCrateLetter].LimitToList = False
            Me![txtCrateNumber] = 1
            Me![txtFullCrateName] = Me![cboCrateLetter] & Me![txtCrateNumber]
            DoCmd.RunCommand acCmdSaveRecord
            Me![cboCrateLetter].LimitToList = True
        End If
    End If

Exit Sub

err_new:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected crate from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
        DoCmd.GoToControl "txtFullCrateName"
        DoCmd.FindRecord Me![cboFindUnit]
        Me![cboFindUnit] = ""
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboFindUnit].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtCrateNumber].Enabled = True
    Me![txtCrateNumber].Locked = False
    Me![txtCrateNumber].BackColor = 16777215
    Me![cboCrateLetter].Enabled = True
    Me![cboCrateLetter].Locked = False
    Me![cboCrateLetter].BackColor = 16777215
 
    DoCmd.GoToControl "cboCrateLetter"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCodes_Click()
'just a quick view of the crate codes
On Error GoTo err_codes

    DoCmd.OpenForm "frm_pop_cratecodes", acNormal, , , acFormReadOnly, acDialog
    
Exit Sub

err_codes:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'allow administrators to delete a crate and all its contents
On Error GoTo err_delete

    Dim retVal
    retVal = MsgBox("This action will delete crate " & Me![Crate Number] & " and all its contents, are you really sure you want to delete this crate?", vbCritical + vbYesNo, "Confirm Deletion")
    If retVal = vbYes Then
        'ok proceed
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace, sql1, sql2, myq As QueryDef
        Dim myrs As DAO.Recordset
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
        
        'this method doesn't seem to work as it says you need to use dbSeeChanges for tables that use an identity field
        'sql1 = "DELETE FROM [Store: Units in Crates] WHERE [Crate Number] = '" & Me![Crate Number] & "';"
        'sql2 = "DELETE FROM [Store: Crate Register] WHERE [Crate Number] = '" & Me![Crate Number] & "';"
        'Set myq = mydb.CreateQueryDef("")
        'myq.sql = sql1
        'myq.Execute
                
        'myq.close
        'Set myq = Nothing
        
        'Set myq = mydb.CreateQueryDef("")
        'myq.sql = sql2
        'myq.Execute
                
        'myq.close
        'Set myq = Nothing
        If spString <> "" Then


            Set myq = mydb.CreateQueryDef("")
            myq.Connect = spString
            myq.ReturnsRecords = False
            myq.sql = "sp_Store_Delete_AllCrateContents " & Me![txtCrateNumber] & ",'" & Me![cboCrateLetter] & "'"
            myq.Execute
    
            myq.sql = "sp_Store_Delete_Crate " & Me![txtCrateNumber] & ",'" & Me![cboCrateLetter] & "'"
            myq.Execute
            myq.Close
            Set myq = Nothing
    
        Else
            MsgBox "Sorry but this crate cannot be deleted at the moment, restart the database and try again", vbCritical, "Error"
        
        End If
        
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Crate Deleted"
            Me.Requery
            Me![cboFindUnit].Requery
        Else
            wrkdefault.Rollback
            MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    
    End If
    
    
Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdFindDescrip_Click()
Call FindFacility("Descrip")

End Sub

Private Sub cmdLocateFind_Click()
Call FindFacility("find")
End Sub

Private Sub cmdMaterial_Click()
Call FindFacility("material")

End Sub

Private Sub cmdPrint_Click()
On Error GoTo err_print

    Dim stDocName As String

    stDocName = "Finds Store: Crate Register"
    DoCmd.OpenReport stDocName, acPreview, , "[Crate Number] = '" & Me![txtFullCrateName] & "'"


    Exit Sub

err_print:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMove_Click()
'new season 2007 - move entire contents of one crate into another
On Error GoTo err_cmdMove

    DoCmd.OpenForm "frm_subform_AdminMoveCrateContents", acNormal, , , acFormPropertySettings, acDialog

Exit Sub

err_cmdMove:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdRenameCrate_Click()
'rename a crate and associate contents with new name
On Error GoTo err_cmdRename_Click

    Dim retVal, nwname, nwnum, sql1, sql2
    nwname = InputBox("Please enter the new Crate name below - just characters here, the number will entered next", "Crate Name")
    If nwname <> "" Then
        nwname = UCase(nwname)
        nwnum = InputBox("Please enter the new Crate number - just numbers here", "Crate Number")
        If nwnum <> "" Then
            retVal = MsgBox("This action will rename crate " & Me![Crate Number] & " and all its contents to " & nwname & nwnum & ", are you really sure you want to continue?", vbCritical + vbYesNo, "Confirm Rename")
            If retVal = vbYes Then
                sql1 = "UPDATE [Store: Units in Crates] SET [Store: Units in Crates].[Crate Number] = '" & nwname & nwnum & "', [Store: Units in Crates].CrateNumber = " & nwnum & ", [Store: Units in Crates].CrateLetter = '" & nwname & "' WHERE [Crate Number] ='" & Me![Crate Number] & "';"
                sql2 = "UPDATE [Store: Crate Register] SET [Store: Crate Register].[Crate Number] = '" & nwname & nwnum & "', [Store: Crate Register].CrateNumber = " & nwnum & ", [Store: Crate Register].CrateLetter = '" & nwname & "' WHERE [Crate Number] ='" & Me![Crate Number] & "';"
                On Error Resume Next
                Dim mydb As DAO.Database, wrkdefault As Workspace, myq As QueryDef
                Set wrkdefault = DBEngine.Workspaces(0)
                Set mydb = CurrentDb
        
                ' Start of outer transaction.
                wrkdefault.BeginTrans
                Set myq = mydb.CreateQueryDef("")
                myq.sql = sql1
                myq.Execute
    
                myq.sql = sql2
                myq.Execute
                
                myq.Close
                Set myq = Nothing
            
                If Err.Number = 0 Then
                    wrkdefault.CommitTrans
                    MsgBox "Crate Renamed"
                    Me.Requery
                    Me![cboFindUnit].Requery
                    DoCmd.GoToControl "txtFullCrateName"
                    DoCmd.FindRecord nwname & nwnum
                Else
                    wrkdefault.Rollback
                    MsgBox "A problem has occured and the rename has been cancelled. The error message is: " & Err.Description
                End If

                mydb.Close
                Set mydb = Nothing
                wrkdefault.Close
                Set wrkdefault = Nothing
            End If
        End If
    End If
    
    
Exit Sub

err_cmdRename_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command21_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub find_unit_Click()

Call FindFacility("unit")

'On Error GoTo Err_find_unit_Click
'
'
'    Dim stDocName As String
'    Dim stLinkCriteria As String
'    Dim message As String, title As String, Unit As String, default As String
'
'message = "Enter a unit number"   ' Set prompt.
'title = "Searching Crate Register" ' Set title.
'default = "1000"   ' Set default.
'' Display message, title, and default value.
'Unit = InputBox(message, title, default)
'
'    stDocName = "Store: Find Unit in Crate2"
'    stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
'    'DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria
'    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria
'
'Exit_find_unit_Click:
'    Exit Sub
'
'Err_find_unit_Click:
'    MsgBox Err.Description
'    Resume Exit_find_unit_Click
    
End Sub


Private Sub Form_AfterUpdate()
'moved from before update
On Error GoTo err_after

'this was looping and not letting move on thro records - dirty check seems to cure this
If Me.Dirty Then
    Me![Date Changed] = Now()
End If
Exit Sub

err_after:
    Call General_Error_Trap
    Exit Sub
End Sub





Sub find_Click()
On Error GoTo Err_find_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_Click:
    Exit Sub

Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub


Sub Close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub

Private Sub Form_Current()
'new code for 2006
On Error GoTo err_current

    If (Me![cboCrateLetter] = "" Or IsNull(Me![cboCrateLetter])) And (Me![txtCrateNumber] = "" Or IsNull(Me![txtCrateNumber])) Then
        'new record allow GID entry
        Me![cboCrateLetter].Enabled = True
        Me![cboCrateLetter].Locked = False
        Me![cboCrateLetter].BackColor = 16777215
        Me![txtCrateNumber].Enabled = True
        Me![txtCrateNumber].Locked = False
        Me![txtCrateNumber].BackColor = 16777215
        Me![cboMainMaterial].Locked = False
        Me![cboMainMaterial].Enabled = True
        Me![cboMainMaterial].BackStyle = 1
    Else
        'existing entry lock
        Me![cboCrateLetter].Enabled = False
        Me![cboCrateLetter].Locked = True
        Me![cboCrateLetter].BackColor = Me.Section(0).BackColor
        Me![txtCrateNumber].Enabled = False
        Me![txtCrateNumber].Locked = True
        Me![txtCrateNumber].BackColor = Me.Section(0).BackColor
        Me![cboMainMaterial].Locked = True
        Me![cboMainMaterial].Enabled = False
        Me![cboMainMaterial].BackStyle = 0
    End If
    
    If GetGeneralPermissions = "Admin" Then
        Me![cboMainMaterial].Locked = False
        Me![cboMainMaterial].Enabled = True
        Me![cboMainMaterial].BackStyle = 1
    End If
Exit Sub


err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'on error goto err_open

If GetGeneralPermissions = "Admin" Then
    Me![cmdDelete].Visible = True
    Me![cmdRenameCrate].Visible = True
    'jules request this hidden season 2008 - v3.1
    'Me![cmdMove].Visible = True
Else
    Me![cmdDelete].Visible = False
    Me![cmdRenameCrate].Visible = False
     'jules request this stay hidden season 2008 - v3.1
    'Me![cmdMove].Visible = False
End If

DoCmd.GoToControl "cboFindUnit"
Exit Sub


err_open:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    Call General_Error_Trap
    Resume Exit_go_next_Click
End Sub

Private Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous2_Click:
    Exit Sub

Err_go_previous2_Click:
    Call General_Error_Trap
    Resume Exit_go_previous2_Click
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    Call General_Error_Trap
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    Call General_Error_Trap
    Resume Exit_go_last_Click
End Sub

Private Sub txtCrateNumber_AfterUpdate()
'update field that holds crate number and letter together
On Error GoTo err_txtCrateNum

    Update_GID
    
Exit Sub

err_txtCrateNum:
    Call General_Error_Trap
    Exit Sub
End Sub
