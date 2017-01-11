Option Compare Database
Option Explicit 'saj
Private Sub FindFacility(what)
'idea copied from crates register and utilised here, kept basic
'saj season 2008, v3.3
On Error GoTo Err_find


    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim message As String, title As String, Unit As String, default As String
    Dim Material As String, descrip As String
    Dim un, lett, num
    
    If what = "unit" Then
        message = "Enter a unit number"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = "1000"   ' Set default.
        ' Display message, title, and default value.
        Unit = InputBox(message, title, default)
        If Unit = "" Then Exit Sub 'saj catch no entry
        'stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
        'saj - jules wants to just find numbers directly
        stLinkCriteria = "[Unit Number] =" & Unit
    ElseIf what = "material" Then
        message = "Enter a material"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        Material = InputBox(message, title, default)
        If Material = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Material Group] like '*" & Material & "*'"
    ElseIf what = "descrip" Then
        message = "Enter a description"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        descrip = InputBox(message, title, default)
        If descrip = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Description] like '*" & descrip & "*'"
    ElseIf what = "subgroup" Then
        'message = "Enter a subgroup"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        message = "Enter a Material Group"   ' Set prompt.
        un = InputBox(message, title, default)
        If un = "" Then Exit Sub 'saj catch no entry
        message = "Enter a material subgroup"   ' Set prompt.
        lett = InputBox(message, title, default)
        If lett = "" Then Exit Sub 'saj catch no entry
        'message = "Enter a number"   ' Set prompt.
        'num = InputBox(message, title, default)
        'If num = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Material Group] ='" & un & "' AND [Material Subgroup] ='" & lett & "'"
    ElseIf what = "object" Then
        message = "Enter object type"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        Material = InputBox(message, title, default)
        If Material = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Object Type] = '" & Material & "'"
    ElseIf what = "all" Then
        'message = "Enter a subgroup"   ' Set prompt.
        title = "Searching Finds Register" ' Set title.
        default = ""   ' Set default.
        ' Display message, title, and default value.
        message = "Enter a Material Group"   ' Set prompt.
        un = InputBox(message, title, default)
        If un = "" Then Exit Sub 'saj catch no entry
        message = "Enter a material subgroup"   ' Set prompt.
        lett = InputBox(message, title, default)
        If lett = "" Then Exit Sub 'saj catch no entry
        message = "Enter an object type"   ' Set prompt.
        num = InputBox(message, title, default)
        If num = "" Then Exit Sub 'saj catch no entry
        stLinkCriteria = "[Material Group] ='" & un & "' AND [Material Subgroup] ='" & lett & "' AND [Object Type] = '" & num & "'"
    Else
        Exit Sub
    End If
    stDocName = "frm_pop_search_finds:BasicData"
    'stLinkCriteria = "[Unit Number] like '*" & Unit & "*'"
    'DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria, acFormReadOnly
    
Exit_find:
    Exit Sub

Err_find:
    MsgBox Err.Description
    Resume Exit_find
End Sub

Private Sub Update_GID()
'sub used by gid fields written by anja adapted by saj to error trap and include letter code fld
On Error GoTo err_updategid

'Me![GID] = Me![Unit] & "." & Me![Find Number]

Me![GID] = Me![txtUnit] & "." & Me![cboFindLetter] & Me![txtFindNumber]
If Me![txtUnit] <> "" And Me![cboFindLetter] <> "" And Me![txtFindNumber] <> "" Then
    Me.Refresh
End If
Exit Sub

err_updategid:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindLetter_AfterUpdate()
'new season 2006 - saj
On Error GoTo err_cbofindletter

    Update_GID
    'Forms![Finds: Basic Data].Refresh

Exit Sub

err_cbofindletter:
    Call General_Error_Trap
    Exit Sub
    
End Sub



Private Sub cboFindUnit_AfterUpdate()
'********************************************
'Find the selected gid from the list
'********************************************
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'for existing number the field will be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![GID].Enabled = False Then Me![txtUnit].Enabled = True
        DoCmd.GoToControl "GID"
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

Private Sub Close_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()

On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    'new record allow GID entry
    Me![txtUnit].Enabled = True
    Me![txtUnit].Locked = False
    Me![txtUnit].BackColor = 16777215
    Me![cboFindLetter].Enabled = True
    Me![cboFindLetter].Locked = False
    Me![cboFindLetter].BackColor = 16777215
    Me![txtFindNumber].Enabled = True
    Me![txtFindNumber].Locked = False
    Me![txtFindNumber].BackColor = 16777215
    
    DoCmd.GoToControl "txtUnit"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
'new version 4.4 - 2009
On Error GoTo err_delete

    Dim retVal
    retVal = MsgBox("This action will delete the record for find " & Me![txtUnit] & "." & Me![cboFindLetter] & Me![txtFindNumber] & ", are you really sure you want to delete this record?", vbCritical + vbYesNo, "Confirm Deletion")
    If retVal = vbYes Then
        'ok proceed
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace, sql1, sql2, myq As QueryDef
        Dim myrs As DAO.Recordset
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
        
        sql1 = "DELETE FROM [Finds: Basic Data] WHERE [Unit] = " & Me!txtUnit & " AND [FindLetter] = '" & Me!cboFindLetter & "' AND [FindNumber] = " & Me!txtFindNumber & ";"
        Set myq = mydb.CreateQueryDef("")
        myq.sql = sql1
        myq.Execute
    
        sql2 = "DELETE FROM [Finds_Basic_Data_Materials_and_Type] WHERE [Unit] = " & Me!txtUnit & " AND [FindLetter] = '" & Me!cboFindLetter & "' AND [FindNumber] = " & Me!txtFindNumber & ";"
        Set myq = mydb.CreateQueryDef("")
        myq.sql = sql2
        myq.Execute
        
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Record Deleted"
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

Private Sub cmdEdit_Click()
'new 2008, allow GID edit in finds:basic data and Finds_Basic_Data_Materials_and_Type
On Error GoTo err_cmdEdit
    
    If Me![txtUnit] <> "" Then
        Dim getUnit, getLetter, getNum
        getUnit = InputBox("Please edit the Unit number:", "Unit", Me![txtUnit])
        If getUnit = "" Then Exit Sub
        getLetter = InputBox("Please edit the Letter:", "Letter Code", Me![cboFindLetter])
        If getLetter = "" Then Exit Sub
        getNum = InputBox("Please edit the Find number:", "Find Number", Me![txtFindNumber])
        If getNum = "" Then Exit Sub
        
        'ok must check if new number already exists
        Dim checkit, sql
        checkit = DLookup("Unit", "[finds: basic data]", "Unit = " & getUnit & " AND FindLetter = '" & getLetter & "' AND FindNumber = " & getNum)
        If IsNull(checkit) Then
            'ok can make the change, now check if the GID is in Finds_Basic_Data_Materials_and_Type and change there first
            checkit = DLookup("unit", "Finds_Basic_Data_Materials_and_Type", "Unit = " & Me![txtUnit] & " AND FindLetter = '" & Me![cboFindLetter] & "' AND FindNumber = " & Me![txtFindNumber])
            If Not IsNull(checkit) Then
                sql = "UPDATE [Finds_Basic_Data_Materials_and_Type] SET Unit = " & getUnit & ", FindLetter = '" & getLetter & "', FindNumber = " & getNum & " WHERE Unit = " & Me![txtUnit] & " AND FindLetter = '" & Me![cboFindLetter] & "' AND FindNumber = " & Me![txtFindNumber] & ";"
                DoCmd.RunSQL sql
            End If
            Me![txtUnit] = getUnit
            Me![cboFindLetter] = getLetter
            Me![txtFindNumber] = getNum
            Me![GID] = getUnit & "." & getLetter & getNum
            MsgBox "GID changed successfully", vbInformation, "Operation Complete"
        Else
            MsgBox "Sorry but this GID exists in the database already, you cannot make this change. Use the find facility to view the record with this GID.", vbInformation, "Key Violation"
        End If
    End If

Exit Sub

err_cmdEdit:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdfindlocate_Click()
On Error GoTo err_cmdfindlocate
'new 2015, allow assigning crate number to a x-find - DL

Dim sql As String

'[Store: Units in Crates subform]
[Store: Units in Crates subform].Form.AllowEdits = True
[Store: Units in Crates subform].Form![Crate Letter].Enabled = True
[Store: Units in Crates subform].Form![Crate Letter].Locked = False
[Store: Units in Crates subform].Form![Crate Number].Enabled = True
[Store: Units in Crates subform].Form![Crate Number].Locked = False
[Store: Units in Crates subform].Form![Material].Locked = False
[Store: Units in Crates subform].Form![Material].Enabled = True
[Store: Units in Crates subform].Form![Description].Locked = False
[Store: Units in Crates subform].Form![Description].Enabled = True

sql = "INSERT INTO [Store: Units in Crates] ([Unit number], [FindSampleLetter], [FindNumber]) VALUES (" & Me![txtUnit] & ", '" & Me![cboFindLetter] & "', " & Me![txtFindNumber] & ");"
DoCmd.RunSQL sql
Debug.Print sql
Me.Refresh
Me![Store: Units in Crates subform].Form![Crate Letter].SetFocus


'    If Me![Crate Number] <> "" Then
'    Else
'        DoCmd.OpenForm "Store: Crate Register", acNormal, , "[Crate Number] = '" & Me![Store: Units in Crates subform].Form![Crate Number] & "'"
'    End If
Exit Sub

err_cmdfindlocate:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdMaterial_Click()
Call FindFacility("material")
End Sub

Private Sub Command66_Click()
Call FindFacility("subgroup")
End Sub

Private Sub Command67_Click()
Call FindFacility("object")
End Sub

Private Sub Command68_Click()
Call FindFacility("all")
End Sub

Private Sub Command69_Click()
Call FindFacility("descrip")
End Sub



Private Sub Form_BeforeInsert(Cancel As Integer)
'new for v4.2 - 2009
On Error GoTo err_Form_AfterUpdate
Me![LastUpdated] = Now()
Exit Sub

err_Form_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'on error goto err_open

If GetGeneralPermissions = "Admin" Then
    Me![cmdDelete].Visible = True
    Me![cmdEdit].Visible = True
Else
    Me![cmdDelete].Visible = False
    Me![cmdEdit].Visible = False
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

Private Sub gotocrate_Click()
On Error GoTo err_gotocrate_Click

DoCmd.OpenForm "Store: Crate Register", , , "CrateLetter = '" & [Store: Units in Crates subform].Form![Crate Letter] & "' AND CrateNumber = " & [Store: Units in Crates subform].Form![Crate Number]
Forms![Store: Crate Register].SetFocus
Exit Sub

err_gotocrate_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtFindNumber_AfterUpdate()
'find num call to Update GID removed from On enter and On change events and just left here
'plus error trap intro - season 2006 - saj
On Error GoTo err_txtfindnumber

    Update_GID
    'Forms![Finds: Basic Data].Refresh

Exit Sub

err_txtfindnumber:
    Call General_Error_Trap
    Exit Sub
    
End Sub


Private Sub Form_Current()
'disabled SAJ
'If Me![Conservation Ref] <> nil Then
' Me![conservation].Enabled = True
' Else
' Me![conservation].Enabled = False
'End If
 
'new code for 2006
On Error GoTo err_current

'add possibility to locate a find directly into a crate

    Dim strSQL As String
    Dim rst As DAO.Recordset
    
    If Not IsNull(Me![txtUnit]) Then
    'maybe findletter and number can be Null: Or IsNull(Me![cboFindLetter]) Or IsNull(Me![txtFindNumber])
    strSQL = "SELECT [CrateLetter], [CrateNumber], [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber] FROM [Store: Units in Crates] WHERE [Unit number]=" & Me![txtUnit] & " AND [FindSampleLetter]='" & Me![cboFindLetter] & "' AND [FindNumber]=" & Me![txtFindNumber] & " ORDER BY [CrateLetter], [CrateNumber] ASC;"
                Debug.Print strSQL
                Set rst = CurrentDb.OpenRecordset(strSQL)
                
                If Not (rst.EOF) Then
                    Debug.Print rst![CrateLetter] & " " & rst![CrateNumber]
                    Me![cmdfindlocate].Visible = False
                    Me![gotocrate].Visible = True
                Else
                    Debug.Print "X-Find does not exist"
                    Me![cmdfindlocate].Visible = True
                    Me![gotocrate].Visible = False
                End If
    Else
        Me![cmdfindlocate].Visible = True
        Me![gotocrate].Visible = False
    End If
    'If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![cboFindLetter] = "" Or IsNull(Me![cboFindLetter])) And (Me![txtFindNumber] = "" Or IsNull(Me![txtFindNumber])) Then
    'don't include find number as defaults to x
    If (Me![txtUnit] = "" Or IsNull(Me![txtUnit])) And (Me![txtFindNumber] = "" Or IsNull(Me![txtFindNumber])) Then
        'new record allow GID entry
        Me![txtUnit].Enabled = True
        Me![txtUnit].Locked = False
        Me![txtUnit].BackColor = 16777215
        Me![cboFindLetter].Enabled = True
        Me![cboFindLetter].Locked = False
        Me![cboFindLetter].BackColor = 16777215
        Me![txtFindNumber].Enabled = True
        Me![txtFindNumber].Locked = False
        Me![txtFindNumber].BackColor = 16777215
    Else
        'existing entry lock
        Me![txtUnit].Enabled = False
        Me![txtUnit].Locked = True
        Me![txtUnit].BackColor = Me.Section(0).BackColor
        Me![cboFindLetter].Enabled = False
        Me![cboFindLetter].Locked = True
        Me![cboFindLetter].BackColor = Me.Section(0).BackColor
        Me![txtFindNumber].Enabled = False
        Me![txtFindNumber].Locked = True
        Me![txtFindNumber].BackColor = Me.Section(0).BackColor
    End If
    
Exit Sub

'Me![frm_subform_materialstypes].Requery
'Me![frm_subform_materialstypes].Form![cboMaterialSubGroup].Requery

    
    
err_current:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_AfterUpdate()
'Unit call to Update GID removed from On enter and On change events and just left here
'plus error trap intro - season 2006 - saj
On Error GoTo err_unit

Update_GID
'don't need
'Forms![Finds: Basic Data].Refresh

Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub







Sub first_Click()
On Error GoTo Err_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_first_Click:
    Exit Sub

Err_first_Click:
    MsgBox Err.Description
    Resume Exit_first_Click
    
End Sub
Sub prev_Click()
On Error GoTo Err_prev_Click


    DoCmd.GoToRecord , , acPrevious

Exit_prev_Click:
    Exit Sub

Err_prev_Click:
    MsgBox Err.Description
    Resume Exit_prev_Click
    
End Sub
Sub next_Click()
On Error GoTo Err_next_Click


    DoCmd.GoToRecord , , acNext

Exit_next_Click:
    Exit Sub

Err_next_Click:
    MsgBox Err.Description
    Resume Exit_next_Click
    
End Sub
Sub last_Click()
On Error GoTo Err_last_Click


    DoCmd.GoToRecord , , acLast

Exit_last_Click:
    Exit Sub

Err_last_Click:
    MsgBox Err.Description
    Resume Exit_last_Click
    
End Sub
Sub new_Click()
On Error GoTo Err_new_Click


    DoCmd.GoToRecord , , acNewRec

Exit_new_Click:
    Exit Sub

Err_new_Click:
    MsgBox Err.Description
    Resume Exit_new_Click
    
End Sub
Sub closeCommand45_Click()
On Error GoTo Err_closeCommand45_Click


    DoCmd.Close

Exit_closeCommand45_Click:
    Exit Sub

Err_closeCommand45_Click:
    MsgBox Err.Description
    Resume Exit_closeCommand45_Click
    
End Sub
Sub find_Click()
On Error GoTo Err_find_Click


    Screen.PreviousControl.SetFocus
    GID.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_Click:
    Exit Sub

Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
    
End Sub
Sub cons_Click()
On Error GoTo Err_cons_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[Conserv: Basic Record.GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cons_Click:
    Exit Sub

Err_cons_Click:
    MsgBox Err.Description
    Resume Exit_cons_Click
    
End Sub
Sub conservation_Click()
On Error GoTo Err_conservation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Conserv: Basic Record"
    
    stLinkCriteria = "[GID1]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_conservation_Click:
    Exit Sub

Err_conservation_Click:
    MsgBox Err.Description
    Resume Exit_conservation_Click
    
End Sub


Private Sub txtUnit_AfterUpdate()
'new season 2006 - saj
On Error GoTo err_txtUnit

    Update_GID
    'Forms![Finds: Basic Data].Refresh
    

Exit Sub

err_txtUnit:
    Call General_Error_Trap
    Exit Sub
End Sub
