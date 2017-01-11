Option Compare Database   'Use database order for string comparisons
Option Explicit 'saj
Private Sub UpdateGID()
'this is called by this form to update the GID field if either the
'Unit, find letter or number fields are altered
'replaces calls to macro Bone.Update contact Unit and GID
' SAJ
On Error GoTo err_UpdateGID

 Me![GID] = [Forms]![Fauna_Bone_Contact]![Unit number] & "." & [Forms]![Fauna_Bone_Contact]![letter code] & [Forms]![Fauna_Bone_Contact]![find number]
    

Exit Sub

err_UpdateGID:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub button_goto_measurement_Click()
'altered season 2006 - saj
'check if any measurement records exist and if not ask user to create
'this button is new here, requested by Rebecca and approved by Rissa
On Error GoTo Err_button_goto_measurement_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim retVal, checknum, sql, getMeasurementSet
    stDocName = "Fauna_Bone_Measurements"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
        
    'new for season 2006 - see if the modification record exists if not create it - SAJ
    checknum = DLookup("[GID]", "[Fauna_Bone_Measurements]", "[GID] = '" & Me![GID] & "'")
    If IsNull(checknum) Then
        'gid not exist there yet
        retVal = MsgBox("No Measurement records exist yet for GID " & Me![GID] & ", do you want to create one now?", vbQuestion + vbYesNo, "Create New Modification Record")
        If retVal = vbNo Then
            'do nothing
            Exit Sub
        Else
            getMeasurementSet = InputBox("Please enter the measurement set number below:", "Measurement Set")
            If getMeasurementSet <> "" Then
                'so create it ready for data entry
                sql = "INSERT INTO [Fauna_Bone_Measurements] ([GID], [Unit number], [Letter code], [Find number], [Measurement Set]) VALUES ('" & Me![GID] & "'," & Me![Unit number] & ", '" & Me![letter code] & "', " & Me![find number] & ", " & getMeasurementSet & ");"
                DoCmd.RunSQL sql
            End If
        End If
    End If
        
        
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No Unit Number"
End If

Exit_button_goto_measurement_Click:
    Exit Sub

Err_button_goto_measurement_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_measurement_Click
End Sub

Private Sub Button23_Click()
'altered season 2006 - saj
'error trap and new way of creating new rec
On Error GoTo err_23

Dim tempGID
Dim tempUnit
Dim tempLetter
Dim tempFind
Dim temporder

tempGID = GID
tempUnit = Unit_number
tempLetter = Letter_code
tempFind = Find_number
temporder = ContactOrder

    'get the last record entered
        Dim mydb As DAO.Database, myrs As DAO.Recordset, lastrec
        Set mydb = CurrentDb()
        Set myrs = mydb.OpenRecordset("Select [ContactOrder] FROM [Fauna_Bone_Contact] WHERE [Unit Number] = " & Me![Unit number] & " AND Ucase([Letter Code]) = '" & Me![letter code] & "' AND [Find Number] = " & Me![find number] & " ORDER BY [Find number];", dbOpenSnapshot)
        If Not (myrs.BOF And myrs.EOF) Then
            myrs.MoveLast
            lastrec = myrs![ContactOrder]
        Else
            lastrec = ""
        End If
        myrs.Close
        Set myrs = Nothing
        mydb.Close
        Set mydb = Nothing

'DoCmd.DoMenuItem MenuBar:=acFormBar, MenuName:=3, Command:=0
DoCmd.RunCommand acCmdRecordsGoToNew

GID = tempGID
Unit_number = tempUnit
Letter_code = tempLetter
Find_number = tempFind
If lastrec = "" Then
    ContactOrder = temporder + 1
Else
    ContactOrder = lastrec + 1
End If
Exit Sub

err_23:
    Call General_Error_Trap
    Exit Sub
End Sub


Sub button_goto_artefacts_Click()
'season 2006, functionality modified slightly - SAJ
On Error GoTo Err_button_goto_artefacts_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Artifacts"
   
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    
    'saj - take over gid what ever the situation
    'If [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = 0 Then
        stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    'End If
        
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No GID Number"
End If

Exit_button_goto_artefacts_Click:
    Exit Sub

Err_button_goto_artefacts_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_artefacts_Click
    
End Sub
Sub button_goto_unitBFD_Click()
'season 2006, functionality modified slightly - SAJ

On Error GoTo Err_button_goto_unitBFD_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
If Me![GID] <> "" Then
    'new check for GID entered by saj
    'the form is only minimised so must save data manually here - saj
    DoCmd.RunCommand acCmdSaveRecord
    
    'why is this working on just unit number not GID?
    'If [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = 0 Then
    '    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    'End If
    stLinkCriteria = "[GID]='" & Me![GID] & "'"
    DoCmd.Minimize

    DoCmd.OpenForm stDocName, , , stLinkCriteria
Else
    MsgBox "Please enter the GID fields for this record or select a GID first", vbInformation, "No GID Number"
End If

Exit_button_goto_unitBFD_Click:
    Exit Sub

Err_button_goto_unitBFD_Click:
    Call General_Error_Trap
    Resume Exit_button_goto_unitBFD_Click
    
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    'If Me.Filter <> "" Then
    '        If Me.Filter <> "[GID] = '" & Me![cboFind].Column(1) & "'" Then
    '            MsgBox "This form was opened to only show the record " & Me.Filter & ". This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
    '            Me.FilterOn = False
    '            Me.Filter = ""
    '        End If
    '    End If
    
    Me.Filter = "[GID] = '" & Me![cboFind] & "' AND [ContactOrder] = " & Me![cboFind].Column(1)
    Me.FilterOn = True
    'DoCmd.GoToControl "GID"
    'DoCmd.FindRecord Me![cboFind]

End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry this value cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMenu_Click()
'new menu button
On Error GoTo err_cmdMenu

    DoCmd.OpenForm "Bone", acNormal
    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdMenu:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Field135_AfterUpdate()
'Type combo
'changed from macro Bone.update type
'new version 2.2 season 2006, make description 1 and 2 combos dependant on type selected here
On Error GoTo err135
Dim val

    If Me![Field135] <> "" Then
        val = CInt(Me![Field135])
        
        Select Case val
        Case 1
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type1_Desc1.code, Fauna_Code_Bone_Contact_Type1_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type1_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type1_Desc2.code, Fauna_Code_Bone_Contact_Type1_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type1_Desc2; "
        Case 2
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type2_Desc1.code, Fauna_Code_Bone_Contact_Type2_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type2_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type2_Desc2.code, Fauna_Code_Bone_Contact_Type2_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type2_Desc2; "
        Case 3
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type3_Desc1.code, Fauna_Code_Bone_Contact_Type3_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type3_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type3_Desc2.code, Fauna_Code_Bone_Contact_Type3_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type3_Desc2; "
        Case 4
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type4_Desc1.code, Fauna_Code_Bone_Contact_Type4_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type4_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type4_Desc2.code, Fauna_Code_Bone_Contact_Type4_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type4_Desc2; "
        Case 5
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type5_Desc1.code, Fauna_Code_Bone_Contact_Type5_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type5_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type5_Desc2.code, Fauna_Code_Bone_Contact_Type5_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type5_Desc2; "
        End Select
    End If


Exit Sub

err135:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Find_number_AfterUpdate()
' added here 2006
Call UpdateGID
End Sub

Private Sub Form_Current()
'Type combo
'changed from macro Bone.update type
'new version 2.2 season 2006, make description 1 and 2 combos dependant on type selected here
On Error GoTo err_current
Dim val

    If Me![Field135] <> "" Then
        val = CInt(Me![Field135])
        
        Select Case val
        Case 1
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type1_Desc1.code, Fauna_Code_Bone_Contact_Type1_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type1_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type1_Desc2.code, Fauna_Code_Bone_Contact_Type1_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type1_Desc2; "
        Case 2
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type2_Desc1.code, Fauna_Code_Bone_Contact_Type2_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type2_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type2_Desc2.code, Fauna_Code_Bone_Contact_Type2_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type2_Desc2; "
        Case 3
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type3_Desc1.code, Fauna_Code_Bone_Contact_Type3_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type3_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type3_Desc2.code, Fauna_Code_Bone_Contact_Type3_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type3_Desc2; "
        Case 4
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type4_Desc1.code, Fauna_Code_Bone_Contact_Type4_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type4_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type4_Desc2.code, Fauna_Code_Bone_Contact_Type4_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type4_Desc2; "
        Case 5
            Me![Description 1].RowSource = "SELECT Fauna_Code_Bone_Contact_Type5_Desc1.code, Fauna_Code_Bone_Contact_Type5_Desc1.[text equivalent] FROM Fauna_Code_Bone_Contact_Type5_Desc1; "
            Me![Description 2].RowSource = "SELECT Fauna_Code_Bone_Contact_Type5_Desc2.code, Fauna_Code_Bone_Contact_Type5_Desc2.[text equivalent] FROM Fauna_Code_Bone_Contact_Type5_Desc2; "
        End Select
    End If


Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Letter_code_AfterUpdate()
' added here 2006
Call UpdateGID
End Sub

Private Sub Unit_number_AfterUpdate()
' originally on got focus was called macro: Bone.Update contact Unit and GID
' this translated into code and call moved to here
Call UpdateGID

End Sub

