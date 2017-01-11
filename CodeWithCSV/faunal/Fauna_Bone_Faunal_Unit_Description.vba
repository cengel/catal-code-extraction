Option Compare Database   'Use database order for string comparisons
Option Explicit 'added saj

Sub Button_Goto_BFD_Click()
'adapted season 2006 to error trap no unit number- SAJ
On Error GoTo Err_Button_Goto_BFD_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

If Me![Unit number] <> "" Then
    'check intro by SAJ
    stDocName = "Fauna_Bone_Basic_Faunal_Data"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    
    ' MR July 18 2005
    'DoCmd.Save 'SAJ comment out as throwing error  29068, Microsoft Access cannot complete this operation. You must stop the code and try again.
    DoCmd.RunCommand acCmdSaveRecord
    
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria

    'If [Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = 0 Then ' MR July 17 2005
    ''[Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number] = [Forms]![Fauna_Bone_Faunal_Unit_Description].[Unit number]
    'End If ' MR July 17 2005
    'SAJ comment out line 2 above, replace with below
    If IsNull([Forms]![Fauna_Bone_Basic_Faunal_Data].[Unit number]) Then
        Forms![Fauna_Bone_Basic_Faunal_Data].[Unit number] = Forms![Fauna_Bone_Faunal_Unit_Description].[Unit number]
    End If
Else
    MsgBox "Please enter or select a Unit Number first", vbInformation, "No Unit Number"
End If

Exit_Button_Goto_BFD_Click:
    Exit Sub

Err_Button_Goto_BFD_Click:
    If Err.Number = 2046 And Me.Dirty = False Then
        'save record has failed  - DB might be Read only, added dirty check to make sure nothing added
        Resume Next
    Else
        Call General_Error_Trap
        Resume Exit_Button_Goto_BFD_Click
    End If
End Sub
Sub Command25_Click()
On Error GoTo Err_Command25_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

Exit_Command25_Click:
    Exit Sub

Err_Command25_Click:
    MsgBox Err.Description
    Resume Exit_Command25_Click
    
End Sub

Private Sub Button67_Click()
'replace macro bone.bone button
On Error GoTo err_but67

    DoCmd.OpenForm "Bone", acNormal
    DoCmd.Close acForm, Me.Name
Exit Sub

err_but67:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_AfterUpdate()
'new find combo by SAJ - NR remove filter msg 5/7/06
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    If Me.Filter <> "" Then
        If Me.Filter <> "[Unit Number] = '" & Me![cboFind] & "'" Then
    '        MsgBox "This form was opened to only show a particular Unit. This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
            Me.FilterOn = False
        End If
    End If

    DoCmd.GoToControl "Unit Number"
    DoCmd.FindRecord Me![cboFind]

End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "Sorry there is no FUD for this Unit", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdMenu_Click()
'replace macro bone.bone button on button67
On Error GoTo err_but67

    DoCmd.OpenForm "Bone", acNormal
    DoCmd.Close acForm, Me.Name
Exit Sub

err_but67:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Activate()

On Error GoTo err_activate
    'request by NR 5/7/06 so when revisited is uptodate (mainly last F# entered)
    Me.Requery
    'saj 2008 - faunal wishlist, ensure focus is in cboFind everytime opened so unit number not overwritten
    DoCmd.GoToControl "cboFind"
Exit Sub

err_activate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'new season 2006, saj
'show number of basic records on screen for each record
On Error GoTo err_current
Dim recCount, lastrec, pieces, elements

'new 2009 lock unit number field if not a new entry
If Me![Unit number] <> "" Then
    Me![Unit number].Locked = True
    'Me![Unit number].BackColor = 16763904 'kathy not like colour
    Me![Unit number].BackColor = 12632256 'darker blue at top is too dark here
Else
    Me![Unit number].Locked = False
    Me![Unit number].BackColor = 16777215
End If

If Me![Unit number] <> "" Then
    recCount = DCount("[Unit Number]", "Fauna_Bone_Basic_Faunal_Data", "[Unit Number] = " & Me![Unit number])
    Me![txtCount] = recCount
    
    'v2.2 add in counts of elements and pieces
    pieces = DLookup("[TotalPieces]", "Q_Total_Pieces_and_Elements_Per_Unit", "[Unit Number] = " & Me![Unit number])
    Me![txtPieces] = pieces
    elements = DLookup("[TotalElements]", "Q_Total_Pieces_and_Elements_Per_Unit", "[Unit Number] = " & Me![Unit number])
    Me![txtElements] = elements
    
    If recCount > 0 Then
        'get the last record entered
        Dim mydb As DAO.Database, myrs As DAO.Recordset
        Set mydb = CurrentDb()
        Set myrs = mydb.OpenRecordset("Select [GID] FROM [Fauna_Bone_Basic_Faunal_Data] WHERE [Unit Number] = " & Me![Unit number] & " AND Ucase([Letter Code]) = 'F' ORDER BY [Find number];", dbOpenSnapshot)
        If Not (myrs.BOF And myrs.EOF) Then
            myrs.MoveLast
            lastrec = myrs![GID]
        Else
            lastrec = "No F Numbers"
        End If
        myrs.Close
        Set myrs = Nothing
        mydb.Close
        Set mydb = Nothing
    
        Me![txtLast] = lastrec
    Else
         Me![txtLast] = "No F Numbers"
    End If
End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub

End Sub



Private Sub Form_Open(Cancel As Integer)
'NEW 2009 - show hide post ex fields depending on permissions
On Error GoTo err_open

If GetGeneralPermissions = "Admin" Then
    Me![cboWorkingPhase].Visible = True
    Me![cboConsumptionContext].Visible = True
    Me![cboDepositionalContext].Visible = True
Else
    Me![cboWorkingPhase].Visible = False
    Me![cboConsumptionContext].Visible = False
    Me![cboDepositionalContext].Visible = False
End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub open_short_Click()
'altered season 2006 - capture no unit number - SAJ
On Error GoTo Err_open_short_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
If Me![Unit number] <> "" Then
    'check intro by SAJ
    stDocName = "Bone: Short Faunal Data"
    
    stLinkCriteria = "[Unit number]=" & Me![Unit number]
    
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    'saj comment out and replace
    'If [Forms]![Bone: Short Faunal Data].[Unit number] = 0 Then
    '[Forms]![Bone: Short Faunal Data].[Unit number] = [Forms]![Fauna_Bone_Faunal_Unit_Description].[Unit number]
    'End If
    If IsNull([Forms]![Bone: Short Faunal Data].[Unit number]) Then
        [Forms]![Bone: Short Faunal Data].[Unit number] = Forms![Fauna_Bone_Faunal_Unit_Description].[Unit number]
    End If
Else
    MsgBox "Please enter or select a Unit Number first", vbInformation, "No Unit Number"
End If
Exit_open_short_Click:
    Exit Sub

Err_open_short_Click:
    MsgBox Err.Description
    Resume Exit_open_short_Click
    
End Sub
Sub New_record_Click()
On Error GoTo Err_New_record_Click


    DoCmd.GoToRecord , , acNewRec
    DoCmd.GoToControl "Unit Number" 'added by saj
    Me![txtCount] = ""
    Me![txtLast] = ""
    '2008 wishlist saj
    Me![txtPieces] = ""
    Me![txtElements] = ""

Exit_New_record_Click:
    Exit Sub

Err_New_record_Click:
    MsgBox Err.Description
    Resume Exit_New_record_Click
    
End Sub

Private Sub Unit_number_AfterUpdate()
'check existence of unit number - new 2008 wishlist - saj
On Error GoTo err_unit

    If IsNull(Me![Unit number].OldValue) Then
        Dim checknum, unit
        checknum = DLookup("[Unit number]", "[Fauna_Bone_Faunal_Unit_Description]", "[Unit number] = " & Me![Unit number])
        If Not IsNull(checknum) Then
            'exists
            MsgBox "This unit number exists already, the system will take you to the record", vbInformation, "Duplicate Unit Number"
            unit = Me![Unit number]
            'Me![txtBag] = ""
            Me.Undo
            DoCmd.GoToControl Me![Unit number].Name
            DoCmd.FindRecord unit
        End If
    End If

Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub
