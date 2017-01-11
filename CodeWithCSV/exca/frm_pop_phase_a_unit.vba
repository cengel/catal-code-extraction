Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
'close with no action
On Error GoTo err_cancel

    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_cancel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOK_Click()
'NEW 2009 season ok take the phase number and fill out field
On Error GoTo err_cmdOK
    
    'first round implementation when field was simply in Unit Sheet table (assumption unit can only
    'be phased to 1 space or building.
    'If Me![cboSelect] <> "" Then
    '    Forms![Exca: Unit Sheet]![Phase] = Me!cboSelect
    'End If
    
    'second round - a unit can be phased to more than one space - but only once within that space
    'must take unit number from unit sheet
    Dim Unit, getBuildingorSpace, getDivider
    If Me![cboSelect] <> "" Then
        Unit = Forms![Exca: Unit Sheet]![Unit Number]
        getDivider = InStr(Me!cboSelect, ".") 'format is B42.A or Sp115.1 etc etc
        getBuildingorSpace = Left(Me!cboSelect, getDivider - 1)
    
        'is this a new entry for this unit or an overwrite of an existing entry for this space or building
        Dim checkRec, sql
        checkRec = DLookup("OccupationPhase", "[Exca: Units Occupation Phase]", "[Unit] = " & Unit & " AND [OccupationPhase] like '" & getBuildingorSpace & "%'")
        If IsNull(checkRec) Then
            'no phasing yet for this building or space so simply add
            sql = "INSERT INTO [Exca: Units Occupation Phase] ([Unit], [OccupationPhase]) VALUES (" & Unit & ",'" & Me!cboSelect & "');"
            DoCmd.RunSQL sql
        Else
            'it exists so must update it
            sql = "UPDATE [Exca: Units Occupation Phase] SET [OccupationPhase] = '" & Me![cboSelect] & "' WHERE Unit = " & Unit & " AND [OccupationPhase] = '" & checkRec & "';"
            DoCmd.RunSQL sql
        End If
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "You must select a phase from the list or press cancel to leave this form", vbInformation, "No Phase Selected"
        
    End If
    
    
Exit Sub

err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRemove_Click()
'remove phase number for this unit
On Error GoTo err_cmdRemove
    'first round implementation when field was simply in Unit Sheet table (assumption unit can only
    'Forms![Exca: Unit Sheet]![Phase] = ""
    
    'second round - a unit can be phased to more than one space - but only once within that space
    'must take unit number from unit sheet
    Dim Unit, getEquals, Phase, getNumber, sql, resp

    Unit = Forms![Exca: Unit Sheet]![Unit Number]
 resp = MsgBox("This will remove all the phasing associated with Unit " & Unit & " - ARE YOU SURE?" & Chr(13) & Chr(13) & "To remove one phase item only: on the main unit sheet click over the arrow to the right of the specific phase and press delete.", vbCritical + vbYesNo, "Confirm Action")
 If resp = vbYes Then
    getEquals = InStr(Me!cboSelect.RowSource, "=") 'format is =Sp115.1 or B42. etc etc
    'getNumber = right(Me!cboSelect.RowSource, Len(Me!cboSelect.RowSource) - (getEquals - 1))
    getNumber = Mid(Me!cboSelect.RowSource, getEquals + 1, (Len(Me!cboSelect.RowSource) - 1) - getEquals)
    If InStr(Me!cboSelect.RowSource, "Space") > 0 Then
        'its space
        Phase = "Sp" & getNumber & "."
    Else
        'its building
        Phase = "B" & getNumber & "."
        
    End If
    
    'delete phasing for this space or building
    ''sql = "DELETE FROM [Exca: Units Occupation Phase] WHERE Unit = " & Unit & " AND [OccupationPhase] like '" & Phase & "%';"
    ''2010 - this code would only work where the phase has been put in exactly in the correct format of eg: sp1004. but this doesn't always
    ''happen as not enfored = any not in this format eg: just in as '5' will not be deleted. This is giving in correct impression of functionality
    ''simply taking this SQL down to unit number which is the same effect
    sql = "DELETE FROM [Exca: Units Occupation Phase] WHERE Unit = " & Unit & ";"
    DoCmd.RunSQL sql
       
End If
    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdRemove:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'this form is new 2009 to deal with phasing - it must present a list of phases relevant to
'the units space or building
On Error GoTo err_open

    'get the recordsource from the combo from the openargs
    If Not IsNull(Me.OpenArgs) Then
        Me!cboSelect.RowSource = Me.OpenArgs
        Me!cboSelect.Requery
        
        Dim Unit, getEquals, getNumber, Phase, phasedalready, sql
        Unit = Forms![Exca: Unit Sheet]![Unit Number]
        getEquals = InStr(Me!cboSelect.RowSource, "=") 'format is =Sp115.1 or B42. etc etc
        'getNumber = right(Me!cboSelect.RowSource, Len(Me!cboSelect.RowSource) - (getEquals - 1))
        getNumber = Mid(Me!cboSelect.RowSource, getEquals + 1, (Len(Me!cboSelect.RowSource) - 1) - getEquals)
        If InStr(Me!cboSelect.RowSource, "Space") > 0 Then
            'its space
            Phase = "Sp" & getNumber & "."
        Else
            'its building
            Phase = "B" & getNumber & "."
        End If
        
        phasedalready = DCount("[OccupationPhase]", "[Exca: Units Occupation Phase]", "[OccupationPhase] like '" & Phase & "%'")
        Me!cmdRemove.Caption = "Remove Unit from Phasing of " & Phase
        If phasedalready >= 1 Then
            Me!cmdRemove.Enabled = True
            
        Else
            Me!cmdRemove.Enabled = False
        End If
    Else
        MsgBox "Form opened with no parametres. Invalid action. The form will now close.", vbInformation, "No OpenArgs"
        DoCmd.Close acForm, Me.Name
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
