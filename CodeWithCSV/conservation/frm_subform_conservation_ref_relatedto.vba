Option Compare Database
Option Explicit

Private Sub cmdRange_Click()
'new season 2006, saj
'allow user to enter a range of x number easily for the ame unit
On Error GoTo err_cmdRange
Dim retVal, assoUnit, assoLetter, assoStartNum, assoEndNum, sql
    If Me![txtExcavationIDNumber] <> "" Then
        retVal = MsgBox("Enter the object number range for Unit " & Me![txtExcavationIDNumber] & "?", vbYesNo + vbQuestion, "Unit to associate range to")
        If retVal = vbYes Then
            assoUnit = Me![txtExcavationIDNumber]
        Else
            assoUnit = InputBox("Please enter the Unit number associated with all the objects:", "Unit Number")
        End If
        
        If assoUnit = "" Then
            MsgBox "Operation Cancelled"
        Else
            assoLetter = InputBox("Please enter the finds letter associate with all the objects:", "Finds Letter", "X")
            If assoLetter = "" Then
                MsgBox "Operation Cancelled"
            Else
                'assoStartNum, assoEndNum
                assoStartNum = InputBox("Please enter the first number of the range:", "Start of number range")
                If assoStartNum = "" Then
                    MsgBox "Operation Cancelled"
                Else
                    assoEndNum = InputBox("Please enter the last number of the range:", "Start of number range")
                    If assoEndNum = "" Then
                        MsgBox "Operation Cancelled, both a start and end number are required."
                    Else
                        If CInt(assoStartNum) > CInt(assoEndNum) Then
                            MsgBox "Sorry but the start number is greater than the end number , invalid entry, please try again"
                        Else
                            Do Until CInt(assoStartNum) > CInt(assoEndNum)
                                sql = "INSERT INTO [Conservation_ConservRef_RelatedTo] ([ConservationRef_Year], [ConservationRef_ID], [RelatedToID],[RelatedToSubTypeID], [ExcavationIDNumber], [FindLetter], [FindSampleNumber])"
                                sql = sql & " VALUES ('" & Forms![Conserv: Basic Record]![txtConservationRef_Year] & "', " & Forms![Conserv: Basic Record]![txtConservationRef_ID] & "," & Forms![Conserv: Basic Record]![RelatedToID]
                                sql = sql & ", " & Me![cboRelatedToSubTypeID] & "," & assoUnit & ", '" & assoLetter & "'," & assoStartNum & ");"
                                DoCmd.RunSQL sql
                                assoStartNum = assoStartNum + 1
                            Loop
                            Me.Requery
                        End If
                    End If
                End If
            End If
        End If
    End If

Exit Sub

err_cmdRange:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdView_Click()
'set up which form to popup
'saj
On Error GoTo err_view

    If Me![RelatedToID] = 1 Then
        'building, space, feature, unit
        If Me![RelatedToSubTypeID] = 1 Then
            DoCmd.OpenForm "frm_subform_ExcaBuilding", acNormal, , "[Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        ElseIf Me![RelatedToSubTypeID] = 2 Then
            DoCmd.OpenForm "frm_subform_ExcaSpace", acNormal, , "[Space Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        ElseIf Me![RelatedToSubTypeID] = 3 Then
            DoCmd.OpenForm "frm_subform_ExcaFeature", acNormal, , "[Feature Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        ElseIf Me![RelatedToSubTypeID] = 4 Then
            DoCmd.OpenForm "frm_subform_ExcaUnit", acNormal, , "[Unit Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        Else
            MsgBox "Sorry option not recognised", vbExclamation, "Unknown Selection"
        End If
    ElseIf Me![RelatedToID] = 2 Then
        'object
        'temporarily we will use unit details here
        'DoCmd.OpenForm "frm_subform_ExcaUnit", acNormal, , "[Unit Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        'however could be faunal etc and no link at present so open unit info = 2009 v5.1
        If LCase(Me![cboFindLetter]) <> "x" Then
            DoCmd.OpenForm "frm_subform_ExcaUnit", acNormal, , "[Unit Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
        Else
            DoCmd.OpenForm "frm_subform_materialstypes", acNormal, , "[GID] = '" & Me![txtExcavationIDNumber] & "." & Me![cboFindLetter] & Me![txtFindSampleNumber] & "'", acFormReadOnly, acDialog
        End If
    ElseIf Me![RelatedToID] = 3 Then
        'sample
        'temporarily we will use unit details here
        DoCmd.OpenForm "frm_subform_ExcaUnit", acNormal, , "[Unit Number] = " & Me![txtExcavationIDNumber], acFormReadOnly, acDialog
    ElseIf Me![RelatedToID] = 4 Then
        'other

    End If
Exit Sub

err_view:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Form_Current()
'Alter fields to view dependant on type of data to enter
On Error GoTo err_current


    Me![txtComment].Enabled = True
    Me![lblRelatedTo].Caption = ""
    Me![lblExIDNumber].Caption = ""
    Me![lblFindSampleNumber].Caption = ""
    Me![lblFindLetter].Caption = ""
    Me![lblComment].Caption = "Comment"
    Me![cmdView].Visible = True
    
    Me![cboRelatedToSubTypeID].RowSource = "SELECT Conservation_Code_ConservRef_RelatedToSubType.RelatedToSubTypeID, Conservation_Code_ConservRef_RelatedToSubType.RelatedToTypeID, Conservation_Code_ConservRef_RelatedToSubType.RelatedToSubTypeText FROM Conservation_Code_ConservRef_RelatedToSubType WHERE RelatedToTypeID = " & Forms![Conserv: Basic Record]![RelatedToID] & ";"
    Me![cboRelatedToSubTypeID].Requery
    
    Me![cboFindLetter].RowSource = "SELECT [Exca: Letter codes].Letter, [Exca: Letter codes].ShortDescription FROM [Exca: Letter codes] WHERE ((([Exca: Letter codes].Letter)<>'E' And ([Exca: Letter codes].Letter)<>'G' And ([Exca: Letter codes].Letter)<>'T' And ([Exca: Letter codes].Letter)<>'W'));"
    
    'If Me![RelatedToID] = 1 Then
    If Forms![Conserv: Basic Record]![RelatedToID] = 1 Then
        'just an excavation id needed
        Me![txtExcavationIDNumber].Enabled = True
        Me![txtExcavationIDNumber].Locked = False
        
        Me![cboFindLetter].Enabled = False
        Me![cboFindLetter].Locked = True
        Me![cboFindLetter].BackColor = -2147483633
        
        Me![txtFindSampleNumber].Enabled = False
        Me![txtFindSampleNumber].Locked = True
        Me![txtFindSampleNumber].BackColor = -2147483633
        
        Me![lblRelatedTo].Caption = "Type"
        Me![lblExIDNumber].Caption = "Number"
        Me![lblFindSampleNumber].Caption = ""
        Me![lblFindLetter].Caption = ""
    ElseIf Forms![Conserv: Basic Record]![RelatedToID] = 4 Then
        'other category just needs comment field filled at present
        Me![txtExcavationIDNumber].Enabled = False
        Me![txtExcavationIDNumber].Locked = True
        Me![txtExcavationIDNumber].BackColor = -2147483633
        
        Me![cboFindLetter].Enabled = False
        Me![cboFindLetter].Locked = True
        Me![cboFindLetter].BackColor = -2147483633
        
        Me![txtFindSampleNumber].Enabled = False
        Me![txtFindSampleNumber].Locked = True
        Me![txtFindSampleNumber].BackColor = -2147483633
        Me![lblFindSampleNumber].Caption = ""
        Me![lblFindLetter].Caption = ""
        'nothing to view so hide button
        Me![cmdView].Visible = False
    Else
        'all other situations
        Me![txtExcavationIDNumber].Enabled = True
        Me![txtExcavationIDNumber].Locked = False
        Me![txtExcavationIDNumber].BackColor = 16777215
        
        Me![cboFindLetter].Enabled = True
        Me![cboFindLetter].Locked = False
        Me![cboFindLetter].BackColor = 16777215
        
        Me![txtFindSampleNumber].Enabled = True
        Me![txtFindSampleNumber].Locked = False
        Me![txtFindSampleNumber].BackColor = 16777215
        
        Me![lblRelatedTo].Caption = "Type"
        Me![lblExIDNumber].Caption = "Unit No."
        
        'If Me![RelatedToID] = 2 Then 'object
        If Forms![Conserv: Basic Record]![RelatedToID] = 2 Then
            'object
            Me![lblFindSampleNumber].Caption = "Find No."
            'set list to default to object
            If IsNull(Me![cboRelatedToSubTypeID]) Then Me![cboRelatedToSubTypeID] = 5
            If IsNull(Me![cboFindLetter]) Then Me![cboFindLetter] = "X"
            Me![lblComment].Caption = "Find Type"
        'ElseIf Me![RelatedToID] = 3 Then 'sample
        ElseIf Forms![Conserv: Basic Record]![RelatedToID] = 3 Then
            'sample
            Me![lblFindSampleNumber].Caption = "Sample No."
            Me![cboFindLetter].RowSource = "SELECT [Exca: Letter codes].Letter, [Exca: Letter codes].ShortDescription FROM [Exca: Letter codes] WHERE ([Exca: Letter codes].ShortDescription like '%sample%');"
            'set list to default to sample
            If IsNull(Me![cboRelatedToSubTypeID]) Then Me![cboRelatedToSubTypeID] = 6
            If IsNull(Me![cboFindLetter]) Then Me![cboFindLetter] = "s"
        Else
            Me![lblFindSampleNumber].Caption = "Number"
        End If
        Me![lblFindLetter].Caption = "Letter"
    End If

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub
