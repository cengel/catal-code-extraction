Option Compare Database
Option Explicit

Private Sub Amount_AfterUpdate()
'2009 first step on the road to sorting out volume issues, amount field still in process of being
'cleaned so is not numeric, trying here to at least force user to put in a number
'they get this message then they can move off this field so not foolproof, I tried it in LostFocus
'but could get a way of knowing if its a new entry or existing one and don't want people getting stuck
'in old entries that they don't know how to convert and making more of a mess.
'Once cleaned this field will be data type numeric and this problem will go away
On Error GoTo err_Amount

    If Not IsNumeric(Me!Amount) Then
        MsgBox Me!Amount & " is not a numeric amount, please enter the amount if Litres but as a number only", vbInformation, "Invalid Amount"
        DoCmd.GoToControl "SampleType"
        Me!Amount.SetFocus
        'Me![Amount] = Me![Amount] & " " 'was trying to change its value to afterupdate to trigger until it was numeric but no effect
    End If
Exit Sub

err_Amount:
    Call General_Error_Trap
    Exit Sub
End Sub





Private Sub Amount__ltrs__AfterUpdate()

End Sub

Private Sub Amount__ltrs__Change()

End Sub

Private Sub Amount__ltrs__LostFocus()

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Me![Date changed] = Now()
End Sub




Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' Set up form view depending on permissions
' SAJ v9.1
'**********************************************************************
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ToggleFormReadOnly Me, False
    Else
        'set read only form here, just once
        ToggleFormReadOnly Me, True
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub SampleType_AfterUpdate()
'new 2009 if sample amount present fill it in
On Error GoTo err_sampletype
    
    If Me![SampleType].Column(1) <> "" Then
        If Me![Amount] <> "" Then
                      
            Dim response
            response = MsgBox("There is a default amount for this sample type of " & Me![SampleType].Column(1) & ". Do you wish to overwrite the current amount?", vbYesNo + vbQuestion, "Amount?")
            If response = vbYes Then Me![Amount] = Me![SampleType].Column(1)
        Else
            Me![Amount] = Me![SampleType].Column(1)
        End If
    End If
    
    '2009 sample type list is now locked down so remind to make sure sub samples get entered correctly
    If InStr(Me![SampleType], "subsample") > 0 Then
        MsgBox "You must write the original sample number from which you are taking the sample in the Comment field as well as the details of the purpose of the sample and amount. No amount is required in the amount column itself", vbExclamation, "Sub sample requirements"
        ''2010 - LOCK AMOUNT field
        Me![Amount (ltrs)].Locked = True
        Me![Amount (ltrs)].Enabled = False
    Else
        ''2010 ensure amount field unlocked
        Me![Amount (ltrs)].Locked = False
        Me![Amount (ltrs)].Enabled = True
    End If
    If Me![SampleType] = "" Or IsNull(Me![SampleType]) Then
        MsgBox "YOU MUST ENTER A SAMPLE TYPE", vbExclamation, "Missing Sample Type"
        '2010 - enough of this nonsense - insist
        ''Me![SampleType].SetFocus
    End If
Exit Sub

err_sampletype:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub SampleType_LostFocus()
'new 2010 - this is getting tough - see if it works. You cannot leave this field until a value is entered. Could lead to
'crap being entered here I guess but its worth a try.
'SAJ 7July10
On Error GoTo err_sampletype

If Me![SampleType] = "" Or IsNull(Me![SampleType]) Then
    MsgBox "YOU MUST ENTER A SAMPLE TYPE", vbExclamation, "Missing Sample Type"
    '2010 - enough of this nonsense - insist
    ''Me![SampleType].SetFocus
    DoCmd.GoToControl Me![X].Name
    
    DoCmd.GoToControl Me![SampleType].Name
End If
Exit Sub

err_sampletype:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub SampleType_NotInList(NewData As String, response As Integer)
'***********************************************************************
' Intro of a validity check to make users a little more aware of the data
' they are entering here. The combo here is trying to prevent different entries
' that represent the same thing. Users are allowed to enter new values but just made aware
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Sampletype_NotInList

''2009 locked down the list as people not using it sensibly at all
''Dim retVal, sql
''retVal = MsgBox("This Sample Type does not appear in the pre-defined list. Have you checked the list to make sure there is no match?", vbQuestion + vbYesNo, "New Sample Type")
''If retVal = vbYes Then
''    MsgBox "Ok this sample type will now be added to the list", vbInformation, "New Sample Type Allowed"
''    'allow value,
''    response = acDataErrAdded
''     'Me![SampleType].LimitToList = False 'turn off limit to list so record can be saved
''
''    sql = "INSERT INTO [Exca:SampleTypeLOV] ([SampleType]) VALUES ('" & NewData & "');"
''    DoCmd.RunSQL sql
''    'dont need any of this for this situation
''    'DoCmd.GoToControl "SampleType"
''    'DoCmd.RunCommand acCmdSaveRecord 'save rec
''    'Me![SampleType].Requery 'requery combo to get new value in list
''    'Me![SampleType].LimitToList = True 'put back on limit to list
''Else
''    'no leave it so they can edit it
''    response = acDataErrContinue
''End If

MsgBox "This Sample Type is not found in the current list, look carefully and consult the Type list via the button above. " & Chr(13) & Chr(13) & "There is a new format for sample types. This is: main type-subtype " & Chr(13) & Chr(13) & "eg: Flotation-routine" & Chr(13) & Chr(13) & "If you really cannot find your sample type then please use: Other and write specific details in the comment field. Then tell your Supervisor who will inform the project team.", vbExclamation, "Sample Types"
SendKeys "{ESC}{ESC}"
Exit Sub

err_Sampletype_NotInList:
    Call General_Error_Trap
''    Me![SampleType].LimitToList = True
    Exit Sub


End Sub
