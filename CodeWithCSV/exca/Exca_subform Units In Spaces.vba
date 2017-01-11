Option Compare Database
Option Explicit
'**********************************************************
' This subform is new in version 9.2 - as a feature can be
' in many spaces the space field has been removed from the
' Feature tables and normalised out int Exca: Features in Spaces
' SAJ v9.2
'**********************************************************

Private Sub cmdGoToSpace_Click()
'***********************************************************************
' Open space form read only from here.
' Also becuase they can enter a space num that not exist yet (SF requirement)
' need to see if the record exists before opening the form (otherwise be blank)
'
' SAJ v9.2
'***********************************************************************
On Error GoTo Err_cmdGoToSpace_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    
    stDocName = "Exca: Space Sheet"
    
    If Not IsNull(Me![txtIn_Space]) Or Me![txtIn_Space] <> "" Then
        'check that space num does exist
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![txtIn_Space])
        If IsNull(checknum) Then
            'number not exist - now see what permissions user has
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Space Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
        
                If retval = vbNo Then
                    MsgBox "No space record to view, please alert the your team leader about this.", vbExclamation, "Missing Space Record"
                Else
                    'add new records behind scences
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Space Sheet] ([Space Number], [Area]) VALUES (" & Me![txtIn_Space] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = " & Me![txtIn_Space], acFormEdit, acDialog
                End If
            Else
                'user is readonly so just tell them record not exist
                MsgBox "Sorry but this space record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Space Record"
            End If
        Else
            'record exists - open it
            stLinkCriteria = "[Space Number]=" & Me![txtIn_Space]
            'DoCmd.OpenForm stDocName, , , stLinkCriteria
            'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
            'decided against dialog as can open other forms on the feature form and they would appear underneath it
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Space number to view", vbInformation, "No Space Number"
    End If
    
Exit_cmdGoToSpace_Click:
    Exit Sub


Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToSpace_Click
    

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub


Private Sub Form_Current()
'v9.2 - enable/disable button depending on contents of space field
On Error GoTo err_Current
    
    If Me![txtIn_Space] = "" Or IsNull(Me![txtIn_Space]) Then
        Me![cmdGoToSpace].Enabled = False
    Else
        Me![cmdGoToSpace].Enabled = True
    End If


Exit Sub
err_Current:
    Call General_Error_Trap
    Exit Sub
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





Private Sub txtIn_Space_AfterUpdate()
'***********************************************************************
' Validity check to ensure space num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter space num when they can
'
' SAJ v9.2
'***********************************************************************
On Error GoTo err_txtIn_Space_AfterUpdate

Dim checknum, msg, retval, sql, insertArea

If Me![txtIn_Space] <> "" Then
    'first check its valid
    If IsNumeric(Me![txtIn_Space]) Then
    
        'check that space num does exist
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![txtIn_Space])
        If IsNull(checknum) Then
            msg = "This Space Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                'add new records behind scences
                If Forms![Exca: Unit Sheet]![Area] <> "" Then
                    insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                Else
                    insertArea = Null
                End If
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number], [Area]) VALUES (" & Me![txtIn_Space] & ", " & insertArea & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = " & Me![txtIn_Space], acFormEdit, acDialog
                ''DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![In_feature] & ",Area:" & Forms![Exca: Unit Sheet]![Area]
            End If
        Else
            'valid number, enable view button
            Me![cmdGoToSpace].Enabled = True
        End If
    
    Else
        'not a vaild numeric space number
        MsgBox "The Space number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_txtIn_Space_AfterUpdate:
    Call General_Error_Trap
    Exit Sub



End Sub

Private Sub txtIn_Space_BeforeUpdate(Cancel As Integer)
'new 2009 winter - if a user gets number wrong they can't delete it and many then put 0
'space 0 keeps appearing and we don't want it so put a check in
On Error GoTo err_spacebefore

If Me![txtIn_Space] = 0 Then
        MsgBox "Space 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
      
        Cancel = True
        'Me![txtIn_Building].Undo
        SendKeys "{ESC}" 'seems to need it done 3x
        SendKeys "{ESC}"
        SendKeys "{ESC}"
End If
Exit Sub

err_spacebefore:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtIn_Space_LostFocus()
'May 2009 - when the form loses focus requery the subform that shows the building numbers as this
'is dependant on the spaces entered
On Error GoTo err_lost

    Forms![Exca: Unit Sheet]![Exca: subform Units  in Buildings].Form.Requery

Exit Sub

err_lost:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub

