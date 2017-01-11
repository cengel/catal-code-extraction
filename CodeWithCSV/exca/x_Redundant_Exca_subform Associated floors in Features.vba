Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Existing date change update - added error trap v9.1
On Error GoTo err_Form_BeforeUpdate

Me![Date changed] = Now()
Forms![Exca: Feature Sheet]![dbo_Exca: FeatureHistory].Form![lastmodify].Value = Now()

Exit Sub
err_Form_BeforeUpdate:
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
    If permiss = "ADMIN" Or permiss = "RW" Then
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

Private Sub Unit_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure unit num entered here is exists and
' that it has the data category - floors (use) OR construction/makeup/packaging
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_Unit_AfterUpdate

'old code that was here
'Me.Requery
'DoCmd.GoToRecord , , acLast
Dim checknum, msg, retval, checknum2

If Me![Unit] <> "" Then
    'first check its valid
    If IsNumeric(Me![Unit]) Then
    
        'check that Unit num does exist
        checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![Unit])
        If IsNull(checknum) Then
            msg = "This Unit Number DOES NOT EXIST in the database, it cannot be used here until it has been entered."
            'retVal = MsgBox(msg, vbInformation + vbOK, "Unit Number does not exist")
            MsgBox msg, vbInformation, "Unit Number does not exist"
            If Not IsNull(Me![Unit].OldValue) Then
                Me![Unit] = Me![Unit].OldValue
            Else
                Me.Undo
            End If
            DoCmd.GoToControl "Unit"
        Else
            'valid number, now check its data category
            checknum2 = DLookup("[Data Category]", "[Exca: Unit Data Categories]", "[Unit Number] = " & Me![Unit])
                If Not IsNull(checknum2) Then 'there is a space for this related feature
                    If UCase(checknum2) <> "FLOORS (USE)" And UCase(checknum2) <> "CONSTRUCTION/MAKE-UP/PACKING" Then
                        'do not allow entry if units datacategory is not floor or construction
                        msg = "This entry is not allowed:  Unit (" & Me![Unit] & ")"
                        msg = msg & " has the data category " & checknum2 & ", only Units with the category 'Floor(use)' or 'construction/make-up/packing' are valid here."
                        msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
                        MsgBox msg, vbExclamation, "Data Category problem"
                        
                        ''MsgBox "To remove this Associated Floors link completely press ESC", vbInformation, "Help Tip"
                        
                        'reset val to previous val if is one or else remove it completely
                        If Not IsNull(Me![Unit].OldValue) Then
                            Me![Unit] = Me![Unit].OldValue
                        Else
                            Me.Undo
                        End If
                        DoCmd.GoToControl "Unit"
                    End If
                Else
                    'the data category for this unit has not been filled out yet, again do not allow link
                     'other possible actions here would be to allow the link therefore no code here
                     'or to fill out the datacategory automatically in code, but would have to know which one of the 2 cats
                    msg = "This entry is not allowed as Unit (" & Me![Unit] & ")"
                    msg = msg & " has no data category entered" & checknum2 & ", only Units with the category 'Floor(use)' or 'construction/make-up/packing' are valid here."
                    msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please update the Unit record first."
                    MsgBox msg, vbExclamation, "No Data Category"
                        
                    ''MsgBox "To remove this Associated Floors link completely press ESC", vbInformation, "Help Tip"
                        
                    'reset val to previous val if is one or else remove it completely
                    If Not IsNull(Me![Unit].OldValue) Then
                        Me![Unit] = Me![Unit].OldValue
                    Else
                        Me.Undo
                    End If
                    DoCmd.GoToControl "Unit"
                End If
        End If
    
    Else
        'not a vaild numeric unit number
        MsgBox "The Unit number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_Unit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub

End Sub

Sub Command5_Click()
'v9.1 - this control not seem to exist - comment out
'On Error GoTo Err_Command5_Click


'    DoCmd.GoToRecord , , acLast

'Exit_Command5_Click:
'    Exit Sub

'Err_Command5_Click:
'    MsgBox Err.Description
'    Resume Exit_Command5_Click
    
End Sub
Sub go_to_unit_Click()
'********************************************
'Existing code for go to unit button, added
'general error trap and check that Unit num there
' now open readonly
'SAJ v9.1
'********************************************
On Error GoTo Err_go_to_unit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    
    If Me![Unit] <> "" Then
        stLinkCriteria = "[Unit Number]=" & Me![Unit]
        DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormReadOnly
    Else
        MsgBox "No Unit number to show", vbInformation, "No Unit Number"
    End If
Exit_go_to_unit_Click:
    Exit Sub

Err_go_to_unit_Click:
    Call General_Error_Trap
    Resume Exit_go_to_unit_Click
    
End Sub
