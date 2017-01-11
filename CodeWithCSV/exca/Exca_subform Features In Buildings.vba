Option Compare Database
Option Explicit
'**********************************************************
' This subform is new in version 9.2 - as a feature can be
' in many Buildings the Building field has been removed from the
' Feature tables and normalised out int Exca: Features in Buildings
' SAJ v9.2
'**********************************************************


Private Sub cmdGoToBuilding_Click()
'***********************************************************************
' Open Building form read only from here.
' Also becuase they can enter a Building num that not exist yet (SF requirement)
' need to see if the record exists before opening the form (otherwise be blank)
'
' SAJ v9.2
'***********************************************************************
On Error GoTo Err_cmdGoToBuilding_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    
    stDocName = "Exca: Building Sheet"
    
    If Not IsNull(Me![txtIn_Building]) Or Me![txtIn_Building] <> "" Then
        'check that Building num does exist
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![txtIn_Building])
        If IsNull(checknum) Then
            'number not exist - now see what permissions user has
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Building Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
        
                If retval = vbNo Then
                    MsgBox "No Building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
                Else
                    'add new records behind scences
                    If Forms![Exca: Feature Sheet]![Combo27] <> "" Then
                        insertArea = "'" & Forms![Exca: Feature Sheet]![Combo27] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Building Details] ([Number], [Area]) VALUES (" & Me![txtIn_Building] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![txtIn_Building], acFormEdit, acDialog
                End If
            Else
                'user is readonly so just tell them record not exist
                MsgBox "Sorry but this Building record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Building Record"
            End If
        Else
            'record exists - open it
            stLinkCriteria = "[Number]=" & Me![txtIn_Building]
            'DoCmd.OpenForm stDocName, , , stLinkCriteria
            'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
            'decided against dialog as can open other forms on the feature form and they would appear underneath it
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Building number to view", vbInformation, "No Building Number"
    End If
    
Exit_cmdGoToBuilding_Click:
    Exit Sub


Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToBuilding_Click
    

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
''MAY 2009 - building number from space now so this form is now read only
'Me![Date Changed] = Now()
End Sub


Private Sub Form_Current()
'v9.2 - enable/disable button depending on contents of Building field
On Error GoTo err_Current
    
    If Me![txtIn_Building] = "" Or IsNull(Me![txtIn_Building]) Then
        Me![cmdGoToBuilding].Enabled = False
    Else
        Me![cmdGoToBuilding].Enabled = True
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
' MAY 2009 - building now from space so this is always readonly
'**********************************************************************
On Error GoTo err_Form_Open

'    Dim permiss
'    permiss = GetGeneralPermissions
'    If permiss = "ADMIN" Or permiss = "RW" Then
'        ToggleFormReadOnly Me, False
'    Else
'        'set read only form here, just once
        ToggleFormReadOnly Me, True
'    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtIn_Building_AfterUpdate()
'***********************************************************************
' Validity check to ensure building num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter building num when they can
'
' SAJ v9.2

'***********************************************************************
On Error GoTo err_txtIn_Space_AfterUpdate

'Dim checknum, msg, retVal, sql, insertArea
'
'If Me![txtIn_Building] <> "" Then
'    'first check its valid
'    If IsNumeric(Me![txtIn_Building]) Then
'
'        'check that Building num does exist
'        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![txtIn_Building])
'        If IsNull(checknum) Then
'            msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
'            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
'            retVal = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
'
'            If retVal = vbNo Then
'                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
'            Else
'                'add new records behind scences
'                If Forms![Exca: Feature Sheet]![Combo27] <> "" Then
'                    insertArea = "'" & Forms![Exca: Feature Sheet]![Combo27] & "'"
'                Else
'                    insertArea = Null
'                End If
'                sql = "INSERT INTO [Exca: Building Details] ([Number], [Area]) VALUES (" & Me![txtIn_Building] & ", " & insertArea & ");"
'                DoCmd.RunSQL sql
'                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![txtIn_Building], acFormEdit, acDialog
'                ''DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![In_feature] & ",Area:" & Forms![Exca: Unit Sheet]![Area]
'            End If
'        Else
'            'valid number, enable view button
'            Me![cmdGoToBuilding].Enabled = True
'        End If
'
'    Else
'        'not a vaild numeric Building number
'        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
'    End If
'End If

Exit Sub

err_txtIn_Space_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtIn_Building_BeforeUpdate(Cancel As Integer)
'new 2009 winter - if a user gets number wrong they can't delete it and many then put 0
'building 0 keeps appearing and we don't want it so put a check in
'MAY 2009 - building number from space now so this never happens
On Error GoTo err_buildingbefore

'If Me![txtIn_Building] = 0 Then
'        MsgBox "Building 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
'
'        Cancel = True
'        'Me![txtIn_Building].Undo
'        SendKeys "{ESC}" 'seems to need it done 3x
'        SendKeys "{ESC}"
'        SendKeys "{ESC}"
'End If
Exit Sub

err_buildingbefore:
    Call General_Error_Trap
    Exit Sub
End Sub



