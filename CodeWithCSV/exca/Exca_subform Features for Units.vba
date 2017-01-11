Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
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

Private Sub go_to_feature_Click()
'***********************************************************************
' changed open form string to be read only from here, also call general
' error trap, plus no feature num catch.
'
' Also becuase they can enter a feature num that not exist yet (SF requirement)
' need to see if the record exists before opening the form (otherwise be blank)
'
' SAJ v9.1
'***********************************************************************
On Error GoTo Err_go_to_feature_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    
    stDocName = "Exca: Feature Sheet"
    
    If Not IsNull(Me![In_feature]) Or Me![In_feature] <> "" Then
        'check that feature num does exist
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![In_feature])
        If IsNull(checknum) Then
            'number not exist - now see what permissions user has
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Feature Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
        
                If retval = vbNo Then
                    MsgBox "No feature record to view, please alert the your team leader about this.", vbExclamation, "Missing Feature Record"
                Else
                    'add new records behind scences
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Features] ([Feature Number], [Area]) VALUES (" & Me![In_feature] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , "[Feature Number] = " & Me![In_feature], acFormEdit, acDialog
                End If
            Else
                'user is readonly so just tell them record not exist
                MsgBox "Sorry but this feature record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Feature Record"
            End If
        Else
            'record exists - open it
            stLinkCriteria = "[Feature Number]=" & Me![In_feature]
            'DoCmd.OpenForm stDocName, , , stLinkCriteria
            'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
            'decided against dialog as can open other forms on the feature form and they would appear underneath it
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Feature number to view", vbInformation, "No Feature Number"
    End If
    
Exit_go_to_feature_Click:
    Exit Sub


Err_go_to_feature_Click:
    Call General_Error_Trap
    Resume Exit_go_to_feature_Click
    
End Sub

Private Sub In_feature_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure feature num entered here is ok
' if not tell the user and allow them to enter. SF not want it to restrict
' entry and trusts excavators to enter feature num when they can
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_In_feature_AfterUpdate

Dim checknum, msg, retval, sql, insertArea

If Me![In_feature] <> "" Then
    'first check its valid
    If IsNumeric(Me![In_feature]) Then
    
        'check that feature num does exist
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![In_feature])
        If IsNull(checknum) Then
            msg = "This Feature Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                'add new records behind scences
                If Forms![Exca: Unit Sheet]![Area] <> "" Then
                    insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                Else
                    insertArea = Null
                End If
                sql = "INSERT INTO [Exca: Features] ([Feature Number], [Area]) VALUES (" & Me![In_feature] & ", " & insertArea & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , "[Feature Number] = " & Me![In_feature], acFormEdit, acDialog
                ''DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![In_feature] & ",Area:" & Forms![Exca: Unit Sheet]![Area]
            End If
        Else
            'valid number, enable view button
            Me![go to feature].Enabled = True
        End If
    
    Else
        'not a vaild numeric feature number
        MsgBox "The Feature number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_In_feature_AfterUpdate:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub In_feature_BeforeUpdate(Cancel As Integer)
'new 2009 winter - if a user gets number wrong they can't delete it and many then put 0
'building 0 keeps appearing and we don't want it so put a check in
On Error GoTo err_featurebefore

If Me![In_feature] = 0 Then
        MsgBox "Feature 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
      
        Cancel = True
        'Me![txtIn_Building].Undo
        SendKeys "{ESC}" 'seems to need it done 3x
        SendKeys "{ESC}"
        SendKeys "{ESC}"
End If
Exit Sub

err_featurebefore:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub

Sub Command5_Click()
On Error GoTo Err_Command5_Click


    DoCmd.GoToRecord , , acLast

Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click
    
End Sub
