Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Feature Sheet]![dbo_Exca: FeatureHistory].Form![lastmodify].Value = Now()
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
    
    '2009 - OMG the feature type field is being unlocked so can be EDITED here - thats terrible
    Me![Feature Type].Locked = True
    Me![Feature Type].Enabled = False
    Me![FeatureSubType].Locked = True
    Me![FeatureSubType].Enabled = False
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub To_feature_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure related feature num entered here is ok
' if not tell the user and enter it if necess. SF not want it to restrict
' entry and trusts excavators to enter feature num when they can - however
' only allow the entry of existing num if its in the same space
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_To_feature_AfterUpdate

Dim checknum, msg, retval, sql, currentFeature, checknum2, featureRel, checknum3, myrs As DAO.Recordset, mydb As DAO.Database

If Me![To_feature] <> "" Then
    'first check its valid
    If IsNumeric(Me![To_feature]) Then
    
        'check that building num does exist
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![To_feature])
        If IsNull(checknum) Then
            msg = "The Feature Number " & Me![To_feature] & " DOES NOT EXIST in the database. The system can enter it for you ready for you to update later."
            msg = msg & Chr(13) & Chr(13) & "Would you like the system to create this feature number now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
        
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
               'Dim myf As Form
              ' myf.Name = "Exca: Feature Sheet"
               'myf.Show
              ' DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![To_feature] & ",Area:" & Forms![Exca: Feature Sheet]![Combo27]
                'ok add the feature number - add it to table but then need to update form
                'annoyingly when requery a form it retains no memory of where you were and flicks to begining
                'warn user screen will refresh and put in hourglass to make it clear its processing for slow links
                
                'grab current feature number to return to
                currentFeature = Me![Feature Number]
                'create feature record
                sql = "INSERT INTO [Exca: Features] ([Feature Number]) VALUES (" & Me![To_feature] & ");"
                DoCmd.RunSQL sql
                'tell user theres going to be a refresh
                MsgBox "Feature " & Me![To_feature] & " has been created in the database. This screen will now refresh itself.", vbInformation, "System updating"
                DoCmd.Hourglass True
                Forms![Exca: Feature Sheet].Requery
                DoCmd.GoToControl Forms![Exca: Feature Sheet]![Feature Number].Name 'goto main forms feature num
                DoCmd.FindRecord currentFeature 'find the number user was editing before
                DoCmd.Hourglass False
            End If
        Else
            'valid number, but must check its in same space
            'SEASON 2009 - this field went in v9.2 but this bug not picked up until v12.6 - how did it take so long!!!
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'If Not IsNull(Forms![Exca: Feature Sheet]![Space]) Or Forms![Exca: Feature Sheet]![Space] <> "" Then
            '    checknum2 = DLookup("[Space]", "[Exca: Features]", "[Feature Number] = " & Me![To_feature])
            '    If Not IsNull(checknum2) Then 'there is a space for this related feature
            '        If checknum2 <> Forms![Exca: Feature Sheet]![Space] Then 'do not allow entry if space numbers differ
            '            msg = "This entry is not allowed:  feature (" & Me![To_feature] & ")"
            '            msg = msg & " is in Space " & checknum2 & " but Feature " & Forms![Exca: Feature Sheet]![Feature Number]
            '            msg = msg & " is in Space " & Forms![Exca: Feature Sheet]![Space]
            '            msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
            '            MsgBox msg, vbExclamation, "Space mis-match"
            '
            '            MsgBox "To remove this relationship completely press ESC", vbInformation, "Help Tip"
            '
            '            'reset val to previous val if is one or else remove it completely
            '            If Not IsNull(Me![To_feature].OldValue) Then
            '                Me![To_feature] = Me![To_feature].OldValue
            '            Else
            '                featureRel = Me![Relation]
            '                Me.Undo
            '                Me![Relation] = featureRel
            '            End If
            '            DoCmd.GoToControl "Feature Type"
            '            DoCmd.GoToControl "To_Feature"
            '        End If
            '    End If
            'End If
            
            'first check this feature has a space/s
            checknum2 = DLookup("[In_Space]", "[Exca: Features in Spaces]", "[Feature] = " & Me![Feature Number])
            If Not IsNull(checknum2) Then 'there is a space for main feature
                checknum3 = DLookup("[In_Space]", "[Exca: Features in Spaces]", "[Feature] = " & Me![To_feature])
                If Not IsNull(checknum3) Then 'there is a space for related feature
                    'ok so both have at least one space number so lets check they are in same space
                    'never done this before but this query seems to work to see if the features have the same space number
                    sql = "SELECT [Exca: Features in Spaces].Feature, [Exca: Features in Spaces].In_Space, [Exca: Features in Spaces_1].Feature, [Exca: Features in Spaces_1].In_Space" & _
                            " FROM [Exca: Features in Spaces] INNER JOIN [Exca: Features in Spaces] AS [Exca: Features in Spaces_1] ON [Exca: Features in Spaces].In_Space = [Exca: Features in Spaces_1].In_Space " & _
                            " WHERE ([Exca: Features in Spaces].Feature =" & Me![Feature Number] & ")  AND ([Exca: Features in Spaces_1].Feature=" & Me![To_feature] & ");"
                    Set mydb = CurrentDb
                    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
                    
                    If myrs.EOF And myrs.BOF Then
                        'contrary to what had once been conceived, features might relate to features from other spaces. Possibility of inserting such relations is now possible,
                        'but user has to be informed on possible erroneous relation; DL 2015
                        Dim response
                        msg = "This entry is not allowed because these two features are not currently in the same Space. They must be in the same space to create a relationship."
                        msg = msg & Chr(13) & Chr(13) & "Are you sure that " & Parent![Feature Number] & " is " & Me![Relation] & " " & Me![To_feature] & "?"
                        response = MsgBox(msg, vbYesNo + vbQuestion, "Space mis-match")
                        If response = vbYes Then
                        Else
                    
                        'if no old value it reverts to 0 and we can't have the To_Feature number as 0
                        'MsgBox "To remove this relationship completely press ESC", vbInformation, "Help Tip"
            
                       'reset val to previous val if is one or else remove it completely
                        If Not IsNull(Me![To_feature].OldValue) Then
                            Me![To_feature] = Me![To_feature].OldValue
                            DoCmd.GoToControl "To_Feature"
                        Else
                            'featureRel = Me![Relation]
                            Me.Undo
                            'Me![Relation] = featureRel
                            'Me![To_feature] = ""
                             DoCmd.GoToControl "Relation"
                        End If
                        'DoCmd.GoToControl "Feature Type"
                    End If
                        
                End If
            End If
        End If
        
    End If
    
    Else
        'not a vaild numeric feature number
        MsgBox "The Feature number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_To_feature_AfterUpdate:
    Call General_Error_Trap
    'just in case fell over when hourglass on - turn it off
    DoCmd.Hourglass False
    Exit Sub

End Sub
