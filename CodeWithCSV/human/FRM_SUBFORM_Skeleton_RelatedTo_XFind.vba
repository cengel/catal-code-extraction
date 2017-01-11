Option Compare Database
Option Explicit

Private Sub cmdAddRelation_Click()
'new 2009 end of season
'allow relationship to be entered = this individual came down off site as an X find - link it to this x find number
On Error GoTo err_skel

    'obtain X find number
    Dim getXFind, getNotes, sql
    getXFind = InputBox("Enter the X find number that this individual number relates to:", "X Find Number Required")
    If getXFind <> "" Then
        getNotes = InputBox("Enter any notes or comments about this relationship:", "Notes or Comments")
        
        If getNotes <> "" Then
            sql = "INSERT INTO [HR_Skeleton_RelatedTo_XFind] ([Unit], [IndividualNumber], [XfindNumber], [Notes]) VALUES (" & Forms![FRM_SkeletonDescription]![txtUnit] & ", " & Forms![FRM_SkeletonDescription]![txtIndivid] & ", " & getXFind & ", '" & getNotes & "');"
        Else
            sql = "INSERT INTO [HR_Skeleton_RelatedTo_XFind] ([Unit], [IndividualNumber], [XfindNumber]) VALUES (" & Forms![FRM_SkeletonDescription]![txtUnit] & ", " & Forms![FRM_SkeletonDescription]![txtIndivid] & ", " & getXFind & ");"
        End If
        DoCmd.RunSQL sql
        
        Me.Requery
    End If
    DoCmd.GoToControl "cmdAddRelation"

Exit Sub

err_skel:
    MsgBox Err.Description
    Exit Sub

End Sub
Private Sub cmdDelete_Click()
'delete relationship
On Error GoTo err_cmdDelete

    'first check they really want to delete
    Dim resp
    resp = MsgBox("Do you really want to delete the relationship between skeleton " & Me![Unit] & ".B" & Me![IndividualNumber] & " and X find number" & Me![Unit] & ".X" & Me![XFindNumber] & "?", vbCritical + vbYesNo, "Confirm Deletion")
    If resp = vbYes Then
        'ok delete this relationship - must do it both ways
        Dim sql
        sql = "Delete FROM [HR_Skeleton_RelatedTo_XFind] WHERE [Unit] = " & Me![Unit] & " AND [IndividualNumber] = " & Me![IndividualNumber] & " AND [XFindNumber] = " & Me![XFindNumber] & ";"
        DoCmd.RunSQL sql
        
        Me.Requery
        'remove focus from the delete button
        DoCmd.GoToControl "cmdAddRelation"
        
    End If
    
Exit Sub

err_cmdDelete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'new 2009 - disable delete button where not permissions
On Error GoTo err_open

Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Me![cmdDelete].Enabled = True
    Else
        Me![cmdDelete].Enabled = False
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
