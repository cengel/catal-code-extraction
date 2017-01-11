Option Compare Database
Option Explicit

Private Sub cmdDelete_Click()
'new in 2009 - first check they really want to delete
On Error GoTo err_delete

    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Me![cmdDelete].Enabled = True
    
        Dim resp
        resp = MsgBox("Do you really want to delete the skeleton " & Me![UnitNumber] & ".B" & Me![txtIndivid] & "? This will remove this individual from the database completely and permanently. ", vbCritical + vbYesNo, "Confirm Deletion")
        If resp = vbYes Then
            'ok delete this relationship - must do it both ways
            Dim sql
            sql = "Delete FROM [HR_BasicSkeletonData] WHERE [UnitNumber] = " & Me![UnitNumber] & " AND [Individual Number] = " & Me![txtIndivid] & ";"
            DoCmd.RunSQL sql
            Me.Requery
            'remove focus from the delete button
            'DoCmd.GoToControl "cmdAddRelation"
        End If
    Else
        MsgBox "You do not have permissions to delete, please contact your team leader", vbExclamation, "Insufficient permissions"
    End If
   
Exit Sub

err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdView_Click()
On Error GoTo Err_cmdView_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    stLinkCriteria = "[UnitNumber] = " & Me![UnitNumber] & " AND [Individual Number] = " & Me![txtIndivid]
    
    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    Forms!FRM_SkeletonDescription!cboFind.RowSource = "SELECT [HR_BasicSkeletonData].[UnitNumber], [HR_BasicSkeletonData].[Individual number] FROM HR_BasicSkeletonData WHERE [HR_BasicSkeletonData].[UnitNumber] = " & Me![UnitNumber] & " ORDER BY [HR_BasicSkeletonData].[UnitNumber], [HR_BasicSkeletonData].[Individual number];"
    'saj
    DoCmd.Close acForm, "FRM_UnitDescription"
Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'new 2009 - delete button
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
