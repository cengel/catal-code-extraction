Option Compare Database
Option Explicit




Private Sub AddDetail_Click()
On Error GoTo Err_AddDetail_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim relationexists, msg, retVal, sql, permiss
    
    stDocName = "Anthracology: Dendro"
    relationexists = DLookup("[GID]", "Anthracology: Dendro", "[GID] = '" & Me![GID] & "'")
        If IsNull(relationexists) Then
            'number not exist - now see what permissions user has
            msg = "Details for this flotation have not been entered yet."
            msg = msg & Chr(13) & Chr(13) & "Would you like to add this information now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Detail does not exist")
        
            If retVal = vbNo Then
                MsgBox "Details will not be added.", vbExclamation, "Missing Detail Record"
            Else
                'add new records behind scences
                sql = "INSERT INTO [Anthracology: Dendro] ([id],[GID]) VALUES (1,'" & Me![GID] & "');"
                DoCmd.RunSQL sql
                DoCmd.OpenForm stDocName, acNormal, , "[GID] = '" & Me![GID] & "'", acFormEdit, acDialog
            End If
        Else
            MsgBox "Details have already been added.", vbInformation, "Existing Detail Record"
        End If


Exit_AddDetail_Click:
    Exit Sub

Err_AddDetail_Click:
    MsgBox Err.Description
    Resume Exit_AddDetail_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_BUpd

Debug.Print Now()
Forms![Anthracology: Sheet]![timestamp] = Now()

Exit Sub

err_BUpd:
    Call General_Error_Trap
    Exit Sub

End Sub


Private Sub Form_Current()
'check whether taxa determination has relation in dendro_detail

Dim relationexists
relationexists = DLookup("[GID]", "Anthracology: Dendro", "[GID] = '" & Me![GID] & "'")
If Not IsNull(relationexists) Then
    Me![goto_DendroDetails].Visible = True
    Me![AddDetail].Visible = False
Else
    Me![goto_DendroDetails].Visible = False
    Me![AddDetail].Visible = True
End If


End Sub

Private Sub goto_DendroDetails_Click()
On Error GoTo Err_goto_DendroDetails_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim relationexists
    
    stDocName = "Anthracology: Dendro"
    
    relationexists = DLookup("[GID]", "Anthracology: Dendro", "[GID] = '" & Me![GID] & "'")
    If IsNull(relationexists) Then
        'record does not exist
    Else
        'record exists - open it
        stLinkCriteria = "[GID]='" & Me![GID] & "'"
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
    End If
    

Exit_goto_DendroDetails_Click:
    Exit Sub

Err_goto_DendroDetails_Click:
    MsgBox Err.Description
    Resume Exit_goto_DendroDetails_Click
    
End Sub
