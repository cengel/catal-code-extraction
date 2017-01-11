Option Compare Database
Option Explicit

Private Sub Analyst_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_GSAnalyst_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    inputname = InputBox("Please enter the full name of the analyst to go along with these initials:", "Analyst Name")
    If inputname <> "" Then
        sql = "INSERT INTO [GroundStone List of Values: GSAnalyst]([GSAnalystInitials], [GSAnalystName]) VALUES ('" & NewData & "', '" & inputname & "');"
        DoCmd.RunSQL sql
    Else
        Response = acDataErrContinue
    End If
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_GSAnalyst_NotInList:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Artefact_Class_Code_AfterUpdate()
'Artefact Class Text update
On Error GoTo err_arte

    If Me![Artefact Class Code] <> "" Then
    
        If Me![Artefact Class Text] = "" Or IsNull(Me![Artefact Class Text]) Then
            Me![Artefact Class Text] = Me![Artefact Class Code].Column(1)
        ElseIf Me![Artefact Class Text] <> Me![Artefact Class Code].Column(1) Then
            Dim resp
            resp = MsgBox("The artefact class text for this selection is: " & Me![Artefact Class Code].Column(1) & ", differing from that already filled out in the text field (" & Me![Artefact Class Text] & "). Do you wish to update the value field with : " & Me![Artefact Class Code].Column(1) & "?", vbYesNo + vbQuestion, "Artefact Text Mismatch")
            If resp = vbYes Then
                Me![Artefact Class Text] = Me![Artefact Class Code].Column(1)
            End If
        End If
    End If

    Me![Artefact Type Code].Requery
    Me![Subtype 1 Code].Requery
    Me![Artefact Class Text].Requery
    
Exit Sub

err_arte:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Artefact_Type_Code_AfterUpdate()
'Artefact Type Text update
On Error GoTo err_arteType

    If Me![Artefact Type Code] <> "" Then
    
        If Me![Artefact Type Text] = "" Or IsNull(Me![Artefact Type Text]) Then
            Me![Artefact Type Text] = Me![Artefact Type Code].Column(1)
        ElseIf Me![Artefact Type Text] <> Me![Artefact Type Code].Column(1) Then
            Dim resp
            resp = MsgBox("The artefact Type text for this selection is: " & Me![Artefact Type Code].Column(1) & ", differing from that already filled out in the text field (" & Me![Artefact Type Text] & "). Do you wish to update the value field with : " & Me![Artefact Type Code].Column(1) & "?", vbYesNo + vbQuestion, "Artefact Type Mismatch")
            If resp = vbYes Then
                Me![Artefact Type Text] = Me![Artefact Type Code].Column(1)
            End If
        End If
    End If

    Me![Subtype 1 Code].Requery
Exit Sub

err_arteType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Comments_DblClick(Cancel As Integer)
'try to replicate zoom box
On Error GoTo err_comments

    SendKeys "+{F2}", True
Exit Sub

err_comments:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub FindNumber_AfterUpdate()
'update the GID
On Error GoTo err_fn

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_fn:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'*******************************************************************
' Update lastupdated field
' SAJ
'*******************************************************************
On Error GoTo err_Form_BeforeUpdate

Me![Last updated] = Date

Exit Sub

err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'events to keep everything up to date for the current record
On Error GoTo err_current

    'Me![Rock Type].Requery
    'Me![Artefact Type Code].Requery
    'Me![Subtype 1 Code].Requery
    'Me![Artefact Class Text].Requery
    'Me![Material].Requery

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Lettercode_AfterUpdate()
'update the GID
On Error GoTo err_lc

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_lc:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Material_Group_AfterUpdate()
'make sure rocktype gets requeried
On Error GoTo err_mat

    Me![Rock Type].Requery

Exit Sub

err_mat:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub MaterialComments_DblClick(Cancel As Integer)
'try to replicate zoom box
On Error GoTo err_Mcomments

    SendKeys "+{F2}", True
Exit Sub

err_Mcomments:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Rock_Type_AfterUpdate()
'material field appears to hold the rocktypetext value so auto write this in
On Error GoTo err_rocktype

    If Me![Rock Type] <> "" Then
    
        If Me![Material] = "" Or IsNull(Me![Material]) Then
            Me![Material] = Me![Rock Type].Column(1)
        ElseIf Me![Material] <> Me![Rock Type].Column(1) Then
            Dim resp
            resp = MsgBox("The material for this rock type is: " & Me![Rock Type].Column(1) & ", differing from that already filled out in the Material field (" & Me![Material] & "). Do you wish to update the Material field with : " & Me![Rock Type].Column(1) & "?", vbYesNo + vbQuestion, "Material Mismatch")
            If resp = vbYes Then
                Me![Material] = Me![Rock Type].Column(1)
            End If
        End If
    End If

Exit Sub

err_rocktype:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Subtype_1_Code_AfterUpdate()
'SubType 1 Text update
On Error GoTo err_subType

    If Me![Subtype 1 Code] <> "" Then
    
        If Me![Subtype 1 Text] = "" Or IsNull(Me![Subtype 1 Text]) Then
            Me![Subtype 1 Text] = Me![Subtype 1 Code].Column(1)
        ElseIf Me![Subtype 1 Text] <> Me![Subtype 1 Code].Column(1) Then
            Dim resp
            resp = MsgBox("The subtype text for this selection is: " & Me![Subtype 1 Code].Column(1) & ", differing from that already filled out in the text field (" & Me![Subtype 1 Text] & "). Do you wish to update the value field with : " & Me![Subtype 1 Code].Column(1) & "?", vbYesNo + vbQuestion, "Subtype 1 Mismatch")
            If resp = vbYes Then
                Me![Subtype 1 Text] = Me![Subtype 1 Code].Column(1)
            End If
        End If
    End If

    
Exit Sub

err_subType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_AfterUpdate()
'update the GID
On Error GoTo err_unit

    Me![GID] = Me![Unit] & "." & Me![Lettercode] & Me![FindNumber]

Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Unit_DblClick(Cancel As Integer)
On Error GoTo Err_cmdUnitDesc_Click

If Me![Unit] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[Groundstone: Unit Description]", "[Unit] = " & Me![Unit])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [Groundstone: Unit Description] ([Unit]) VALUES (" & Me![Unit] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "Frm_GS_UnitDescription", acNormal, , "[Unit] = " & Me![Unit], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
