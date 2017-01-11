Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
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

Sub open_skell_Click()
'****************************************************************************
' This button triggers a parametre box to appear for a feature number  - its predates
' the time features were normalised and pulled out of the Unit table. Then it leads
' off to a skeleton recording sheet.
' SF said to hide this button as everything is now recorded on the main Unit form
' SAJ v9.1
'****************************************************************************
On Error GoTo Err_open_skell_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Skeleton Sheet"
    
    stLinkCriteria = "[Unit Number]=" & Me![To_Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_skell_Click:
    Exit Sub

Err_open_skell_Click:
    MsgBox Err.Description
    Resume Exit_open_skell_Click
    
End Sub

Private Sub To_Unit_AfterUpdate()
'***********************************************************************
' Intro of a validity check to ensure unit num entered here is exists
' that it is a skeleton
'
' SAJ v9.1
'***********************************************************************
On Error GoTo err_To_Unit_AfterUpdate

Dim checknum, msg, retval, checknum2

If Me![To_Unit] <> "" Then
    'first check its valid
    If IsNumeric(Me![To_Unit]) Then
    
        'check that Unit num does exist
        checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![To_Unit])
        If IsNull(checknum) Then
            msg = "This Unit Number DOES NOT EXIST in the database yet, please ensure it is entered soon."
            MsgBox msg, vbInformation, "Unit Number does not exist yet"
           DoCmd.GoToControl "To_Unit"
            
        Else
            'valid number, now check its  category
            checknum2 = DLookup("[Category]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![To_Unit])
                If Not IsNull(checknum2) Then 'category found this unit
                    If UCase(checknum2) <> "SKELETON" Then
                        'do not allow entry if units category is not skeleton
                        msg = "This entry is not allowed:  Unit (" & Me![To_Unit] & ")"
                        msg = msg & " has the category " & checknum2 & ", only Units with the category 'Skeleton' are valid here."
                        msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
                        MsgBox msg, vbExclamation, "Category problem"
                        
                        'reset val to previous val if is one or else remove it completely
                        If Not IsNull(Me![To_Unit].OldValue) Then
                            Me![To_Unit] = Me![To_Unit].OldValue
                        Else
                            Me.Undo
                        End If
                        DoCmd.GoToControl "To_Unit"
                    End If
                Else
                    'the category for this unit has not been filled out yet, SF says allow link
                    msg = "The Unit (" & Me![To_Unit] & ")"
                    msg = msg & " has no category entered yet. Please correct this as soon as possible"
                    MsgBox msg, vbInformation, "Category Missing"
                    'but do nothign
                    DoCmd.GoToControl "To_Unit"
                End If
        End If
    
    Else
        'not a vaild numeric unit number
        MsgBox "The Unit number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If

Exit Sub

err_To_Unit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
