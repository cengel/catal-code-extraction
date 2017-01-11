Option Compare Database
Option Explicit

Sub copy_layer_Click()
On Error GoTo Err_copy_layer_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Copy Layer description"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_copy_layer_Click:
    Exit Sub

Err_copy_layer_Click:
    MsgBox Err.Description
    Resume Exit_copy_layer_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'*****************************************************************
' Setup display parametres based on permissions
' SAJ v9.1
'*****************************************************************
On Error GoTo err_Form_Open

Dim permiss
permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
        Me![copy layer].Enabled = False
        'in toggleformreadonly allowadditions is set to false and the fields are not showing
        'for records where no data - just a blank grey area where subform should be
        'data cannot be entered anyway due to set up of main form so set allowadditions to
        'true to at least show subform fields - this doesn't happen with all subforms eg: dimensions
        'must be something to do with the type of controls here
        If Me.AllowAdditions = False Then Me.AllowAdditions = True
    End If

Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
