Option Compare Database
Option Explicit
'**************************************************************************
' This form has been reformatted - only allow entry of priority data, no
' editing of data from unit sheet allowed. Formatted to show what can be
' edited.
'
' SAJ v9.1
'***************************************************************************

Private Sub cboFindPriority_AfterUpdate()
'********************************************
'Find the selected priority unit number from the list
'
'SAJ v9.1
'********************************************
On Error GoTo err_cboFindPriority_AfterUpdate

    If Me![cboFindPriority] <> "" Then
        'if field disabled, enable it for find then disable again
        If Me![Unit Number].Enabled = False Then Me![Unit Number].Enabled = True
        DoCmd.GoToControl "Unit Number"
        DoCmd.FindRecord Me![cboFindPriority]
        'send focus down to the main editable field here - enabling if necessary
        'If Me![Short Description].Enabled = False Then Me![Short Description].Enabled = True
        'DoCmd.GoToControl "Short Description"
        '2009 - no don't as it might get overwritten by mistake so keep focus here
        DoCmd.GoToControl "cboFindPriority"
        Me![Unit Number].Enabled = False
        Me![cboFindPriority] = ""
    End If
Exit Sub

err_cboFindPriority_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdClose_Click()
'***************************************************************
' original closing form pic replaced by common trowel close
' and gen error trap intro, clse specified as this form - rest is orig code
' SAJ v9.1
'***************************************************************
On Error GoTo err_close_Click


    DoCmd.Close acForm, "Exca: Priority Detail", acSaveYes

Exit_close_Click:
    Exit Sub

err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
    
End Sub



Private Sub Form_Open(Cancel As Integer)
'***************************************************************
' New permissions check
' SAJ v9.1
'***************************************************************
On Error GoTo err_Form_Open
Dim permiss
    permiss = GetGeneralPermissions
    'due to amount of field always locked on this form not going to use togglformreadonly
    'but set it here
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ''ToggleFormReadOnly Me, False
        Me![Priority].Locked = False
        Me![Priority].Enabled = True
        Me![Priority].BackColor = 16777215
        Me![Discussion].Locked = False
        Me![Discussion].Enabled = True
        Me![Discussion].BackColor = 16777215
        Me![Short Description].Locked = False
        Me![Short Description].Enabled = True
        Me![Short Description].BackColor = 16777215
    Else
        'set read only form here, just once
        ''ToggleFormReadOnly Me, True
        Me![Priority].Locked = True
        Me![Priority].Enabled = False
        Me![Priority].BackColor = Me.Section(0).BackColor
        Me![Discussion].Locked = True
        Me![Discussion].Enabled = False
        Me![Discussion].BackColor = Me.Section(0).BackColor
        Me![Short Description].Locked = True
        Me![Short Description].Enabled = False
        Me![Short Description].BackColor = Me.Section(0).BackColor
    End If

    'new 2009 to ensure when opened from unit sheet it disables the search as filter is on
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        'disable find and add new in this instance
        Me![cboFindPriority].Enabled = False
        Me.AllowAdditions = False
        DoCmd.GoToControl "cmdClose"
        
    Else
        'new end of season 2008 to ensure not within record when opened, try and prevent overwriting
        DoCmd.GoToControl "cboFindPriority"
    End If

Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub prevrep_Click()
'***************************************************************
' preview report existing code - gen error trap intro
' SAJ v9.1 (BTW the report has not been touched in this version)
'***************************************************************
On Error GoTo Err_prevrep_Click

    Dim stDocName As String

    stDocName = "Exca: Priority Units"
    DoCmd.OpenReport stDocName, acPreview

Exit_prevrep_Click:
    Exit Sub

Err_prevrep_Click:
    Call General_Error_Trap
    Resume Exit_prevrep_Click
    
End Sub
