Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
'find combo
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then

    DoCmd.GoToControl "txtUnitNumber"
    DoCmd.FindRecord Me![cboFind]
    Me![cboFind] = ""
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, response As Integer)
'stop not in list msg loop
On Error GoTo err_cbofindNot

    MsgBox "This unit number has not been entered", vbInformation, "No Match"
    response = acDataErrContinue
    
    Me![cboFind].Undo
    DoCmd.GoToControl "cmdAddNew"
Exit Sub

err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'add new unit
On Error GoTo err_cmdAddNew

    DoCmd.GoToRecord acActiveDataObject, Me.Name, acNewRec
    DoCmd.GoToControl "txtUnitNumber"
    

Exit Sub

err_cmdAddNew:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCloseForm_Click()
'close form to return to database window
On Error GoTo err_closeform
    DoCmd.OpenForm "FRM_menu"
    DoCmd.Restore
    DoCmd.Close acForm, Me.Name, acSaveYes
    

Exit Sub

err_closeform:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdDelete_Click()
'new 2009 - control the delete of a record to ensure both tables are clear
On Error GoTo err_del

Dim response
    response = MsgBox("Do you really want to remove Unit " & Me!txtUnitNumber & " and all its related species identification from your database (this does not effect the excavation database)?", vbYesNo + vbQuestion, "Remove Record")
    If response = vbYes Then
        Dim sql
        sql = "Delete FROM [Shell_Level_One_Data] WHERE [UnitNumber] = " & Me![txtUnitNumber] & ";"
        DoCmd.RunSQL sql
        
        sql = "Delete from [Shell_UnitDescription] WHERE [UnitNumber] = " & Me![txtUnitNumber] & ";"
        DoCmd.RunSQL sql
        Me.Requery
        MsgBox "Deletion completed", vbInformation, "Done"
    End If
Exit Sub

err_del:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
'quit application
On Error GoTo err_cmdQuit

    DoCmd.Quit acQuitSaveAll
    
Exit Sub

err_cmdQuit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRefreshCount_Click()
'requery subform to refresh count of bags
On Error GoTo err_refreshcount
    Me![FRM_SUB_TOTAL_BAGS_IN_A_UNIT].Requery
    
Exit Sub

err_refreshcount:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'things to do for the display of each unit
On Error GoTo err_current
    
    If Me!txtUnitNumber <> "" Then
        Me!txtUnitNumber.Locked = True
        Me!txtUnitNumber.BackColor = "-2147483633"
    Else
        Me!txtUnitNumber.Locked = False
        Me!txtUnitNumber.BackColor = "16777215"
    End If

Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'maximise
On Error GoTo err_open
    DoCmd.GoToControl "cboFind"
    DoCmd.Maximize

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtUnitNumber_AfterUpdate()
'new 2009 - update screen after new unit entered
On Error GoTo err_unit

    DoCmd.RunCommand acCmdSaveRecord
    Me!frm_subform_Exca_Unit_Sheet.Requery
    Me!cboFind.Requery
Exit Sub

err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub
