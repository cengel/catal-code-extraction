

Private Sub cmdChangeStatus_Click()
On Error GoTo err_cmdAddNew_Click

    Me![combostatus].Locked = False
    Me![statusdate].Locked = False
    Me![statuswho].Locked = False

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    Me![statusdate].Value = Now()
    DoCmd.GoToControl Me![unitstatus_determination]
Exit Sub

err_cmdAddNew_Click:
    If Err.Number = 2498 Then
        Resume Next
    Else
    Call General_Error_Trap
    End If
    Exit Sub
End Sub




Private Sub combostatus_BeforeUpdate(Cancel As Integer)
On Error GoTo err_combostatus_BeforeUpdate

If Me![combostatus].Value = "to be checked" And _
IsNull(Forms![Exca: Unit Sheet]![Exca: Unit Data Categories LAYER subform].Form![Data Category].Value) And _
IsNull(Forms![Exca: Unit Sheet]![Exca: Unit Data Categories SKELL subform].Form![Data Category].Value) And _
IsNull(Forms![Exca: Unit Sheet]![Exca: Unit Data Categories CUT subform].Form![Data Category].Value) And _
IsNull(Forms![Exca: Unit Sheet]![Exca: Unit Data Categories CLUSTER subform].Form![Data Category].Value) Then

    'this event will trigger when move to subform, which is really hard
    'on this form as there are so many. So really can only catch this
    'when they are editing at the bottom of the form
   
    MsgBox "There is no Data Category entered for this Unit. This information is mandatory and has to be inserted!" & Chr(13) & Chr(13) & "Click okay, press 'ESC' and enter a valid data category.", vbInformation, "What is the Category?"
    Cancel = True
End If
Exit Sub

err_combostatus_BeforeUpdate:

    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Current()
If Me![status].Value <> "" Then
    Debug.Print okay
    Me![combostatus].Locked = True
    Me![statusdate].Locked = True
    Me![statuswho].Locked = True
Else
End If
End Sub



