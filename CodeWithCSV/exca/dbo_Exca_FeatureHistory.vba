

Private Sub cmdChangeStatus_Click()
On Error GoTo err_cmdAddNew_Click
    Me![combostatus].Locked = False
    Me![statusdate].Locked = False
    Me![statuswho].Locked = False
    
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    
    Me![combostatus].Locked = False
    Me![statusdate].Locked = False
    Me![statuswho].Locked = False
    
    Me![statusdate].Value = Now()
    DoCmd.GoToControl Me![featurestatus_determination]
Exit Sub

err_cmdAddNew_Click:
    If Err.Number = 2498 Then
        Resume Next
    Else
    Call General_Error_Trap
    End If
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
