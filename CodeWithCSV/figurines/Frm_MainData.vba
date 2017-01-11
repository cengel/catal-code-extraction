Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind

If Me![cboFind] <> "" Then
    'me![txtIdnumber].Enabled = true
    If Me.FilterOn = True Then
        Me.FilterOn = False
        Me.Filter = ""
        Me!cmdRemoveFilter.Visible = False
        Me![lblFilter].Visible = False
    End If
    
    DoCmd.GoToControl Me![txtIDnumber].Name
    DoCmd.FindRecord Me![cboFind]
    Me![cboFind] = ""
    Me![cboFindUnit] = ""
End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_NotInList(NewData As String, Response As Integer)
On Error GoTo err_notUnit
    
    MsgBox "GID number not found", vbInformation, "Not In List"
    Response = acDataErrContinue
    Me![cboFind].Undo

Exit Sub

err_notUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_cboFindUnit

'If Me![cboFindUnit] <> "" Then
'    DoCmd.GoToControl Me![UnitNumber].Name
'    DoCmd.FindRecord Me![cboFindUnit]
'    Me![cboFindUnit] = ""
'End If

If Me![cboFindUnit] <> "" Then
    Me.Filter = "[UnitNumber] = " & Me![cboFindUnit]
    Me.FilterOn = True
    Me![cmdRemoveFilter].Visible = True
    Me!lblFilter.Visible = True
    Me![cboFind] = ""
End If

Exit Sub

err_cboFindUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_NotInList(NewData As String, Response As Integer)
On Error GoTo err_notUnit
    
    MsgBox "Unit number not found", vbInformation, "Not In List"
    Response = acDataErrContinue
    Me![cboFindUnit].Undo

Exit Sub

err_notUnit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub chkFreeStanding_Click()
'update the text in the free standing field
On Error GoTo err_chkFree
    
    If Me![chkFreeStanding] = True Then
        Me![free-standing] = "free-standing"
        
    Else
        Me![free-standing] = ""
    End If

    
Exit Sub

err_chkFree:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub cmdImage1_Click()
Application.FollowHyperlink "http://10.2.1.1/Fig_Images/displayimage.asp?size=full_size&img=" & Me![Image ids], , True
End Sub

Private Sub Close_Click()
'***************************************************
' Standard close
'***************************************************
On Error GoTo err_close_Click
    
    DoCmd.OpenForm "Frm_Menu", , , , acFormPropertySettings
    
    DoCmd.Close acForm, Me.Name
    
    Exit Sub

err_close_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click

    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    Me![txtIDnumber].Locked = False
    DoCmd.GoToControl "txtIDnumber"
Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
On Error GoTo Err_gofirst_Click


    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_gofirst_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdGoNext_Click()
On Error GoTo Err_goNext_Click


    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_goNext_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdGoPrev_Click()
On Error GoTo Err_goPrev_Click


    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_goPrev_Click:
    Call General_Error_Trap
End Sub

Private Sub cmdLastGID_Click()
'return last gid for the current unit
On Error GoTo err_lastgid

    If Me![UnitNumber] = "" Or IsNull(Me![UnitNumber]) Then
        MsgBox "No current Unit to check"
    Else
        Dim num
        num = FindLastGIDforUnit(Me![UnitNumber])
        If num = "Not found" Then
            MsgBox "The last GID number for this unit cannot be found", vbInformation, "No number returned"
        Else
            MsgBox "Last GID entered for this unit was: " & Me![UnitNumber] & "." & "H" & num & Chr(13) & Chr(13) & "Please use " & num + 1 & " as next number.", vbInformation, "Last GID"
        End If
    End If
Exit Sub

err_lastgid:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdLocate_Click()
'allow user to locate this artefact in  crate
On Error GoTo err_locate

    'get find number from main form to pass as openargs
    Dim current
    current = Forms![Frm_MainData]![ID number]
    DoCmd.OpenForm "frm_subform_newlocation", acNormal, , , acFormPropertySettings, acDialog, current
    
Exit Sub

err_locate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'ImageLocationOnSite
On Error GoTo err_current

Dim fullimagepath, imagename

'this will work for on-site if you want to have an image control - one is hidden on this form
'If Me![Image small ids] <> "" Or Not IsNull(Me![Image small ids]) Then
'    If InStr(Me![Image small ids], ".jpg") = 0 Then
'        fullimagepath = ImageLocationOnSite & Me![Image small ids] & ".jpg"
'    Else
'        fullimagepath = ImageLocationOnSite & Me![Image small ids]
''    End If
'
'    Me!Image1.Picture = fullimagepath
'End If

'this will work on site
Me![WebBrowser1].Navigate URL:="http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?img=" & Me![Image small ids] & "&size=thumbnail&height=100&width=100"
Me![WebBrowser2].Navigate URL:="http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?img=" & Me![image 2 small id] & "&size=thumbnail&height=100&width=100"
Me![WebBrowser3].Navigate URL:="http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?img=" & Me![image 3 small id] & "&size=thumbnail&height=100&width=100"
Me![WebBrowser4].Navigate URL:="http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?img=" & Me![image 4 small id] & "&size=thumbnail&height=100&width=100"
'this will work off site
''Me![WebBrowser1].Navigate URL:="http://www.catalhoyuk.com/Fig_Images/displayimage.asp?img=" & Me![Image small ids] & "&size=thumbnail&height=80&width=80"
''MsgBox Me![WebBrowser1].LocationURL

If Me![free-standing] <> "" Then
    Me![chkFreeStanding] = True
Else
    Me![chkFreeStanding] = False
End If

'lock id number so not overwritten
If Me![txtIDnumber] <> "" And Not IsNull(Me![txtIDnumber]) Then
    Me![txtIDnumber].Locked = True
Else
    Me![txtIDnumber].Locked = False
End If

If Me!frm_subform_location.Form.RecordsetClone.RecordCount = 0 Then
    Me!lblCrateRegMsg.Visible = True
    Me!frm_subform_location.Visible = False
    Me![cmdLocate].Visible = True
Else
    Me!lblCrateRegMsg.Visible = False
    Me!frm_subform_location.Visible = True
    Me![cmdLocate].Visible = False
End If
Me!frm_subform_location.Form.Refresh


Exit Sub

err_current:
    If Err.Number = 2220 Then
        'Me!Image1Picture = ""
    Else
        Call General_Error_Trap
    End If
End Sub

Private Sub frmComplete_Click()
On Error GoTo err_frmComplete

    If Me![frmComplete] = 1 Then
        Me![Data Entry] = "complete"
    ElseIf Me![frmComplete] = 2 Then
        Me![Data Entry] = "incomplete"
    Else
        Me![Data Entry] = ""
    End If
    

Exit Sub

err_frmComplete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'make sure focus in search combo
On Error GoTo err_open

    DoCmd.Maximize
    
    DoCmd.GoToControl "cbofind"
    
    'i was going to have it so that all filters get turned off on open BUT
    'finder screen now passes in where/filters so needs to stay on
    If Me.FilterOn = True Then
        Me![cmdRemoveFilter].Visible = True
        If InStr(Me.Filter, "[id number]") > 0 Then
            'Me![cboFind] = Left(Me.Filter, Len(Me.Filter) - 1 - InStr(Me.Filter, "'")) 'WRONG
            Me![cboFind] = Mid(Me.Filter, InStr(Me.Filter, "'") + 1, Len(Me.Filter) - 1 - InStr(Me.Filter, "'"))
        ElseIf InStr(Me.Filter, "unit") > 0 Then
            'Me![cboFindUnit] = Left(Me.Filter, Len(Me.Filter) - InStr(Me.Filter, "="))
            Me![cboFindUnit] = Mid(Me.Filter, InStr(Me.Filter, "=") + 1, Len(Me.Filter) - InStr(Me.Filter, "="))
        End If
        Me!lblFilter.Visible = True
    Else
        Me![cmdRemoveFilter].Visible = False
        Me.Filter = ""
        Me.FilterOn = False
        Me!lblFilter.Visible = False
        Me![cboFindUnit] = ""
        Me![cboFind] = ""
    End If
    
    ''MsgBox Forms!Frm_MainData!UnitNumber

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub image_2_id_Click()
'open hyperlink?

Application.FollowHyperlink "http://10.2.1.1/Fig_Images/displayimage.asp?size=full_size&img=" & Me![image 2 id], , True
End Sub

Private Sub image_2_small_id_AfterUpdate()
'get web browser control to refresh
On Error GoTo err_img_sm_2
   
   Me![WebBrowser2].Navigate URL:="http://10.2.1.1/Fig_Images/displayimage.asp?img=" & Me![image 2 small id] & "&size=thumbnail&height=100&width=100"
err_img_sm_2:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub image_3_id_Click()
'open hyperlink?

Application.FollowHyperlink "http://10.2.1.1/Fig_Images/displayimage.asp?size=full_size&img=" & Me![image 3 id], , True
End Sub

Private Sub image_3_small_id_AfterUpdate()
'get web browser control to refresh
On Error GoTo err_img_sm_3
   
   Me![WebBrowser3].Navigate URL:="http://10.2.1.1/Fig_Images/displayimage.asp?img=" & Me![image 3 small id] & "&size=thumbnail&height=100&width=100"
err_img_sm_3:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub image_4_id_Click()
'open hyperlink?

Application.FollowHyperlink "http://10.2.1.1/Fig_Images/displayimage.asp?size=full_size&img=" & Me![image 4 id], , True
End Sub

Private Sub image_4_small_id_AfterUpdate()
'get web browser control to refresh
On Error GoTo err_img_sm_4
   
   Me![WebBrowser4].Navigate URL:="http://10.2.1.1/Fig_Images/displayimage.asp?img=" & Me![image 4 small id] & "&size=thumbnail&height=100&width=100"

err_img_sm_4:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Image_ids_Click()
'open hyperlink?
'Me![WebBrowser1].Navigate URL:="http://10.2.1.1/Fig_Images/displayimage.asp?img=" & Me![Image small ids] & "&size=thumbnail&height=100&width=100"

Application.FollowHyperlink "http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?size=full_size&img=" & Me![Image ids], , True

End Sub

Private Sub Image_small_ids_AfterUpdate()
'get web browser control to refresh
On Error GoTo err_img_sm_1

    Me![WebBrowser1].Navigate URL:="http://www.catalhoyuk.lan/Fig_Images/displayimage.asp?img=" & Me![Image small ids] & "&size=thumbnail&height=100&width=100"

Exit Sub

err_img_sm_1:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub txtIDnumber_AfterUpdate()
'make sure GID is entered
On Error GoTo err_ID

    If Me![txtIDnumber] = "" Or IsNull(Me![txtIDnumber]) Then
        MsgBox "ID number must be entered", vbCritical, "Missing ID"
        If Me![txtIDnumber].OldValue <> "" Then Me![txtIDnumber] = Me![txtIDnumber].OldValue
        DoCmd.GoToControl "Unitnumber"
        DoCmd.GoToControl "txtIDNumber"
    Else
        'auto fill out unit number field
        Me![UnitNumber] = Left(Me![txtIDnumber], InStr(Me![txtIDnumber], "."))
        
    End If
    
Exit Sub

err_ID:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub txtIDnumber_LostFocus()

'txtIDnumber_AfterUpdate

End Sub
Private Sub cmdRemoveFilter_Click()
On Error GoTo Err_cmdRemoveFilter_Click

    Me.Filter = ""
    Me.FilterOn = False
    Me!cboFindUnit = ""
    DoCmd.GoToControl "cboFindUnit"
    Me![cmdRemoveFilter].Visible = False
    Me!lblFilter.Visible = False
    
    Exit Sub

Err_cmdRemoveFilter_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub
