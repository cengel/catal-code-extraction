Option Compare Database
Option Explicit

Private Sub cmdAddtoPReport_Click()
'check if the record already exists in priority report
'and if not insert it and do calculation required
'SAJ season 2006
On Error GoTo err_AddtoPReport

    Dim checknum, sql, woodml, dungml, parencml, seedchaff
    checknum = DLookup("[Flot Number]", "[Bot: Priority Report]", "[Flot Number] = " & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number])
    If IsNull(checknum) Then
        sql = "INSERT INTO [Bot: Priority Report] ([Flot Number]) VALUES (" & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number] & ");"
        DoCmd.RunSQL sql
        'calc the values required
        woodml = ""
        If Forms![FrmPriority]![FrmSubPriority].Form![4 mm wood] <> "" And Not IsNull(Forms![FrmPriority]![FrmSubPriority].Form![4 mm wood]) Then
            
            woodml = Calc_WoodParenceDung_ml_per_litre(Forms![FrmPriority]![FrmSubPriority].Form![Flot Number], Forms![FrmPriority]![FrmSubPriority].Form![4 mm wood], Forms![FrmPriority]![FrmSubPriority].Form![4 mm fraction])
            If woodml <> "" Then
                woodml = Round(woodml, 2)
                sql = "UPDATE [Bot: Priority Report] SET [Wood_ml_Per_Litre] = " & woodml & " WHERE [Flot Number] = " & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number] & ";"
                DoCmd.RunSQL sql
            End If
        End If
        
        dungml = ""
        If Forms![FrmPriority]![FrmSubPriority].Form![4 mm dung] <> "" And Not IsNull(Forms![FrmPriority]![FrmSubPriority].Form![4 mm dung]) Then
            dungml = Calc_WoodParenceDung_ml_per_litre(Forms![FrmPriority]![FrmSubPriority].Form![Flot Number], Forms![FrmPriority]![FrmSubPriority].Form![4 mm dung], Forms![FrmPriority]![FrmSubPriority].Form![4 mm fraction])
        
            If dungml <> "" Then
                dungml = Round(dungml, 2)
                sql = "UPDATE [Bot: Priority Report] SET [Dung_ml_Per_Litre]= " & dungml & " WHERE [Flot Number] = " & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number] & ";"
                DoCmd.RunSQL sql
            End If
        End If
        
        parencml = ""
        If Forms![FrmPriority]![FrmSubPriority].Form![4 mm parenc] <> "" And Not IsNull(Forms![FrmPriority]![FrmSubPriority].Form![4 mm parenc]) Then
            parencml = Calc_WoodParenceDung_ml_per_litre(Forms![FrmPriority]![FrmSubPriority].Form![Flot Number], Forms![FrmPriority]![FrmSubPriority].Form![4 mm parenc], Forms![FrmPriority]![FrmSubPriority].Form![4 mm fraction])
    
            If parencml <> "" Then
                parencml = Round(parencml, 2)
                sql = "UPDATE [Bot: Priority Report] SET [Parenc_ml_Per_Litre] = " & parencml & " WHERE [Flot Number] = " & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number] & ";"
                DoCmd.RunSQL sql
            End If
        End If
        
        seedchaff = ""
        seedchaff = Calc_seedchaff_per_litre(Forms![FrmPriority]![FrmSubPriority].Form![Flot Number])
        If seedchaff <> "" Then
            seedchaff = Round(seedchaff, 2)
            sql = "UPDATE [Bot: Priority Report] SET [Seeds_Chaff_Per_Litre] = " & seedchaff & " WHERE [Flot Number] = " & Forms![FrmPriority]![FrmSubPriority].Form![Flot Number] & ";"
            DoCmd.RunSQL sql
        End If
        
        
        If woodml = "" And parencml = "" And dungml = "" And seedchaff = "" Then
            MsgBox "The system attempted to undetake the necessary calculations automatically but one or more of the necessary fields was missing", vbInformation, "Auto Calculate not performed"
        End If
    
    End If

    DoCmd.OpenForm "FrmPriorityReport", acNormal, , , , , Forms![FrmPriority]![FrmSubPriority].Form![Flot Number]
Exit Sub

err_AddtoPReport:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Name
    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_cmdgofirst_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoLast_Click()
'********************************************************************
' Go to last record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoLast_Click

    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Name
    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_cmdgoLast_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoNext_Click()
'********************************************************************
' Go to next record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoNext_Click

    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Name
    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_cmdgoNext_Click:
    If Err.Number = 2105 Then
        MsgBox "No more records to show", vbInformation, "End of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdgoprevious_Click()
'********************************************************************
' Go to previous record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgoprevious_Click

    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Name
    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_cmdgoprevious_Click:
    If Err.Number = 2105 Then
        MsgBox "Already at the begining of the recordset", vbInformation, "Begining of records"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdGoToBasic_Click()
'*******************************************************************
' Open Basic data screen at selected flot number
' SAJ
'*******************************************************************
On Error GoTo err_gotoBasic
Dim openarg

If Me![FrmSubPriority].Form![Flot Number] <> "" Then
    openarg = Me![FrmSubPriority].Form![Flot Number]
Else
    openarg = Null
End If

DoCmd.OpenForm "FrmBasicData", acNormal, , , , , openarg
DoCmd.Close acForm, "FrmPriority"
Exit Sub

err_gotoBasic:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdgotonew_Click()
'********************************************************************
' Create new record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgonew_Click

    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Name
    DoCmd.GoToRecord , , acNewRec
    DoCmd.GoToControl Forms![FrmPriority]![FrmSubPriority].Form![Flot Number].Name

    Exit Sub

Err_cmdgonew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cboFindFlot_AfterUpdate()
'******************************************************************
' Search for a flot number from the list
' SAJ
'******************************************************************
On Error GoTo err_FindFlot

If Me![cboFindFlot] <> "" Then

    DoCmd.GoToControl "FrmSubPriority"
    DoCmd.GoToControl "Flot Number"
    'DoCmd.GoToControl Me!FrmSubBasicData.Form![Flot Number].Name
    DoCmd.FindRecord Me![cboFindFlot]
    DoCmd.GoToControl "4 mm fraction"
End If

Exit Sub

err_FindFlot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
'********************************************************************
' Close form and return to main menu
' SAJ
'********************************************************************
On Error GoTo err_close
    DoCmd.OpenForm "FrmMainMenu"
    DoCmd.Close acForm, "FrmPriority"

Exit Sub
err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![FrmSubPriority].Form![Flot Number] <> "" Then
        DoCmd.OpenForm "FrmPopDataOutputOptions", acNormal, , , , acDialog, "Bot: Priority Sample;" & Me![FrmSubPriority].Form![Flot Number]
    Else
        MsgBox "The output options form cannot be shown when there is no Flot Number on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Command22_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
'*****************************************************************************
' Check for any open args to set record to focus on and set up view of form
'
' SAJ
'*****************************************************************************
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    'flot number passed in must find it
    DoCmd.GoToControl "FrmSubPriority"
    DoCmd.GoToControl "Flot Number"
    DoCmd.FindRecord Me.OpenArgs
    DoCmd.GoToControl "4 mm fraction"
End If

If Me!FrmSubPriority.Form.DefaultView = 2 Then
    Me!tglDataSheet = True
    Me!tglFormV = False
    'Me!FrmSubBasicData.Form!cboOptions.Visible = True
Else
    Me!tglDataSheet = False
    Me!tglFormV = True
    'Me!FrmSubBasicData.Form!cboOptions.Visible = False
End If
    
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub tglDataSheet_Click()
'********************************************************************
' The user wants to see the basic data in datasheet view
' SAJ
'********************************************************************
On Error GoTo Err_tglDataSheet

If Me!tglDataSheet = True Then
    Me!FrmSubPriority.SetFocus
    'Me!FrmSubBasicData.Form![cboOptions].Visible = True
    DoCmd.RunCommand acCmdSubformDatasheet
    Me!tglFormV = False
End If
Exit Sub

Err_tglDataSheet:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub tglFormV_Click()
'********************************************************************
' The user wants to see the basic data in form view
' SAJ
'********************************************************************
On Error GoTo Err_tglFormV

If Me!tglFormV = True Then
    
    Me!FrmSubPriority.SetFocus
    Me!FrmSubPriority![Flot Number].SetFocus
    DoCmd.RunCommand acCmdSubformDatasheet

    'Me!FrmSubBasicData.Form.DefaultView = 2
    Me!tglDataSheet = False
End If
Exit Sub

Err_tglFormV:
    Call General_Error_Trap
    Exit Sub
End Sub
