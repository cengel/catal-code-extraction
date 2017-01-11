Option Compare Database
Option Explicit

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    DoCmd.GoToControl "txtFlotNumber"
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

    DoCmd.GoToControl "txtFlotNumber"
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

    DoCmd.GoToControl "txtFlotNumber"
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

    DoCmd.GoToControl "txtFlotNumber"
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

Private Sub cmdgotonew_Click()
'********************************************************************
' Create new record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgonew_Click

    DoCmd.GoToRecord , , acNewRec
    DoCmd.GoToControl "txtFlotNumber"

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

    DoCmd.GoToControl "txtFlotNumber"
    DoCmd.FindRecord Me![cboFindFlot]
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
    DoCmd.Close acForm, Me.Name
    

Exit Sub
err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![txtFlotNumber] <> "" Then
        DoCmd.OpenForm "FrmPopDataOutputOptions", acNormal, , , , acDialog, "Bot: Priority Report;" & Me![txtFlotNumber]
    Else
        MsgBox "The output options form cannot be shown when there is no Flot Number on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRecalc_Click()
'*******************************************************************
' Recalc the wood, parenc, dung and seed/chaff values
' SAJ
'*******************************************************************
On Error GoTo err_cmdRecalc
    
    Dim getFourmmFraction, getWood, getParenc, getDung, result1, result2, result3, result4
    getFourmmFraction = DLookup("[4 mm Fraction]", "[Bot: Priority Sample]", "[Flot Number] = " & Me![txtFlotNumber])
    If Not IsNull(getFourmmFraction) Then
        'calc the values required
        getWood = DLookup("[4 mm Wood]", "[Bot: Priority Sample]", "[Flot Number] = " & Me![txtFlotNumber])
        If Not IsNull(getWood) Then
            result1 = Calc_WoodParenceDung_ml_per_litre(Me![txtFlotNumber], getWood, getFourmmFraction)
            Me![Wood_ml_Per_Litre] = Round(result1, 2)
        End If
    
        getParenc = DLookup("[4 mm Parenc]", "[Bot: Priority Sample]", "[Flot Number] = " & Me![txtFlotNumber])
        If Not IsNull(getParenc) Then
            result2 = Calc_WoodParenceDung_ml_per_litre(Me![txtFlotNumber], getParenc, getFourmmFraction)
            Me![Parenc_ml_Per_Litre] = Round(result2, 2)
        End If
    
        getDung = DLookup("[4 mm Dung]", "[Bot: Priority Sample]", "[Flot Number] = " & Me![txtFlotNumber])
        If Not IsNull(getDung) Then
            result3 = Calc_WoodParenceDung_ml_per_litre(Me![txtFlotNumber], getDung, getFourmmFraction)
            Me![Dung_ml_Per_Litre] = Round(result3, 2)
        End If
        
        result4 = Calc_seedchaff_per_litre(Me![txtFlotNumber])
        Me![Seeds_Chaff_Per_Litre] = Round(result4, 2)
        
    Else
        MsgBox "The system cannot obtain the 4mm fraction value so cannot recalculate the fields", vbCritical, "Error Obtaining Fraction"
    End If
Exit Sub

err_cmdRecalc:
    Call General_Error_Trap
    Exit Sub
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
    DoCmd.GoToControl "txtFlotNumber"
    DoCmd.FindRecord Me.OpenArgs
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
'On Error GoTo Err_tglDataSheet
'
'If Me!tglDataSheet = True Then
'    Me!FrmSubPriority.SetFocus
'    'Me!FrmSubBasicData.Form![cboOptions].Visible = True
'    DoCmd.RunCommand acCmdSubformDatasheet
'    Me!tglFormV = False
'End If
'Exit Sub

'Err_tglDataSheet:
'    Call General_Error_Trap
'    Exit Sub
End Sub

Private Sub tglFormV_Click()
'********************************************************************
' The user wants to see the basic data in form view
' SAJ
'********************************************************************
'On Error GoTo Err_tglFormV

'If Me!tglFormV = True Then
'
'    Me!FrmSubPriority.SetFocus
'    Me!FrmSubPriority![Flot Number].SetFocus
'    DoCmd.RunCommand acCmdSubformDatasheet

'    'Me!FrmSubBasicData.Form.DefaultView = 2
'    Me!tglDataSheet = False
'End If
'Exit Sub

'Err_tglFormV:
'    Call General_Error_Trap
'    Exit Sub
End Sub
