Option Compare Database
Option Explicit

Private Sub cboFindFlot_AfterUpdate()
'******************************************************************
' Search for a flot number from the list
' SAJ
'******************************************************************
On Error GoTo err_FindFlot

If Me![cboFindFlot] <> "" Then
    DoCmd.GoToControl "FrmSubBasicData"
    DoCmd.GoToControl "Flot Number"
    'DoCmd.GoToControl Me!FrmSubBasicData.Form![Flot Number].Name
    DoCmd.FindRecord Me![cboFindFlot]
End If

Exit Sub

err_FindFlot:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddRange_Click()
'new in 2006 - request from Nikki to add range of number automatically
'july 2007 get them to add the float date
On Error GoTo err_range

    Dim startnum, endnum, sql, sql1, floatdate
    startnum = InputBox("Please enter the number that starts this set of Flot numbers", "Start Number")
    If startnum <> "" Then
        endnum = InputBox("Please enter the number that ends this set of Flot numbers", "End Number")
        If endnum <> "" Then
            
            floatdate = InputBox("Please enter the Flot date for these records (dd/mm/yyyy)", "Float Date")
            If floatdate <> "" Then
            
                If startnum < endnum Then
                    Do While CInt(startnum) <= CInt(endnum)
                        'sql = "INSERT INTO [Bot: Basic Data] ([Flot Number], [Float date]) VALUES (" & startnum & ", #" & Date & "#);"
                        sql = "INSERT INTO [Bot: Basic Data] ([Flot Number], [Float date]) VALUES (" & startnum & ", #" & floatdate & "#);"
                        DoCmd.RunSQL sql
                    
                        'sql1 = "INSERT INTO [Bot: Sample Scanning] ([Flot Number], [YearScanned]) VALUES (" & startnum & ", " & Year(Date) & ");"
                        sql1 = "INSERT INTO [Bot: Sample Scanning] ([Flot Number], [YearScanned]) VALUES (" & startnum & ", " & Year(floatdate) & ");"
                        DoCmd.RunSQL sql1
                        startnum = startnum + 1
                    Loop
                    Me!FrmSubBasicData.Requery
                    DoCmd.GoToControl "FrmSubBasicData"
                    DoCmd.GoToControl "Flot Number"
                    DoCmd.GoToRecord acActiveDataObject, , acLast
                Else
                    MsgBox "Invalid number range the start number is greater than the end number, please try again", vbInformation, "Invalid Action"
                    Exit Sub
                End If
            Else
                MsgBox "Sorry this function only works if a start and end number are entered, please use the new record button instead", vbExclamation, "No end number entered"
            End If
        End If
    End If
Exit Sub

err_range:
    If Err.Number = 2501 Then
        MsgBox "An error has occured, the record you were trying to enter probably already exisits", vbInformation, "Error"
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub cmdClose_Click()
'********************************************************************
' Close form and return to main menu
' SAJ
'********************************************************************
On Error GoTo err_close
    DoCmd.OpenForm "FrmMainMenu"
    DoCmd.Close acForm, "FrmBasicData"

Exit Sub
err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
'********************************************************************
' Go to first record
' SAJ
'********************************************************************
On Error GoTo Err_cmdgofirst_Click

    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Name
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

    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Name
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

    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Name
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

    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Name
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

    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Name
    DoCmd.GoToRecord , , acNewRec
    DoCmd.GoToControl Forms![FrmBasicData]![FrmSubBasicData].Form![Flot Number].Name

    Exit Sub

Err_cmdgonew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdHistoric_Click()
'Open the old bots interface
On Error GoTo err_Historic
    DoCmd.OpenForm "Bots98: Flot Sheet", , , "[FLot Number] = " & Me![FrmSubBasicData].Form![Flot Number], acFormReadOnly

Exit Sub

err_Historic:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdOutput_Click()
'open output options pop up
On Error GoTo err_Output

    If Me![FrmSubBasicData].Form![Flot Number] <> "" Then
        'DoCmd.OpenForm "FrmPopDataOutputOptions", acNormal, , , , acDialog, "Bot: Basic Data;" & Me![FrmSubBasicData].Form![Flot Number]
        'Q_ExportBasicData_AllRecs_withUnit
        'will this be too slow - but request from A & M to have excavation data as well
        DoCmd.OpenForm "FrmPopDataOutputOptions", acNormal, , , , acDialog, "Q_ExportBasicData_AllRecs_withUnit;" & Me![FrmSubBasicData].Form![Flot Number]
    Else
        MsgBox "The output options form cannot be shown when there is no Flot Number on screen", vbInformation, "Action Cancelled"
    End If

Exit Sub

err_Output:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdUnitDescr_Click()
On Error GoTo Err_cmdUnitDesc_Click

If Me![FrmSubBasicData].Form![Unit Number] <> "" Then
    'check the unit number is in the unit desc form
    Dim checknum, sql
    checknum = DLookup("[Unit]", "[Bot: Unit Description]", "[Unit] = " & Me![FrmSubBasicData].Form![Unit Number])
    If IsNull(checknum) Then
        'must add the unit to the table
        sql = "INSERT INTo [Bot: Unit Description] ([Unit]) VALUES (" & Me![FrmSubBasicData].Form![Unit Number] & ");"
        DoCmd.RunSQL sql
    End If
    
    DoCmd.OpenForm "FrmBotUnitDescription", acNormal, , "[Unit] = " & Me![FrmSubBasicData].Form![Unit Number], acFormPropertySettings
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub

Err_cmdUnitDesc_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
'********************************************************************
' Set up how the form should look
' SAJ
'********************************************************************
On Error GoTo err_open

If Not IsNull(Me.OpenArgs) Then
    'flot number passed in must find it
    DoCmd.GoToControl "FrmSubBasicData"
    DoCmd.GoToControl "Flot Number"
    DoCmd.FindRecord Me.OpenArgs
    DoCmd.GoToControl "Sample Number"
End If

If Me!FrmSubBasicData.Form.DefaultView = 2 Then
    'if the default view of the subform is datasheet then set the datasheet
    'button to true, the cbo is shown for the datasheet
    Me!tglDataSheet = True
    Me!tglFormV = False
    Me!FrmSubBasicData.Form!cboOptions.Visible = True
Else
    'if the default view of the subform is form then set the datasheet
    'button to false, the cbo is hidden for the form
    Me!tglDataSheet = False
    Me!tglFormV = True
    Me!FrmSubBasicData.Form!cboOptions.Visible = False
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
    'set the sub form to datasheet view and make the combo of actions visible
    'as this replaces the form action buttons visible in form view
    Me!FrmSubBasicData.SetFocus
    Me!FrmSubBasicData.Form![cboOptions].Visible = True
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
    'set the sub form to form view and make the combo of actions invisible
    'as this is replaced by the form action buttons when in form view
    Me!FrmSubBasicData.SetFocus
    Me!FrmSubBasicData![Flot Number].SetFocus
    Me!FrmSubBasicData.Form![cboOptions].Visible = False
    DoCmd.RunCommand acCmdSubformDatasheet
    Me!tglDataSheet = False
End If

Exit Sub

Err_tglFormV:
    Call General_Error_Trap
    Exit Sub
End Sub
