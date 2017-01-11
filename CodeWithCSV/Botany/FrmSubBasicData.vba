Option Compare Database
Option Explicit

'saj july 2007, carry the last date input for all subsequent records
Dim dateTouse




Private Sub cboOptions_AfterUpdate()
'*******************************************
' The form is datasheet view which means no buttons
' are available, so the action list takes their place.
' This replicates the actions of the buttons that are shown
' in form view
' SAJ
'*******************************************
' Something very odd is going on in this procedure - it replicates the code
' used by cmdAddPriority where this form is then closed. However if you try
' to close the form here it quits the whole application. I have no idea why
' so here this form is left open.
'*******************************************
On Error GoTo err_cboOptions_click
    If Me![cboOptions] = "Put record in Priority" Then
        If AddRecordToPriorityTable(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name 'do not do this will quit app
        End If
    ElseIf Me![cboOptions] = "View Priority Record" Then
        If ViewPriorityRecord(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name 'do not do this will quit app
        End If
    ElseIf Me![cboOptions] = "View Scanning" Then
        If ViewScanRecord(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name 'do not do this will quit app
        End If
    ElseIf Me![cboOptions] = "Put record in Scanning" Then
        If AddRecordToScanTable(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name
        End If
    ElseIf Me![cboOptions] = "View Sieve Scanning" Then
        If ViewSieveScanRecord(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name 'do not do this will quit app
        End If
    ElseIf Me![cboOptions] = "Put record in Sieve Scanning" Then
        If AddRecordToSieveScanTable(Me) = True Then
            ''DoCmd.Close acForm, Forms![frmBasicData].Name
        End If
    Else
        MsgBox "Action not known to the system", vbCritical, "Unknown Action"
    End If
    
Exit Sub

err_cboOptions_click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddPriority_Click()
'*******************************************
' Add the flot record to the priority table
' SAJ
'*******************************************
On Error GoTo err_Priority_click

If AddRecordToPriorityTable(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_Priority_click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddScan_Click()
'*******************************************
' Add the flot record to the scanning table
' SAJ
'*******************************************
On Error GoTo err_Scan_click

If AddRecordToScanTable(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_Scan_click:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cmdAddSieveScan_Click()
'*******************************************
' Add the flot record to the scanning table
' SAJ
'*******************************************
On Error GoTo err_Scan_click

If AddRecordToSieveScanTable(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_Scan_click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdExportBasicOnly_Click()
'export record withOUT unit data as this is faster
'functionality replace by cmdOutput on FrmBasicData
On Error GoTo err_ExportRec
    'DoCmd.OutputTo acOutputQuery, "Q_ExportBasicData_FormRecord", acFormatXLS, , True
    'DoCmd.OutputTo acOutputQuery, "Q_ExportBasicData_FormRecord", acFormatXLS, , True
    DoCmd.OutputTo acOutputForm, Me, acFormatXLS, , True
Exit Sub

err_ExportRec:
    If Err.Number = 2302 Then
        MsgBox "The data cannot be exported at present. Please close all Excel files you may have open and try again", vbInformation, "Error"
    Else
        MsgBox "An error has occured trying to export the record to Excel,  the message is as follows: " & Err.Description
    End If
Exit Sub

End Sub

Private Sub cmdExportRecToExcel_Click()
'export record with unit data as well - can be quite slow
'functionality replace by cmdOutput on FrmBasicData
On Error GoTo err_ExportRec
    'DoCmd.OutputTo acOutputQuery, "Q_ExportBasicData_FormRecord", acFormatXLS, , True
    DoCmd.OutputTo acOutputQuery, "Q_ExportBasicData_FormRecord_withUnit", acFormatXLS, , True
Exit Sub

err_ExportRec:
    If Err.Number = 2302 Then
        MsgBox "The data cannot be exported at present. Please close all Excel files you may have open and try again", vbInformation, "Error"
    Else
        MsgBox "An error has occured trying to export the record to Excel,  the message is as follows: " & Err.Description
    End If
Exit Sub
End Sub

Private Sub cmdGotoPriority_Click()
'*******************************************
' Open the priority record for this the flot num
' SAJ
'*******************************************
On Error GoTo err_GoToPriority_click


If ViewPriorityRecord(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_GoToPriority_click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdGoToScanning_Click()
'*******************************************
' Open the sample scan record for this the flot num
' SAJ
'*******************************************
On Error GoTo err_GoToScanning_click

If ViewScanRecord(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_GoToScanning_click:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cmdGoToSieveScanning_Click()
'*******************************************
' Open the sample scan record for this the flot num
' SAJ
'*******************************************
On Error GoTo err_GoToSieveScanning_click

If ViewSieveScanRecord(Me) = True Then
    'closing form here causes no problem (see cboOptions_AfterUpdate for more info)
    DoCmd.Close acForm, Forms![FrmBasicData].Name
End If

Exit Sub
err_GoToSieveScanning_click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Float_date_AfterUpdate()
'saj july 2007 - grab the date entered and allow it to be used for next record
On Error GoTo err_floatdate

    
    dateTouse = Me![Float Date]
    

Exit Sub

err_floatdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Floater_NotInList(NewData As String, Response As Integer)
'*******************************************
' allow new floater names to be added after a prompt
' SAJ
'*******************************************
On Error GoTo err_floater
Dim retVal
retVal = MsgBox("This floater name has not been used on this screen before, are you sure you wish to add it to the list?", vbInformation + vbYesNo, "New floater name")
If retVal = vbYes Then
    'allow value, as this is distinct query based list we must save the record
    'first but need to turn off limittolist first to be able to do so an alternative
    'way to do this would be to dlookup on entry when not limited
    'to list but this method is quicker (but messier) as not require DB lookup 1st
    Response = acDataErrContinue
    Me![Floater].LimitToList = False 'turn off limit to list so record can be saved
    DoCmd.RunCommand acCmdSaveRecord 'save rec
    Me![Floater].Requery 'requery combo to get new value in list
    Me![Floater].LimitToList = True 'put back on limit to list
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub
err_floater:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Current()
'**********************************************************************************
' Set up display of the form based on whether the flot number has related records
' in the priority and sample tables. The command buttons appear on the form view and
' the combo cboOptions appears on the datasheet view to offer the same functionality as
' the buttons in this view.
' NB - the optionrowsource values are referred to in cboOptions_Afterupdate and form_open
' so any changes must be reflected there as well
' SAJ
'***********************************************************************************
On Error GoTo err_current

Dim checkFlotInPriority, checkFlotInScan, checkFlotInSieveScan
Dim OptionRowSource

If Not IsNull(Me![Flot Number]) Then
    checkFlotInPriority = DLookup("[Flot Number]", "Bot: Priority Sample", "[Flot Number] = " & Me![Flot Number])
    If IsNull(checkFlotInPriority) Then
    '    Me![chkInPriority] = False
        OptionRowSource = OptionRowSource & "Put record in Priority;"
        Me![cmdAddPriority].Visible = True
        Me![cmdGotoPriority].Visible = False
    Else
    '    Me![chkInPriority] = True
        OptionRowSource = OptionRowSource & "View Priority Record;"
        Me![cmdGotoPriority].Visible = True
        Me![cmdAddPriority].Visible = False
    End If

    checkFlotInScan = DLookup("[Flot Number]", "Bot: Sample Scanning", "[Flot Number] = " & Me![Flot Number])
    If IsNull(checkFlotInScan) Then
    '    Me![chkInScanning] = False
        OptionRowSource = OptionRowSource & "Put record in Scanning;"
        Me![cmdAddScan].Visible = True
        Me![cmdGoToScanning].Visible = False
    Else
    '    Me![chkInScanning] = True
        OptionRowSource = OptionRowSource & "View Scanning;"
        Me![cmdAddScan].Visible = False
        Me![cmdGoToScanning].Visible = True
    End If
    
    checkFlotInSieveScan = DLookup("[Flot Number]", "Bot: SieveScanning", "[Flot Number] = " & Me![Flot Number])
    If IsNull(checkFlotInSieveScan) Then
    '    Me![chkInSieveScanning] = False
        OptionRowSource = OptionRowSource & "Put record in Sieve Scanning;"
        Me![cmdAddSieveScan].Visible = True
        Me![cmdGoToSieveScanning].Visible = False
    Else
    '    Me![chkInSieveScanning] = True
        OptionRowSource = OptionRowSource & "View Sieve Scanning;"
        Me![cmdAddSieveScan].Visible = False
        Me![cmdGoToSieveScanning].Visible = True
    End If
    Me!cboOptions.RowSource = OptionRowSource
    
    'check for historic data for this flot number
    Dim checknum
    checknum = DLookup("[Flot Number]", "Bots98: Basic Flot Details", "[Flot Number] = " & Me![Flot Number])
    If Not IsNull(checknum) Then
        Forms![FrmBasicData]![cmdHistoric].Enabled = True
    Else
        Forms![FrmBasicData]![cmdHistoric].Enabled = False
    End If

Else
    'no flot number - new record
    Me![cmdAddPriority].Visible = True
    Me![cmdGotoPriority].Visible = False
    Me![cmdAddScan].Visible = True
    Me![cmdGoToScanning].Visible = False
    Me![cmdAddSieveScan].Visible = True
    Me![cmdGoToSieveScanning].Visible = False
    OptionRowSource = OptionRowSource & "Put record in Priority;"
    OptionRowSource = OptionRowSource & "Put record in Scanning;"
    OptionRowSource = OptionRowSource & "Put record in Sieve Scanning;"
    Me!cboOptions.RowSource = OptionRowSource
    
    Forms![FrmBasicData]![cmdHistoric].Enabled = False
    
    'saj july 2007 - if a float date was altered in a previous record then carry it across here
    If dateTouse <> "" Then
        Me![Float Date] = dateTouse
    End If
End If
Exit Sub
err_current:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Form_Open(Cancel As Integer)
'**********************************************************************************
' Set up display of the form based on whether the flot number has related records
' in the priority and sample tables. The command buttons appear on the form view and
' the combo cboOptions appears on the datasheet view to offer the same functionality as
' the buttons in this view.
' NB - the optionrowsource values are referred to in cboOptions_Afterupdate and form_open
' so any changes must be reflected there as well
' SAJ
'***********************************************************************************
On Error GoTo err_open
Dim OptionRowSource

'new on site - goto last record on open
DoCmd.GoToRecord acActiveDataObject, , acLast

If Me![chkInPriority] = False Then
    OptionRowSource = OptionRowSource & "Put record in Priority;"
    Me![cmdAddPriority].Visible = True
    Me![cmdGotoPriority].Visible = False
Else
'    Me![chkInPriority] = True
    OptionRowSource = OptionRowSource & "View Priority Record;"
    Me![cmdGotoPriority].Visible = True
    Me![cmdAddPriority].Visible = False
End If

If Me![chkInScanning] = False Then
'    Me![chkInScanning] = False
    OptionRowSource = OptionRowSource & "Put record in Scanning;"
    Me![cmdAddScan].Visible = True
    Me![cmdGoToScanning].Visible = False
Else
'    Me![chkInScanning] = True
    OptionRowSource = OptionRowSource & "View Scanning;"
    Me![cmdAddScan].Visible = False
    Me![cmdGoToScanning].Visible = True
End If

If Me![chkinsievescanning] = False Then
'    Me![chkInSieveScanning] = False
    OptionRowSource = OptionRowSource & "Put record in Sieve Scanning;"
    Me![cmdAddSieveScan].Visible = True
    Me![cmdGoToSieveScanning].Visible = False
Else
'    Me![chkInSieveScanning] = True
    OptionRowSource = OptionRowSource & "View Sieve Scanning;"
    Me![cmdAddSieveScan].Visible = False
    Me![cmdGoToSieveScanning].Visible = True
End If
Me!cboOptions.RowSource = OptionRowSource

'saj july 2007 dateTouse alows the last date input to be used as the default for the net record, set it to "" to begin with
dateTouse = ""
Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub

End Sub
