Option Compare Database
Option Explicit

Private Sub cboFindSample_AfterUpdate()
'find choosen sample id
On Error GoTo err_cboFind
    If Me![cboFindSample] <> "" Then
        DoCmd.GoToControl "txtSampleID"
        DoCmd.FindRecord Me![cboFindSample]
    End If
Exit Sub

err_cboFind:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    MsgBox "An error has occured: " & Err.Description
End Sub

Private Sub cmdDelete_Click()
'delete here so can clean up sub tables
On Error GoTo err_cmdDelete

    'double check is admin
    Dim permiss
    permiss = GetGeneralPermissions
    
    If permiss <> "ADMIN" Then
        MsgBox "You do not have permission to delete records. Contact your supervisor.", vbInformation, "Permission Denied"
    Else
        Dim retVal, sql
        retVal = MsgBox("Really delete Sample ID: " & Me![txtSampleID] & "?", vbCritical + vbYesNoCancel, "Confirm Delete")
        If retVal = vbYes Then
            sql = "Delete from [Phytolith_Sample_Analysis_Details] WHERE [SampleID] = '" & Me![txtSampleID] & "';"
            DoCmd.RunSQL sql
            
            sql = "Delete from [Phytolith_Sample_Core_Data] WHERE [SampleID] = '" & Me![txtSampleID] & "';"
            DoCmd.RunSQL sql
            
            Me.Requery
            DoCmd.GoToRecord acActiveDataObject, , acLast
        End If
    End If
    
Exit Sub

err_cmdDelete:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdExport_Click()
'This exports the core data about the sample shown on screen to excel

On Error GoTo err_cmdExport
'DoCmd.RunCommand acCmdOutputToExcel
DoCmd.OutputTo acOutputQuery, "Q_ExportToExcel_PhytoData_OnScreen", acFormatXLS, "PhytoData for sample " & Me![SampleID] & ".xls", True

Exit Sub

err_cmdExport:
   ' Call general_error_trap
    Exit Sub
End Sub

Private Sub cmdExportAll_Click()
'export all core information entered into database into excel
On Error GoTo err_cmdExportAll

    DoCmd.OutputTo acOutputTable, "Phytolith_Sample_Core_Data", acFormatXLS, "All Phyto Core Data from Catal Database.xls", True
Exit Sub

err_cmdExportAll:
    MsgBox "An error has occured, the description is: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdExportAllSamples_Click()
'export all sample info into excel
On Error GoTo err_cmdExportAllSamples

    DoCmd.OutputTo acOutputTable, "Phytolith_Sample_Analysis_Details", acFormatXLS, "All Phyto Sample data from Catal Database.xls", True
Exit Sub

err_cmdExportAllSamples:
    MsgBox "An error has occured, the description is: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdExportThisSample_Click()
'export just the sample shown on screen to excel
On Error GoTo err_cmdExport

    DoCmd.OutputTo acOutputQuery, "Q_ExportToExcel_Sample_OnScreen", acFormatXLS, "Phyto Sample Data for " & Me![SampleID] & ".xls", True
Exit Sub

err_cmdExport:
    MsgBox "An error has occured, the description is: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdReady_Click()
'set up the phyto names for data entry
On Error GoTo err_cmdReady

If (Forms![frm_Phyto_Data_Entry].[SampleProcessYear] <> "") And (Forms![frm_Phyto_Data_Entry].[LabSampleNumber] <> "") Then

    Dim mydb As Database, myrs As DAO.Recordset, sql, sql1
    Set mydb = CurrentDb
    sql = "SELECT * FROM [PhytolithLOV_PhytoNames] WHERE [PhytoName] <> '' ORDER BY [PhytoID], [PhytoMultiOrSingle]"
    Set myrs = mydb.OpenRecordset(sql)
    Me![lblPressReady].Visible = False
    Me![lblInform].Visible = True
    Me![lblInform].Caption = "Phytolith List is being prepared"
    Me.Repaint
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        Do Until myrs.EOF
            sql1 = "INSERT INTO [Phytolith_Sample_Analysis_Details] ([SiteCode], [SampleProcessYear], [LabSampleNumber], [SampleID], [SingleOrMulti], [DicotOrMonocot], [PhytoName])"
            sql1 = sql1 & " VALUES ('" & Forms![frm_Phyto_Data_Entry].[SiteCode] & "', '" & Forms![frm_Phyto_Data_Entry].[SampleProcessYear] & "', '" & Forms![frm_Phyto_Data_Entry].[LabSampleNumber] & "', '" & Forms![frm_Phyto_Data_Entry].[SampleID] & "', '" & myrs![PhytoMultiOrSingle] & "', '" & myrs![DicotOrMonocot] & "','" & myrs![PhytoName] & "');"
            DoCmd.RunSQL sql1
            Me![lblInform].Caption = "Phytolith List is being prepared: " & myrs![PhytoName] & " added"
            Me.Repaint
            myrs.MoveNext
        Loop
    End If
    
    Me![lblInform].Visible = False
    Me![frm_sub_phyto_data].Requery
    Me![frm_sub_phyto_data].Visible = True
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
Else
    MsgBox "Please enter the Sample Process Year and Sample Number first"
End If
Exit Sub

err_cmdReady:
    MsgBox "An error has occured setting up the data ready for entry. The error description is as follows: " & Err.Description, vbCritical, "Error"
    Exit Sub

End Sub

Private Sub cmdRecalc_Click()
'DoCmd.GoToControl Me![frm_sub_phyto_data].Name
'DoCmd.GoToControl Me![frm_sub_phyto_data].Form![PhytoCount].Name


End Sub

Private Sub cmdReport_Click()
'report the data - this sample or all
On Error GoTo err_report

    Dim retVal
    retVal = MsgBox("Do you want to report this sample only?" & Chr(13) & Chr(13) & "Press Yes to report " & Me![txtSampleID] & " only, press No to report on ALL samples", vbYesNoCancel + vbQuestion, "Report current sample only?")
    If retVal = vbNo Then
        DoCmd.OpenReport "Phytolith_Sample_Analysis_Details", acViewPreview
    ElseIf retVal = vbYes Then
        DoCmd.OpenReport "Phytolith_Sample_Analysis_Details", acViewPreview, , "[SampleID] ='" & Me![txtSampleID] & "'"
    End If
Exit Sub

err_report:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub

End Sub

Private Sub Command19_Click()
'create a new record for a new sample
On Error GoTo err_cmd19

DoCmd.RunCommand acCmdRecordsGoToNew

Exit Sub

err_cmd19:
    MsgBox "An error has occured setting up a new record, the error is as follows: " & Err.Description
    Exit Sub
End Sub

Private Sub Form_Current()
'set up the form dispay depending on whether sample details have been entered
On Error GoTo err_Current

If IsNull(Me![frm_sub_phyto_data].Form![SampleID]) Then
    'no sample details added yet
    Me![cmdReady].Visible = True
    'Me![cmdRecalc].Visible = False
    Me![frm_sub_phyto_data].Visible = False
    Me![cmdExportThisSample].Visible = False
    Me![boxReady].Visible = True
    Me![lblPressReady].Visible = True
Else
    Me![cmdReady].Visible = False
    'Me![cmdRecalc].Visible = True
    Me![frm_sub_phyto_data].Visible = True
     Me![cmdExportThisSample].Visible = True
     Me![boxReady].Visible = False
     Me![lblPressReady].Visible = False
End If
Exit Sub

err_Current:
    MsgBox "An error has occured, the description is: " & Err.Description
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)

On Error GoTo err_open
    
    Dim permiss
    permiss = GetGeneralPermissions

    If permiss = "ADMIN" Then
        Me![cmdDelete].Visible = True
    Else
        Me![cmdDelete].Visible = False
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub LabSampleNumber_AfterUpdate()
'update the sample id
On Error GoTo err_samplenum

If Me![SampleProcessYear] <> "" And Me![LabSampleNumber] <> "" Then
    Me![txtSampleID] = Me![SiteCode] & "-" & Me![SampleProcessYear] & "-" & Me![LabSampleNumber]
Else
    If Me![SampleProcessYear] <> "" Then
        MsgBox "Altering this value effects the sample ID"
        Me![SampleID] = Null
    End If
End If
Exit Sub

err_samplenum:
    MsgBox "Error: " & Err.Description
    Exit Sub
End Sub

Private Sub SampleProcessYear_AfterUpdate()
'update the sample id
On Error GoTo err_sampleyr

If Me![SampleProcessYear] <> "" And Me![LabSampleNumber] <> "" Then
    Me![txtSampleID] = Me![SiteCode] & "-" & Me![SampleProcessYear] & "-" & Me![LabSampleNumber]
Else
    If Me![LabSampleNumber] <> "" Then
        MsgBox "Altering this value effects the sample ID"
        Me![SampleID] = ""
    End If
End If
Exit Sub

err_sampleyr:
    MsgBox "Error: " & Err.Description
    Exit Sub
End Sub

Private Sub SampleProcessYear_NotInList(NewData As String, Response As Integer)
'allow entry of new year
On Error GoTo err_Year_NotInList

Dim retVal, sql

retVal = MsgBox("This year is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [PhytolithLOV_AnalysisYear]([AnalysisYear]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
    ''Response = acDataErrContinue
    'DoCmd.RunCommand acCmdSaveRecord
    'Me![SampleProcessYear].Requery
Else
    Response = acDataErrContinue
End If

Exit Sub

err_Year_NotInList:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub TotalMGPhyto_AfterUpdate()
'warn the users that changing this value effects the weight % calc
On Error GoTo err_TotalMGPhyto

Dim retVal

If Me![TotalMGSediment] <> "" Then
    If Me![Weight%] <> "" And Me![TotalMGPhyto].OldValue <> "" Then
        retVal = MsgBox("Changing this value affects the Weight % calculation (Total MG Phyto / Total MG Sediment * 100). Are you sure you want to update this value?", vbQuestion + vbYesNo, "Check Entry")
        If retVal = vbNo Then
            Me![TotalMGPhyto] = Me![TotalMGPhyto].OldValue
            Exit Sub
        End If
    End If
    Me![Weight%] = Me![TotalMGPhyto] / Me![TotalMGSediment] * 100
End If
Exit Sub

err_TotalMGPhyto:
    MsgBox "An error has occured trying to update the Weight % field, the description is as follows: " & Err.Description, vbCritical, "Error"
    Exit Sub

End Sub

Private Sub TotalMGSediment_AfterUpdate()
'warn the users that changing this value effects the weight % calc
On Error GoTo err_TotalMGSediment

Dim retVal

If Me![TotalMGPhyto] <> "" Then
    If Me![Weight%] <> "" And Me![TotalMGSediment].OldValue <> "" Then
        retVal = MsgBox("Changing this value affects the Weight % calculation (Total MG Phyto / Total MG Sediment * 100). Are you sure you want to update this value?", vbQuestion + vbYesNo, "Check Entry")
        If retVal = vbNo Then
            Me![TotalMGSediment] = Me![TotalMGSediment].OldValue
            Exit Sub
        End If
    End If
    Me![Weight%] = Me![TotalMGPhyto] / Me![TotalMGSediment] * 100
End If
Exit Sub

err_TotalMGSediment:
    MsgBox "An error has occured trying to update the Weight % field, the description is as follows: " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub
