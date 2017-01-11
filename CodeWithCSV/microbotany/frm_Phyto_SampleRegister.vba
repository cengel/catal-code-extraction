Option Compare Database
Option Explicit


Private Sub addnewsample_Click()
'create a new record for a new sample
On Error GoTo err_addnewsample_Click

DoCmd.OpenForm "Phytolith_NewSampleRegister", acNormal
DoCmd.GoToRecord , "Phytolith_NewSampleRegister", acNewRec
Forms![Phytolith_NewSampleRegister].Form![UnitNumber].SetFocus

Exit Sub

err_addnewsample_Click:
    MsgBox "An error has occured setting up a new record, the error is as follows: " & Err.Description
    Exit Sub

End Sub

Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_cboFindUnit_AfterUpdate

    If Me![cboFindUnit] <> "" Then
         'for existing number the field with be disabled, enable it as when find num
        'is shown the on current event will deal with disabling it again
        If Me![UnitNumber].Enabled = False Then Me![UnitNumber].Enabled = True
        DoCmd.GoToControl "UnitNumber"
        DoCmd.FindRecord Me![cboFindUnit]
        '2009 move the blank down to after the gotocontrol as code in Year_lostfocus checks
        'for a value when the year looses it - complicated passing of focus nightmare stuff
        'Me![cboFindUnit] = ""
        '2009 focus will bounce on Year fld and can be easily over written to make sure stays here
        DoCmd.GoToControl "cboFindUnit"
        Me![cboFindUnit] = ""
    End If
Exit Sub

err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFindUnit_NotInList(NewData As String, Response As Integer)
'stop not in list msg loop - 2009
On Error GoTo err_cbofindNotInList

    MsgBox "Sorry this Unit cannot be found in the list", vbInformation, "No Match"
    Response = acDataErrContinue
    
    Me![cboFindUnit].Undo
    '2009 if not esc the list will stay pulled down making it hard to go direct to Add new or where ever as
    'have to escape the pull down list first
    SendKeys "{ESC}"
Exit Sub

err_cbofindNotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
Exit Sub

err_close:
    MsgBox "An error has occured: " & Err.Description
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

If (Forms![frm_Phyto_Data_Entry].[SampleProcessYear] <> "") And (Forms![frm_Phyto_Data_Entry].[SampleNumber] <> "") Then

    Dim mydb As Database, myrs As DAO.Recordset, sql, sql1
    Set mydb = CurrentDb
    sql = "SELECT * FROM [PhytolithLOV_PhytoNames] WHERE [PhytoName] <> '' ORDER BY [PhytoID], [PhytoMultiOrSingle]"
    Set myrs = mydb.OpenRecordset(sql)

    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        Do Until myrs.EOF
            sql1 = "INSERT INTO [Phytolith_Sample_Analysis_Details] ([SiteCode], [SampleProcessYear], [SampleNumber], [SampleID], [SingleOrMulti], [DicotOrMonocot], [PhytoName])"
            sql1 = sql1 & " VALUES ('" & Forms![frm_Phyto_Data_Entry].[SiteCode] & "', '" & Forms![frm_Phyto_Data_Entry].[SampleProcessYear] & "', " & Forms![frm_Phyto_Data_Entry].[SampleNumber] & ", '" & Forms![frm_Phyto_Data_Entry].[SampleID] & "', '" & myrs![PhytoMultiOrSingle] & "', '" & myrs![DicotOrMonocot] & "','" & myrs![PhytoName] & "');"
            DoCmd.RunSQL sql1
            myrs.MoveNext
        Loop
    End If
    
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



Private Sub Combo61_NotInList(NewData As String, Response As Integer)
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

Private Sub Form_Current()
'set up display depending on reason for collection
On Error GoTo err_curr

Exit Sub

err_curr:
    Call General_Error_Trap
    Exit Sub
End Sub




Private Sub Befehl99_Click()
On Error GoTo Err_Befehl99_Click


    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdFind

Exit_Befehl99_Click:
    Exit Sub

Err_Befehl99_Click:
    MsgBox Err.Description
    Resume Exit_Befehl99_Click
    
End Sub


Private Sub GoToFull_Click()

On Error GoTo Err_GoToFull_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retVal, sql, insertArea, permiss
    
    stDocName = "frm_Phyto_Data_Entry"
    
    If Not IsNull(Me![UnitNumber]) Or Me![UnitNumber] <> "" Or Not IsNull(Me![SampleNumber]) Or Me![SampleNumber] <> "" Then
        'check that feature num does exist
        checknum = DLookup("[FieldID]", "[Phytolith_Sample_Analysis_Details]", "[Unit] = " & Me![UnitNumber] & " AND [LabSampleNumber] = " & Me![SampleNumber])
        If IsNull(checknum) Then
            'number not exist - now see what permissions user has
            Debug.Print GetGeneralPermissions
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Sample DOES NOT EXIST in 'Sample Details'."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retVal = MsgBox(msg, vbInformation + vbYesNo, "Sample does not exist")
        
                If retVal = vbNo Then
                    MsgBox "No sample record to view, please alert the your team leader about this.", vbExclamation, "Missing Sample Record"
                Else
                    'add new records behind scences
                    sql = "INSERT INTO [Phytolith_Sample_Analysis_Details] ([Unit], [LabSampleNumber]) VALUES (" & Me![UnitNumber] & ", " & Me![SampleNumber] & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm stDocName, acNormal, , "[Unit] = " & Me![UnitNumber] & " AND [LabSampleNumber] = " & Me![SampleNumber], acFormEdit, acDialog
                End If
            Else
                'user is readonly so just tell them record not exist
                MsgBox "Sorry but this sample record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Sample Record"
            End If
        Else
            'record exists - open it
            stLinkCriteria = "[Unit] = " & Me![UnitNumber] & " AND [LabSampleNumber] = " & Me![SampleNumber]
            'DoCmd.OpenForm stDocName, , , stLinkCriteria
            'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
            'decided against dialog as can open other forms on the feature form and they would appear underneath it
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Sample to view", vbInformation, "No Sample Number"
    End If
    
Exit_GoToFull_Click:
    Exit Sub


Err_GoToFull_Click:
    Call General_Error_Trap
    Resume Exit_GoToFull_Click

End Sub

Private Sub GoToSample_Click()
On Error GoTo Err_GoToSample_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retVal, sql, insertArea, permiss
    Dim priorization
    
    stDocName = "frm_Phyto_FieldAnalysis"
    
    If Not IsNull(Me![UnitNumber]) Or Me![UnitNumber] <> "" Or Not IsNull(Me![SampleNumber]) Or Me![SampleNumber] <> "" Then
        'check that feature num does exist
        checknum = DLookup("[FieldID]", "[Phytolith_FieldAnalysis]", "[Unit] = " & Me![UnitNumber] & " AND [SampleNumber] = " & Me![SampleNumber])
        If IsNull(checknum) Then
            'number not exist - now see what permissions user has
            Debug.Print GetGeneralPermissions
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Sample DOES NOT EXIST in 'Sample Details'."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retVal = MsgBox(msg, vbInformation + vbYesNo, "Sample does not exist")
        
                If retVal = vbNo Then
                    MsgBox "No sample record to view, please alert the your team leader about this.", vbExclamation, "Missing Sample Record"
                Else
                    'add new records behind scences
                    If Me![priorityunit] = True Then
                        priorization = "Priority Tour"
                    Else
                        priorization = ""
                    End If
                    sql = "INSERT INTO [Phytolith_FieldAnalysis] ([Unit], [SampleLetter], [SampleNumber], [CheckReason]) VALUES (" & Me![UnitNumber] & ", 's', " & Me![SampleNumber] & ", '" & priorization & "');"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm stDocName, acNormal, , "[Unit] = " & Me![UnitNumber] & " AND [SampleNumber] = " & Me![SampleNumber]
                End If
            Else
                'user is readonly so just tell them record not exist
                MsgBox "Sorry but this sample record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Sample Record"
            End If
        Else
            'record exists - open it
            stLinkCriteria = "[Unit] = " & Me![UnitNumber] & " AND [SampleNumber] = " & Me![SampleNumber]
            'DoCmd.OpenForm stDocName, , , stLinkCriteria
            'DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
            'decided against dialog as can open other forms on the feature form and they would appear underneath it
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria
        End If
    Else
        MsgBox "No Sample to view", vbInformation, "No Sample Number"
    End If
    
Exit_GoToSample_Click:
    Exit Sub


Err_GoToSample_Click:
    Call General_Error_Trap
    Resume Exit_GoToSample_Click
End Sub
