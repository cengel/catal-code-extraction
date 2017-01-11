Option Compare Database
Option Explicit

Private Sub cmdExport_Click()
'On Error GoTo err_cmdExport
'
'    DoCmd.OutputTo acOutputQuery, "Q_ExportToExcel_Sample_OnScreen", acFormatXLS, "Phyto Sample Data for " & Me![SampleID] & ".xls", True
'Exit Sub
'
'err_cmdExport:
'    MsgBox "An error has occured, the description is: " & Err.Description
'    Exit Sub
End Sub

Private Sub PhytoCount_AfterUpdate()
'This is where the calculations are done for the n/slide and n/mg fields

On Error GoTo err_PhytoCount

Dim countedFields, totfields, result1, result2, result


If Me![PhytoCount] <> "" Then
    'only do this if a valid numeric count number has been entered
    If IsNumeric(Me![PhytoCount]) Then
        'n/slide = Count / fields counted on slide * total fields on slide
        'n/gm = n per slide / mg mounted * total mg phyto / Total mg sediment  * 1000
        If Me![SorM] = "Single" Then
            countedFields = Forms![frm_Phyto_Data_Entry]![FieldsCountedSinglePhyto]
        ElseIf Me![SorM] = "Multi" Then
            countedFields = Forms![frm_Phyto_Data_Entry]![FieldsCountedMultiPhyto]
        ElseIf Me![SorM] = "SilicaAgg" Then
            countedFields = Forms![frm_Phyto_Data_Entry]![FieldsCountedSilica]
        ElseIf Me![SorM] = "Other" Then 'ie non phyto
            countedFields = Forms![frm_Phyto_Data_Entry]![FieldsCountedNonPhyto]
        Else
            MsgBox "The Single / Multi / Silica Agg / Other field contains the value " & Me![SorM] & ", no calculation formula is stored for  this type", vbExclamation, "No calculation can be done"
        End If
        
        totfields = Forms![frm_Phyto_Data_Entry]![SlideFields]
        
        If IsNumeric(countedFields) And IsNumeric(totfields) Then
        
            result = Me![PhytoCount] / countedFields * totfields
            
            Me![PhytoN/Slide] = result
            
            If IsNumeric(Forms![frm_Phyto_Data_Entry]![MGMounted]) And IsNumeric(Forms![frm_Phyto_Data_Entry]![TotalMGPhyto]) And IsNumeric(Forms![frm_Phyto_Data_Entry]![TotalMGSediment]) Then
                'MGMounted, TotalMGPhyto,TotalMGSediment
                result2 = result / Forms![frm_Phyto_Data_Entry]![MGMounted] * Forms![frm_Phyto_Data_Entry]![TotalMGPhyto] / Forms![frm_Phyto_Data_Entry]![TotalMGSediment] * 1000
                Me![PhytoN/gm] = result2
            Else
                MsgBox "Cannot undertake n/gm calculation, one or more parameters is not numeric"
            End If
        Else
            MsgBox "Cannot undertake n/slide calculation, one or more parameters is not numeric"
        End If
        
        
    End If
End If
Exit Sub

err_PhytoCount:
    If Err.Number = 11 Then
        MsgBox "A problem has occured undertaking a calculation, the message is: " & Err.Description, vbCritical, "Error"
    Else
        MsgBox "A error has occured, the message is: " & Err.Description, vbCritical, "Error"
    End If
End Sub
