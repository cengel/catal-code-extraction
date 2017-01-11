Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo err_cancel

    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_cancel:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReport_Click()
On Error GoTo err_cmdReport

    If IsNull(Me![optionFrame]) Then
        MsgBox "Select a dataset first"
    Else
        If Me![optionFrame] < 1 Or Me![optionFrame] > 4 Then
            MsgBox "Invalid parameter", vbInformation, "Invalid Operation"
        Else
            Dim numselect
            numselect = Me![optionFrame]
            Select Case numselect
            Case 1
                If FindMissingNumbers("Exca: Building Details", "Number") = True Then
                    DoCmd.OpenReport "R_MissingBuildings", acViewPreview
                End If
            Case 2
                If FindMissingNumbers("Exca: Space Sheet", "Space Number") = True Then
                    DoCmd.OpenReport "R_MissingSpaces", acViewPreview
                End If
            Case 3
                If FindMissingNumbers("Exca: Features", "Feature Number") = True Then
                    DoCmd.OpenReport "R_MissingFeatures", acViewPreview
                End If
            Case 4
                If FindMissingNumbers("Exca: Unit sheet", "unit number") = True Then
                    DoCmd.OpenReport "R_MissingUnits", acViewPreview
                End If
            End Select
            
        
        End If
    End If
Exit Sub

err_cmdReport:
    Call General_Error_Trap
    Exit Sub
End Sub
