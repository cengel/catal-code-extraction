Option Compare Database
Option Explicit

Private Sub AnalysisLevel_AfterUpdate()

If Me![AnalysisLevel] = "Second" Then
    Me![button.goto.debitage].Enabled = True
Else: Me![button.goto.debitage].Enabled = False

End If

End Sub


Private Sub Bag_AfterUpdate()

Me![cmd.gotonew].Enabled = True

End Sub


Private Sub Form_Current()

If Me![AnalysisLevel] = "first" Or IsNull(Me![AnalysisLevel]) Then
    Me![button.goto.debitage].Enabled = False
Else
    Me![button.goto.debitage].Enabled = True
End If

If Me![Bag] = 0 Then
    Me![cmd.gotonew].Enabled = False
Else
    Me![cmd.gotonew].Enabled = True
End If

End Sub


Sub OpenForm_Blades_Click()

On Error GoTo Err_OpenForm_Blades_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stPrimaryAnalysis As String

    stDocName = "LithicForm:Blades"
    stPrimaryAnalysis = "first"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]

If Me![Regular Blades] = 0 And Me![Non-Regular Blades] = 0 Or Me.AnalysisLevel = stPrimaryAnalysis Then
    MsgBox "No data.", 0, "Error"
        
Else
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
            
End If

Exit_OpenForm_Blades_Click:
    Exit Sub

Err_OpenForm_Blades_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_Blades_Click
    
End Sub


Sub OpenForm_Debitage_Click()
On Error GoTo Err_OpenForm_Debitage_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stPrimaryAnalysis As String
    
    stPrimaryAnalysis = "first"
    stDocName = "LithicForm:Debitage"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]
    
If Me.AnalysisLevel = stPrimaryAnalysis Then
    MsgBox "Secondary analysis data not available.", 0, "Error"
Else
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
            
End If

Exit_OpenForm_Debitage_Click:
    Exit Sub

Err_OpenForm_Debitage_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_Debitage_Click
    
End Sub
Sub OpenForm_Cores_Click()
On Error GoTo Err_OpenForm_Cores_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stPrimaryAnalysis As String

    stDocName = "LithicForm:Cores"
    stPrimaryAnalysis = "first"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]
    

If Me![Blade Cores] = 0 And Me![Flake Cores] = 0 Or Me.AnalysisLevel = stPrimaryAnalysis Then
    MsgBox "No data.", 0, "Error"
    
Else
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
End If

Exit_OpenForm_Cores_Click:
    Exit Sub

Err_OpenForm_Cores_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_Cores_Click
    
End Sub
Sub OpenForm_ModifiedBlanks_Click()
On Error GoTo Err_OpenForm_ModifiedBlanks_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stPrimaryAnalysis As String

    stDocName = "LithicForm:ModifiedBlanks"
    stPrimaryAnalysis = "first"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]

If Me![Retouched Flakes] = 0 And Me![Retouched Blades] = 0 And IsNull(Me![Other Retouched]) Or Me.AnalysisLevel = stPrimaryAnalysis Then
    MsgBox "No data.", 0, "Error"
        
Else
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
            
End If
Exit_OpenForm_ModifiedBlanks_Click:
    Exit Sub

Err_OpenForm_ModifiedBlanks_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_ModifiedBlanks_Click
    
End Sub
Sub OpenForm_Biface_Click()
On Error GoTo Err_OpenForm_Biface_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stPrimaryAnalysis As String

    stDocName = "LithicForm:Bifaces"
    stPrimaryAnalysis = "first"
    
    stLinkCriteria = "[Bag]=" & Me![Bag]
    
If Me![Fragmentary P/Bs] = 0 And Me![Complete P/Bs] = 0 Or Me.AnalysisLevel = stPrimaryAnalysis Then
    MsgBox "No data.", 0, "Error"
        
Else
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria
            
End If

Exit_OpenForm_Biface_Click:
    Exit Sub

Err_OpenForm_Biface_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_Biface_Click
    
End Sub
Sub OpenForm_UnitTotals_Click()
On Error GoTo Err_OpenForm_UnitTotals_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:UnitCounts"
    
    stLinkCriteria = "[Unit]=" & Me![Unit]
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OpenForm_UnitTotals_Click:
    Exit Sub

Err_OpenForm_UnitTotals_Click:
    MsgBox Err.Description
    Resume Exit_OpenForm_UnitTotals_Click
    
End Sub
Sub button_goto_debitage_Click()
On Error GoTo Err_button_goto_debitage_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LithicForm:Debitage"
    stLinkCriteria = "[Bag]=" & Me![Bag]
    
    DoCmd.Minimize
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_goto_debitage_Click:
    Exit Sub

Err_button_goto_debitage_Click:
    MsgBox Err.Description
    Resume Exit_button_goto_debitage_Click
    
End Sub
Sub cmd_gotonew_Click()
On Error GoTo Err_cmd_gotonew_Click


    DoCmd.GoToRecord , , acNewRec
    Bag.SetFocus

Exit_cmd_gotonew_Click:
    Exit Sub

Err_cmd_gotonew_Click:
    MsgBox Err.Description
    Resume Exit_cmd_gotonew_Click
    
End Sub
