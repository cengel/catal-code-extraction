Option Compare Database   'Use database order for string comparisons

Private Sub Area_Sheet_Click()
On Error GoTo Err_Area_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Area Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Area_Sheet_Click:
    Exit Sub

Err_Area_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Area_Sheet_Click
End Sub


Private Sub Building_Sheet_Click()
On Error GoTo Err_Building_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Building Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Building_Sheet_Click:
    Exit Sub

Err_Building_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Building_Sheet_Click
End Sub

Private Sub Button10_Click()
Building_Sheet_Click
End Sub

Private Sub Button11_Click()
Space_Sheet_Button_Click
End Sub


Private Sub Button12_Click()
Feature_Sheet_Button_Click
End Sub


Private Sub Button13_Click()
Unit_Sheet_Click
End Sub


Private Sub Button17_Click()
Area_Sheet_Click
End Sub

Private Sub Button9_Click()
Return_to_Master_Con_Click
End Sub

Private Sub Command18_Click()
Open_priority_Click
End Sub

Private Sub Feature_Sheet_Button_Click()
On Error GoTo Err_Feature_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Feature Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Feature_Sheet_Button_Click:
    Exit Sub

Err_Feature_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Feature_Sheet_Button_Click
End Sub


Sub Open_priority_Click()
On Error GoTo Err_Open_priority_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Priority Detail"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Open_priority_Click:
    Exit Sub

Err_Open_priority_Click:
    MsgBox Err.Description
    Resume Exit_Open_priority_Click
    
End Sub

Private Sub Return_to_Master_Con_Click()
On Error GoTo Err_Return_to_Master_Con_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Excavation"
    
Exit_Return_to_Master_Con_Click:
    Exit Sub

Err_Return_to_Master_Con_Click:
    MsgBox Err.Description
    Resume Exit_Return_to_Master_Con_Click

End Sub

Private Sub Space_Sheet_Button_Click()
On Error GoTo Err_Space_Sheet_Button_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Space Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Space_Sheet_Button_Click:
    Exit Sub

Err_Space_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Space_Sheet_Button_Click

End Sub


Private Sub Unit_Sheet_Click()
On Error GoTo Err_Unit_Sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Exca: Unit Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Unit_Sheet_Click:
    Exit Sub

Err_Unit_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Unit_Sheet_Click

End Sub


Sub Command19_Click()
On Error GoTo Err_Command19_Click


    DoCmd.Close

Exit_Command19_Click:
    Exit Sub

Err_Command19_Click:
    MsgBox Err.Description
    Resume Exit_Command19_Click
    
End Sub
Sub Quit_Click()
On Error GoTo Err_Quit_Click


    DoCmd.Quit

Exit_Quit_Click:
    Exit Sub

Err_Quit_Click:
    MsgBox Err.Description
    Resume Exit_Quit_Click
    
End Sub
Sub open_register_Click()
On Error GoTo Err_open_register_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Store: Crate Register"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_open_register_Click:
    Exit Sub

Err_open_register_Click:
    MsgBox Err.Description
    Resume Exit_open_register_Click
    
End Sub
Sub preview_report_Click()
On Error GoTo Err_preview_report_Click

    Dim stDocName As String

    stDocName = "Finds Store: Crate Register"
    DoCmd.OpenReport stDocName, acPreview

Exit_preview_report_Click:
    Exit Sub

Err_preview_report_Click:
    MsgBox Err.Description
    Resume Exit_preview_report_Click
    
End Sub
Private Sub Command23_Click()
On Error GoTo Err_Command23_Click

    Dim stDocName As String

    stDocName = "Finds Store: Crate Register"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command23_Click:
    Exit Sub

Err_Command23_Click:
    MsgBox Err.Description
    Resume Exit_Command23_Click
    
End Sub
Private Sub SearchbyUnits_Click()
On Error GoTo Err_SearchbyUnits_Click

    Dim stDocName As String

    stDocName = "Unit Based Query Report"
    DoCmd.OpenReport stDocName, acPreview

Exit_SearchbyUnits_Click:
    Exit Sub

Err_SearchbyUnits_Click:
    MsgBox Err.Description
    Resume Exit_SearchbyUnits_Click
    
End Sub
