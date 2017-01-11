Option Compare Database
Option Explicit

Sub Close_Feature_Sheet_Click()
On Error GoTo Err_Close_Feature_Sheet_Click


    DoCmd.Close

Exit_Close_Feature_Sheet_Click:
    Exit Sub

Err_Close_Feature_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Close_Feature_Sheet_Click
    
End Sub

Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Feature Sheet"
    
Exit_Excavation_Click:
    Exit Sub

err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub




Private Sub Feature_Number_Exit(Cancel As Integer)
On Error GoTo Err_Feature_Number_Exit

    Me.refresh
    'DoCmd.Save acTable, "Exca: Unit Sheet"
    
Exit_Feature_Number_Exit:
    Exit Sub

Err_Feature_Number_Exit:
   
    'MsgBox Err.Description
    
    'MsgBox "This unit already exists in the database. Use the 'Find' button to go to it.", vbOKOnly, "duplicate"
    DoCmd.DoMenuItem acFormBar, acEditMenu, 0, , acMenuVer70
    
    Cancel = True
        
    Resume Exit_Feature_Number_Exit
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Feature Sheet]![dbo_Exca: FeatureHistory].Form![lastmodify].Value = Now()
End Sub

Private Sub go_next_Click()
On Error GoTo Err_go_next_Click


    DoCmd.GoToRecord , , acNext

Exit_go_next_Click:
    Exit Sub

Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub


Private Sub go_previous_Click()
On Error GoTo Err_go_previous_Click


    DoCmd.GoToRecord , , acPrevious

Exit_go_previous_Click:
    Exit Sub

Err_go_previous_Click:
    MsgBox Err.Description
    Resume Exit_go_previous_Click
End Sub

Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click


    DoCmd.GoToRecord , , acFirst

Exit_go_to_first_Click:
    Exit Sub

Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
End Sub

Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click


    DoCmd.GoToRecord , , acLast

Exit_go_last_Click:
    Exit Sub

Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
    
End Sub

Private Sub Master_Control_Click()
On Error GoTo Err_Master_Control_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Feature Sheet"
    
Exit_Master_Control_Click:
    Exit Sub

Err_Master_Control_Click:
    MsgBox Err.Description
    Resume Exit_Master_Control_Click
End Sub


Private Sub New_entry_Click()
On Error GoTo Err_New_entry_Click


    DoCmd.GoToRecord , , acNewRec
    Mound.SetFocus
    
Exit_New_entry_Click:
    Exit Sub

Err_New_entry_Click:
    MsgBox Err.Description
    Resume Exit_New_entry_Click
End Sub


Sub find_feature_Click()
On Error GoTo Err_find_feature_Click


    Screen.PreviousControl.SetFocus
    Feature_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_find_feature_Click:
    Exit Sub

Err_find_feature_Click:
    MsgBox Err.Description
    Resume Exit_find_feature_Click
    
End Sub
