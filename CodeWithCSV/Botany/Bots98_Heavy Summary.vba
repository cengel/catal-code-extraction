Option Compare Database
Sub adjust_and_standardize(original As Object, adjust As Object, standard As Object, PERCENT As Object, litres As Object)

If Not IsNull(PERCENT) Then adjust = original * (100 / PERCENT)
If Not IsNull(litres) Then standard = adjust / litres

End Sub
Sub back_to_main_Click()
On Error GoTo Err_back_to_main_Click


Exit_back_to_main_Click:
    Exit Sub

Err_back_to_main_Click:
    MsgBox Err.Description
    Resume Exit_back_to_main_Click
    
End Sub
Sub adj_and_stand_4_mm_Click()
On Error GoTo Err_adj_and_stand_4_mm_Click

Dim litres1 As Object, adjusted As Object, standard As Object
Dim detail_form As Object, standard_form As Object

Set litres1 = Me![Vol in Litres]
Set detail_form = Forms![Bots: Heavy Residue Phase II]![Bots: Heavy Residue II subform]
Set standard_form = Forms![Bots: Heavy Residue Phase II]![Bots: Heavy Residue II standardized subform]

adjust_and_standardize detail_form![4 wood wt], standard_form![4 wood adj wt], standard_form![4 wood stand wt], detail_form![4 wood perc sort], litres1
adjust_and_standardize detail_form![4 cereal wt], standard_form![4 cereal adj wt], standard_form![4 cereal stand wt], detail_form![4 cereal perc sort], litres1
adjust_and_standardize detail_form![4 chaff wt], standard_form![4 CHAFF adj wt], standard_form![4 chaff stand wt], detail_form![4 CHAFF perc sort], litres1

Exit_adj_and_stand_4_mm_Click:
    Exit Sub

Err_adj_and_stand_4_mm_Click:
    MsgBox Err.Description
    Resume Exit_adj_and_stand_4_mm_Click
    
End Sub

Sub standard_Click()
On Error GoTo Err_standard_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Heavy Ph2 Standardize pop-up"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_standard_Click:
    Exit Sub

Err_standard_Click:
    MsgBox Err.Description
    Resume Exit_standard_Click
    
End Sub
