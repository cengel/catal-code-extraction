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


Sub calc_Click()
On Error GoTo Err_calc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Ph2 Standardize pop-up"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_calc_Click:
    Exit Sub

Err_calc_Click:
    MsgBox Err.Description
    Resume Exit_calc_Click
    
End Sub
Sub calculate2_Click()
On Error GoTo Err_calculate2_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Light Ph2 Stand pop-up"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_calculate2_Click:
    Exit Sub

Err_calculate2_Click:
    MsgBox Err.Description
    Resume Exit_calculate2_Click
    
End Sub

Private Sub Command127_Click()

End Sub


Sub Calculate_Sums_Click()
On Error GoTo Err_Calculate_Sums_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Bots98: Heavy Ph3 Standardize Pop-up"
    
    stLinkCriteria = "[GID]=" & "'" & Me![GID] & "' AND [Material]=" & "'" & Me![Material] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Calculate_Sums_Click:
    Exit Sub

Err_Calculate_Sums_Click:
    MsgBox Err.Description
    Resume Exit_Calculate_Sums_Click
    
End Sub
Sub back_Click()
On Error GoTo Err_back_Click

    Dim stDocName As String
    Dim stLinkCriteria As String, findmatch As String

    stDocName = "Bots98: Light Residue Phase 2"
    
    'stLinkCriteria = "[GID]=" & "'" & Me![GID] & "'"
    findmatch = Me![GID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.FindRecord findmatch, , True, , True
    
Exit_back_Click:
    Exit Sub

Err_back_Click:
    MsgBox Err.Description
    Resume Exit_back_Click
    
End Sub
