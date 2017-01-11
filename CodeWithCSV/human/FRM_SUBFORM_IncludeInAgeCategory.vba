Option Compare Database
Option Explicit






Private Sub Check21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'late august 2009
'this recordset actually lockes this field so it cant be edited-am so short on time having to work around
'this with horrible hack to catch mouse click and action like an event
On Error GoTo err_chkInclude

Dim sql
If Me!Check21 = True Then
    'make the field false
    sql = "UPDATE [HR_ageing and sexing] SET [IncludeinAgeSexGrouping] = false WHERE [Unit number] = " & Me!Unit & " AND [Individual Number] = " & Me![IndividualNumber] & ";"
    DoCmd.RunSQL sql
Else
    'make the field true
    sql = "UPDATE [HR_ageing and sexing] SET [IncludeinAgeSexGrouping] = true WHERE [Unit number] = " & Me!Unit & " AND [Individual Number] = " & Me![IndividualNumber] & ";"
    DoCmd.RunSQL sql
End If
Me.Requery
Exit Sub

err_chkInclude:
    Call General_Error_Trap
    
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name
    

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub Form_Open(Cancel As Integer)
'use open args to get the where criteria for this recordsource
'problem is using the query Q_IncludeinAgeCategory doesn't work as its a distinct query that
'needs to hide the related to fields. That means that passing in where criteria including relatedTo
'does not work as this form can't actually see them!
'late august 2009
On Error GoTo err_open

If Me.OpenArgs <> "" Then
    Me.RecordSource = "SELECT DISTINCT [HR_Skeleton_RelatedTo_Skeleton].[Unit], [HR_Skeleton_RelatedTo_Skeleton].[IndividualNumber], [HR_ageing and sexing].[IncludeinAgeSexGrouping] FROM HR_Skeleton_RelatedTo_Skeleton LEFT JOIN [HR_ageing and sexing] ON ([HR_Skeleton_RelatedTo_Skeleton].[Unit]=[HR_ageing and sexing].[unit number]) AND ([HR_Skeleton_RelatedTo_Skeleton].[IndividualNumber]=[HR_ageing and sexing].[Individual number]) WHERE " & Me.OpenArgs & ";"

End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
    
End Sub
