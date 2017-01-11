Option Compare Database
Option Explicit

Private Sub Command47_Click()

Dim s4wt As Single, s2wt As Single, s1wt As Single, s05wt As Single
Dim s4ct As Single, s2ct As Single, s1ct As Single, s05ct As Single

Dim mydb As Database, Recs As DAO.Recordset
Dim strSQL As String

    ' Return Database variable pointing to current database.
    Set mydb = CurrentDb
    strSQL = "SELECT DISTINCTROW [Bots98: Light Ph2 Material].GID, [Bots98: Light Ph2 Material].Material, [Bots98: Basic Flot details].[Vol in Litres], [Bots98: Light Ph2 Material].TotalWeight, [Bots98: Light Ph2 Material].TotalCount, [Bots98: Light Ph2 Material].[4 Weight], [Bots98: Light Ph2 Material].[4 Count], [Bots98: Light Ph2 Material].[4 % sorted], [Bots98: Light Ph2 Material].[2 Weight], [Bots98: Light Ph2 Material].[2 Count], [Bots98: Light Ph2 Material].[2 % sorted], [Bots98: Light Ph2 Material].[1 Weight], [Bots98: Light Ph2 Material].[1 Count], [Bots98: Light Ph2 Material].[1 % sorted], [Bots98: Light Ph2 Material].[0,5 Weight], [Bots98: Light Ph2 Material].[0,5 Count], [Bots98: Light Ph2 Material].[0,5 % sorted] FROM [Bots98: Light Ph2 Material] INNER JOIN [Bots98: Basic Flot details] ON [Bots98: Light Ph2 Material].GID = [Bots98: Basic Flot details].GID;"
    Set Recs = mydb.OpenRecordset(strSQL)
    
Do Until Recs.EOF
With Recs
If ![Vol in Litres] > 0 Then
    'do WEIGHTS and COUNTS for each fraction
    
    If ![4 % sorted] = 0 Then
        s4wt = 0
        s4ct = 0
    Else
        s4wt = ![4 Weight]
        s4ct = ![4 Count]
    End If
    Debug.Print "4%" & s4wt
    
    If ![2 % sorted] = 0 Then
        s2wt = 0
        s2ct = 0
    Else
        s2wt = ![2 Weight]
        s2ct = ![2 Count]
    End If
    
    If ![1 % sorted] = 0 Then
        s1wt = 0
        s1ct = 0
    Else
        s1wt = ![1 Weight]
        s1ct = ![1 Count]
    End If
    
    If ![0,5 % sorted] = 0 Then
        s05wt = 0
        s05ct = 0
    Else
        s05wt = ![0,5 Weight]
        s05ct = ![0,5 Count]
    End If
   
End If 'litres

.Edit
![TotalWeight] = s4wt + s2wt + s1wt + s05wt
![TotalCount] = s4ct + s2ct + s1ct + s05ct

'Forms![Bots98: Light Ph2 Stand pop-up]![TotalWeight] = totwt
'Forms![Bots98: Light Ph2 Stand pop-up]![TotalCount] = totct
.Update
.MoveNext
End With

Loop
Recs.MoveLast
Debug.Print Recs.RecordCount
Recs.Close

End Sub


Private Sub Form_Current()
Dim s4wt As Single, s2wt As Single, s1wt As Single, s05wt As Single
Dim s4ct As Single, s2ct As Single, s1ct As Single, s05ct As Single
Dim totwt As Single, totct As Single
Dim substring As Object

Set substring = Me![Bots98: subform Standardize Heavy2]
'Set substring = Forms![Bots98: Light Ph2 Stand pop-up2]

If substring![Vol in Litres] > 0 Then
    'do WEIGHTS and COUNTS for each fraction
    
    If substring![4 % sorted] = 0 Then
        s4wt = 0
        s4ct = 0
    Else
        s4wt = substring![stand4wt]
        s4ct = substring![Stand4ct]
    End If
    
    If substring![2 % sorted] = 0 Then
        s2wt = 0
        s2ct = 0
    Else
        s2wt = substring![stand2wt]
        s2ct = substring![Stand2ct]
    End If
    
    If substring![1 % sorted] = 0 Then
        s1wt = 0
        s1ct = 0
    Else
        s1wt = substring![Stand1wt]
        s1ct = substring![Stand1ct]
    End If
    
    If substring![0,5 % sorted] = 0 Then
        s05wt = 0
        s05ct = 0
    Else
        s05wt = substring![Stand05wt]
        s05ct = substring![Stand05ct]
    End If
   
End If 'litres

totwt = s4wt + s2wt + s1wt + s05wt
totct = s4ct + s2ct + s1ct + s05ct

Me![TotalWeight] = totwt
Me![TotalCount] = totct


End Sub

Sub Close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub

Sub run_Click()
On Error GoTo Err_run_Click


Dim s4wt As Single, s2wt As Single, s1wt As Single, s05wt As Single
Dim s4ct As Single, s2ct As Single, s1ct As Single, s05ct As Single
Dim totwt As Single, totct As Single
Dim substring As Object
Dim Recs As Recordset

Set Recs = Me.RecordsetClone

Set substring = Forms![Bots98: Light Ph2 Stand pop-up]![Bots98: Standardize subform1]
'Set substring = Forms![Bots98: Light Ph2 Stand pop-up2]
Do Until Recs.EOF

If substring![Vol in Litres] > 0 Then
    'do WEIGHTS and COUNTS for each fraction
    
    If substring![4 % sorted] = 0 Then
        s4wt = 0
        s4ct = 0
    Else
        s4wt = substring![stand4wt]
        s4ct = substring![Stand4ct]
    End If
    Debug.Print "4%" & s4wt
    
    If substring![2 % sorted] = 0 Then
        s2wt = 0
        s2ct = 0
    Else
        s2wt = substring![stand2wt]
        s2ct = substring![Stand2ct]
    End If
    
    If substring![1 % sorted] = 0 Then
        s1wt = 0
        s1ct = 0
    Else
        s1wt = substring![Stand1wt]
        s1ct = substring![Stand1ct]
    End If
    
    If substring![0,5 % sorted] = 0 Then
        s05wt = 0
        s05ct = 0
    Else
        s05wt = substring![Stand05wt]
        s05ct = substring![Stand05ct]
    End If
   
End If 'litres

totwt = s4wt + s2wt + s1wt + s05wt
totct = s4ct + s2ct + s1ct + s05ct

Forms![Bots98: Light Ph2 Stand pop-up]![TotalWeight] = totwt
Forms![Bots98: Light Ph2 Stand pop-up]![TotalCount] = totct
Recs.Update
Recs.MoveNext

Loop
Recs.Close
'If Not Recs.EOF Then
'Recs.MoveLast
'End If

Exit_run_Click:
    Exit Sub

Err_run_Click:
    MsgBox Err.Description
    Resume Exit_run_Click
    
End Sub

Private Sub Form_Load()
Dim s4wt As Single, s2wt As Single, s1wt As Single, s05wt As Single
Dim s4ct As Single, s2ct As Single, s1ct As Single, s05ct As Single
Dim totwt As Single, totct As Single
Dim substring As Object

Set substring = Me![Bots98: subform Standardize Heavy2]

If substring![Vol in Litres] > 0 Then
    'do WEIGHTS and COUNTS for each fraction
    
    If substring![4 % sorted] = 0 Then
        s4wt = 0
        s4ct = 0
    Else
        s4wt = substring![stand4wt]
        s4ct = substring![Stand4ct]
    End If
    
    If substring![2 % sorted] = 0 Then
        s2wt = 0
        s2ct = 0
    Else
        s2wt = substring![stand2wt]
        s2ct = substring![Stand2ct]
    End If
    
    If substring![1 % sorted] = 0 Then
        s1wt = 0
        s1ct = 0
    Else
        s1wt = substring![Stand1wt]
        s1ct = substring![Stand1ct]
    End If
    
    If substring![0,5 % sorted] = 0 Then
        s05wt = 0
        s05ct = 0
    Else
        s05wt = substring![Stand05wt]
        s05ct = substring![Stand05ct]
    End If
   
    totwt = s4wt + s2wt + s1wt + s05wt
    totct = s4ct + s2ct + s1ct + s05ct
Else 'no litres
    MsgBox "Sample Volume is 0 Litres, no calculation possible. All standardized values for this sample will be set to 0."
    totwt = 0
    totct = 0
    
End If 'litres

Me![TotalWeight] = totwt
Me![TotalCount] = totct

End Sub


