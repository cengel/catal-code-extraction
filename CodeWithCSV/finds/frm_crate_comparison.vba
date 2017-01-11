Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
'close this pop up
On Error GoTo err_close
    
    DoCmd.Close acForm, Me.Name
    
Exit Sub

err_close:
    If Err.Number = 2450 Then
        'cant find form ie: not called from the find form, its not open
        Resume Next
    Else
        Call General_Error_Trap
    End If
    Exit Sub
    
End Sub

Private Sub cmdReturn_Click()
On Error GoTo err_cmdReturn

Dim mydb As DAO.Database, firstCrate As DAO.Recordset, SecondCrate As DAO.Recordset
Dim sql, sql1, criteria
Set mydb = CurrentDb

sql1 = "DELETE FROM [Temp_Store: Units in Crates]"
DoCmd.RunSQL sql1

'sql1 = "SELECT * FROM [Store: Units in Crates] WHERE CrateLetter = 'PG' AND CrateNumber = 1"
sql1 = "SELECT * FROM [Store: Units in Crates] WHERE CrateLetter = '" & Me!cboFirstCrate.Column(1) & "' AND CrateNumber = " & Me!cboFirstCrate.Column(2)
Set firstCrate = mydb.OpenRecordset(sql1, dbOpenSnapshot)

If Not (firstCrate.EOF And firstCrate.BOF) Then
    firstCrate.MoveFirst
    
    Do Until firstCrate.EOF
    
        sql = "SELECT * FROM [Store: Units in Crates] WHERE "
        If Not IsNull(firstCrate![Unit number]) Then
            sql = sql & "[Unit number] = " & firstCrate![Unit number]
        Else
             sql = sql & "([Unit number] is null)"
        End If
       
       If Not IsNull(firstCrate![FindSampleLetter]) Then
            sql = sql & " AND [FindSampleLetter] = '" & firstCrate![FindSampleLetter] & "'"
        Else
            sql = sql & " AND ([FindSampleLetter] is null)"
        End If
        
        If Not IsNull(firstCrate![FindNumber]) Then
            sql = sql & " AND [FindNumber] = " & firstCrate![FindNumber]
        Else
            sql = sql & "AND ([FindNumber] is null)"
        End If
    
        If Not IsNull(firstCrate![SampleNumber]) Then
            sql = sql & " AND [SampleNumber] = " & firstCrate![SampleNumber]
        Else
            sql = sql & " AND ([SampleNumber] is null)"
        End If
        
        'sql = sql & " AND CrateLetter = 'PG' AND CrateNumber = 1001;"
        sql = sql & " AND CrateLetter = '" & Me!cboSecondCrate.Column(1) & "' AND CrateNumber = " & Me!cboSecondCrate.Column(2) & ";"
        ''Debug.Print sql
    
        Set SecondCrate = mydb.OpenRecordset(sql, dbOpenSnapshot)
            
            If SecondCrate.BOF And SecondCrate.EOF Then
                'insert temp table
                sql = "INSERT INTO [Temp_Store: Units in Crates] "
                sql = sql & "SELECT * FROM [Store: Units in Crates] "
                sql = sql & " WHERE [RowID] = " & firstCrate![rowID] & ";"
                DoCmd.RunSQL sql
            
            End If
        
        SecondCrate.Close
        Set SecondCrate = Nothing
    firstCrate.MoveNext
    Loop
End If

sql = ""

firstCrate.Close
Set firstCrate = Nothing
mydb.Close
Set mydb = Nothing

Me!frm_subform_crate_comparison.Requery
Me!frm_subform_crate_comparison.Visible = True

Exit Sub

err_cmdReturn:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
    Me!frm_subform_crate_comparison.Visible = False
    

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdPrint_Click()
On Error GoTo Err_cmdPrint_Click

    SecondCrate = Me!cboSecondCrate 'public var declared in globals mod.
    DoCmd.OpenReport "R_crate_comparison", acViewPreview, Me!cboSecondCrate
    'Reports![R_crate_comparison].SetFocus

Exit Sub

Err_cmdPrint_Click:
    Call General_Error_Trap
    Exit Sub
  
End Sub
