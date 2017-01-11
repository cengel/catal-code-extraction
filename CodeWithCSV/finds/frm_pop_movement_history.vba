Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
'close this pop up
On Error GoTo err_close
    Forms![Store: Crate Register]![Store: subform Units in Crates].Requery
    Forms![Store: find unit in crate2].Requery
    'Forms![Store: Crate Register]![Store: subform Units in Crates].Refresh
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
'2011 track movement in the tracker table
        Dim sql, crateLet, crateNum, counter, strLen
        'the crate letter and number are held in one field - split it into its parts
        strLen = Len(Me![MovedFromCrate])
        counter = 1
        Do Until counter = strLen + 1
            'loop thro str
            If IsNumeric(Mid(Me![MovedFromCrate], counter, 1)) Then
                'when hit a number stop as this is the begining of the crate number
                'crate letter is the part of the field to the left of this number
                crateLet = Left(Me![MovedFromCrate], counter - 1)
                'crate number is the part to the right, start at number and work out len to bring back
                crateNum = Mid(Me![MovedFromCrate], counter, strLen - (counter - 1))
                Exit Do
            End If
        
        counter = counter + 1
        Loop
        
        'insert the movement into the tracker
        sql = "INSERT INTO [Store: Crate Movement by Teams] ([OriginalRowID], [Unit Number], [FindSampleLetter], [FindNumber], [SampleNumber], [MovedFromCrate], [MovedToCrate], [MovedBy], [MovedOn]) "
        sql = sql & "SELECT [RowID] as OriginalRowID, [Unit number], [FindSampleLetter], [FindNumber], [SampleNumber], [CrateLetter] & [CrateNumber] as MovedFromCrate, '" & Me![MovedFromCrate] & "' as MovedToCrate, '" & logon & "', '" & Now & "' "
        sql = sql & " FROM [Store: Units in Crates] "
        sql = sql & " WHERE [RowID] = " & Me![OriginalRowID] & ";"
        DoCmd.RunSQL sql
        
        'change the crate number/letter to the movedtocrate value (ie: previous value)
        sql = "UPDATE [Store: Units in Crates] SET [CrateNumber] = " & crateNum & ", [CrateLetter] = '" & crateLet & "' WHERE [RowID] = " & Me![OriginalRowID] & ";"
        DoCmd.RunSQL sql
        
        'Me.Requery
        'Me![cboMoveCrate] = ""
        MsgBox "Move has been successful from " & Me![MovedToCrate] & " back to " & Me![MovedFromCrate]
        Me.Requery
End Sub
