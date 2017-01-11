Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_cmdClose

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmdClose:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdNow_Click()
'new season 2007 version v2
On Error GoTo err_cmdNow

    If (Me![cboFrom] <> "" And Me![cboFrom] <> "") And (Me![cboFrom] <> Me![cboTo]) Then
        Dim Response, sql1
        Response = MsgBox("Are you sure you wish to move the contents of Crate " & Me![cboFrom] & " into Crate " & Me![cboTo] & " (joining any existing contents " & Me![cboTo] & " already has) " & Chr(13) & " and thus emptying " & Me![cboFrom] & " of all its records?", vbQuestion + vbYesNo, "Confirm Action")
        If Response = vbYes Then
        
            'rename cboFrom units in crates to cboTo crate
            '2009 v4.2 change to structure
            'sql1 = "UPDATE [Store: Units in Crates] SET [Store: Units in Crates].[Crate Number] = '" & Me![cboTo] & "', [Store: Units in Crates].CrateNumber = " & Me![cboTo].Column(2) & ", [Store: Units in Crates].CrateLetter = '" & Me![cboTo].Column(1) & "' WHERE [Crate Number] ='" & Me![cboFrom] & "';"
            sql1 = "UPDATE [Store: Units in Crates] SET [Store: Units in Crates].CrateNumber = " & Me![cboTo].Column(2) & ", [Store: Units in Crates].CrateLetter = '" & Me![cboTo].Column(1) & "' WHERE [CrateLetter] ='" & Me![cboFrom].Column(1) & "' AND [CrateNumber] = " & Me![cboFrom].Column(2) & ";"
                On Error Resume Next
                Dim mydb As DAO.Database, wrkdefault As Workspace, myq As QueryDef
                Set wrkdefault = DBEngine.Workspaces(0)
                Set mydb = CurrentDb
        
                ' Start of outer transaction.
                wrkdefault.BeginTrans
                Set myq = mydb.CreateQueryDef("")
                myq.sql = sql1
                myq.Execute
                
                myq.Close
                Set myq = Nothing
            
                If Err.Number = 0 Then
                    wrkdefault.CommitTrans
                    MsgBox "Crate Contents Moved Successfully"
                    
                Else
                    wrkdefault.Rollback
                    MsgBox "A problem has occured and the move has been cancelled. The error message is: " & Err.Description
                End If

                mydb.Close
                Set mydb = Nothing
                wrkdefault.Close
                Set wrkdefault = Nothing
        
        End If
    
    Else
        MsgBox "You must select a valid To and From crate to proceed", vbExclamation, "Invalid Crate Selection"
        Exit Sub
    End If


Exit Sub

err_cmdNow:
    Call General_Error_Trap
    Exit Sub
End Sub



Private Sub Form_Open(Cancel As Integer)
'new season 2007

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
   Exit Sub

End Sub
