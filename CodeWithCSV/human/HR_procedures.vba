Option Compare Database
Option Explicit

Sub ReturnToMenu(frm)
'global return to main menu routine - SAJ
On Error GoTo err_menu

    DoCmd.OpenForm "FRM_MAINMENU"
    DoCmd.Close acForm, frm.Name
Exit Sub

err_menu:
    MsgBox Err.Description
    Exit Sub
End Sub

Sub DoRecordCheck(tblName, Unit, Individ, UnitfldName)
'check if a skeleton number exists in the table related to the form being opened, if not
'add it
'SAJ
On Error GoTo err_check
    Dim mydb As DAO.Database, myrs As DAO.Recordset, sql As String, sql2 As String
    sql = "SELECT [" & UnitfldName & "], [Individual Number] FROM [" & tblName & "] WHERE [" & UnitfldName & "] = " & Unit & " AND [Individual Number] = " & Individ & ";"
    Set mydb = CurrentDb
    Set myrs = mydb.OpenRecordset(sql)
    
    If myrs.BOF And myrs.EOF Then
        sql2 = "INSERT INTO [" & tblName & "] ([" & UnitfldName & "], [Individual Number]) VALUES (" & Unit & "," & Individ & ");"
        DoCmd.RunSQL sql2
    End If
    
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing

Exit Sub

err_check:
    MsgBox Err.Description
    Exit Sub
End Sub

Sub SortOutButtons(frm As Form)
On Error GoTo err_Sort

If frm![cboAgeCategory] <> "" Then
    
        '2010 or =9 added for prenatal
        If frm![cboAgeCategory] = 0 Or frm![cboAgeCategory] = 9 Then
            If frm.Name <> "FRM_Permanent_Teeth" Then frm![CmdOpenNeonateFrm].Enabled = True
            frm![CmdOpenJuvenileFrm].Enabled = False
            If frm.Name <> "FRM_Deciduous_Teeth" Then frm![CmdOpenAdultFrm].Enabled = False
        ElseIf frm![cboAgeCategory] = 1 Or frm![cboAgeCategory] = 2 Or frm![cboAgeCategory] = 3 Then
            If frm.Name <> "FRM_Permanent_Teeth" Then frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = True
            If frm.Name <> "FRM_Deciduous_Teeth" Then frm![CmdOpenAdultFrm].Enabled = False
        ElseIf frm![cboAgeCategory] = 4 Or frm![cboAgeCategory] = 5 Or frm![cboAgeCategory] = 6 Or frm![cboAgeCategory] = 7 Or frm![cboAgeCategory] = 8 Then
            'march 2009 - SH decide all unknowns treated as adult - category 8 added to end of if statement above
            If frm.Name <> "FRM_Permanent_Teeth" Then frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = False
            If frm.Name <> "FRM_Deciduous_Teeth" Then frm![CmdOpenAdultFrm].Enabled = True
        Else
            If frm.Name <> "FRM_Permanent_Teeth" Then frm![CmdOpenNeonateFrm].Enabled = True
            frm![CmdOpenJuvenileFrm].Enabled = True
            If frm.Name <> "FRM_Deciduous_Teeth" Then frm![CmdOpenAdultFrm].Enabled = True
        End If
        
   Else
            If frm.Name <> "FRM_Permanent_Teeth" Then frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = False
            If frm.Name <> "FRM_Deciduous_Teeth" Then frm![CmdOpenAdultFrm].Enabled = False
    End If

Exit Sub
err_Sort:
    MsgBox Err.Description
    Exit Sub
End Sub


Function GetSkeletonAge(frm As Form)
'******************************************************************************
' Return age category of current skeleton
'
' SAJ v9
'******************************************************************************
On Error GoTo err_GetSkeletonAge

Dim retVal
retVal = ""
If DBName <> "" Then
    Dim mydb As DAO.Database, myrs As DAO.Recordset
    Dim sql
    Set mydb = CurrentDb()
    
    sql = "SELECT [age category] FROM [Q_Retrieve_Age_of_Skeleton] WHERE [unit number] = " & frm![txtUnit] & " AND [individual number] = " & frm![txtIndivid] & ";"
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    
     '' Dim myrs As New ADODB.Recordset
   '' myrs.Open sql, CurrentProject.Connection, adOpenKeySet, adLockOptimistic
    
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        retVal = retVal & myrs![age category]
    End If
    
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    
  
Else
    retVal = retVal & "X"
End If

GetSkeletonAge = retVal

Exit Function
err_GetSkeletonAge:
    Call General_Error_Trap
End Function

