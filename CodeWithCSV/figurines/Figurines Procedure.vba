Option Compare Database
Option Explicit

Sub DuplicateCurrentRecord()

On Error GoTo err_dup

    'get user to input new GID - if not quit
    
    
        'check new GID does not already exist



Exit Sub

err_dup:
    Call General_Error_Trap
    Exit Sub

End Sub

Function FindLastGIDforUnit(getUnit)
'find the last GID for a unit
'more complicated than it should be as we the id number is not split into component parts
On Error GoTo err_lastgid

Dim mydb As DAO.Database, myrs As DAO.Recordset, sql, returnVal

If getUnit <> "" Then
    'first tried splitting the id number into unit, letter, num on the fly via Q_Figurine_GID_Split
    'but could not sort this query correctly to get the last number - so hack time - use this query
    'to feed a temporary table which is dumped and refilled every time the user calls this function
    sql = "DELETE * FROM [Temp_Fig_GID_Split];"
    DoCmd.RunSQL sql
    
    sql = "INSERT INTO Temp_Fig_GID_Split ([Id number], [Unit], Letter, Num) SELECT [Id number], [Unit], Letter, Num FROM Q_Figurine_GID_Split;"
    DoCmd.RunSQL sql
    
    
    'data ready now can extract our unit
    sql = "SELECT Max(Temp_Fig_GID_Split.Num) AS LastNum FROM Temp_Fig_GID_Split WHERE (Temp_Fig_GID_Split.Unit = " & getUnit & ") And (Temp_Fig_GID_Split.Letter = 'H') ORDER BY Max(Temp_Fig_GID_Split.Num);"
    
    Set mydb = CurrentDb
    Set myrs = mydb.OpenRecordset(sql, dbOpenDynaset)
    
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveLast
        ''MsgBox myrs![LastNum]
        returnVal = myrs![LastNum]
    Else
        returnVal = "Not found"
    End If
    
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
End If

FindLastGIDforUnit = returnVal
Exit Function

err_lastgid:
    Call General_Error_Trap
    Exit Function

End Function
