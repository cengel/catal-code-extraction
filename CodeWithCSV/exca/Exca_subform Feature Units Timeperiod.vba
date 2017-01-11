Option Compare Database
Option Explicit

Private Sub TimePeriod_DblClick(Cancel As Integer)
'May 2009 - SAJ
'This is a new subform that allows the timeperiod of a units to display on the feature sheet
'Sometimes an error may have occurred and a feature have units from >1 timeperiod, to facilitate
'checking this out when the user double click the period a message will tell them the units assigned
On Error GoTo err_timeperiod

Dim sql, mydb As DAO.Database, myrs As DAO.Recordset, msg
Set mydb = CurrentDb
sql = "SELECT DISTINCT [Exca: Unit Sheet].TimePeriod, [Exca: Units in Features].In_feature, [Exca: Units in Features].Unit " & _
        "FROM [Exca: Unit Sheet] INNER JOIN [Exca: Units in Features] ON " & _
        "[Exca: Unit Sheet].[Unit Number] = [Exca: Units in Features].Unit WHERE [In_feature] = " & Me![In_feature] & " AND "
        
If IsNull(Me![TimePeriod]) Then
    sql = sql & "[timeperiod]is null;"
Else
    sql = sql & "[timeperiod] = '" & Me![TimePeriod] & "';"
End If
Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)

If Not (myrs.BOF And myrs.EOF) Then
    myrs.MoveFirst
    Do Until myrs.EOF
        If msg <> "" Then msg = msg & ", "
        msg = msg & myrs![Unit]
    myrs.MoveNext
    Loop

End If

If msg = "" Then
    MsgBox "No units have been found for this timeperiod, associated with this feature. This should not happen, please let the database administrator know", vbCritical, "Code Issue"
Else
    MsgBox "This features unit numbers that are assigned to the timeperiod '" & Me![TimePeriod] & "': " & Chr(13) & Chr(13) & msg, vbInformation, "Units for this timeperiod"
End If

myrs.Close
Set myrs = Nothing
mydb.Close
Set mydb = Nothing

Exit Sub

err_timeperiod:
    Call General_Error_Trap
    Exit Sub
End Sub
