Option Compare Database

Option Explicit

Sub SortBodySherd()

'first deal with surface treatment - sep from warecode
Dim mydb As DAO.Database, myrs As DAO.Recordset, WareCode As String, sql As String, streatment, newwarecode As String
Set mydb = CurrentDb
Set myrs = mydb.OpenRecordset("ceramics_body_sherds", dbOpenDynaset)

If Not (myrs.BOF And myrs.EOF) Then
    myrs.MoveFirst
    Do Until myrs.EOF
        WareCode = myrs![WARE CODE]
        If InStr(WareCode, ",") Then
            Debug.Print myrs![Unit] & myrs![WARE CODE]
        Else
            streatment = Right(myrs![WARE CODE], 1)
            newwarecode = Left(myrs![WARE CODE], Len(myrs![WARE CODE]) - 1)
            If IsNumeric(streatment) Then
                sql = "INSERT INTO [ceramics_body_sherd_surfacetreatment] ([unit], [ware code], [surfacetreatment]) VALUES (" & myrs![Unit] & ", '" & newwarecode & "'," & streatment & ");"
                DoCmd.RunSQL sql
            End If
            myrs.Edit
                myrs![WARE CODE] = newwarecode
            myrs.Update
            
        End If
        myrs.MoveNext
    Loop
End If

myrs.Close
Set myrs = Nothing
mydb.Close
Set mydb = Nothing

End Sub
