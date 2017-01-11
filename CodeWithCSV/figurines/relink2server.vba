Option Compare Database

'Sub to relink tables to local/online server when going on- or offsite
'Only necessary to change DSN in tdf.connect

Private Sub relinkTables()
Dim tdf As DAO.TableDef
 
    For Each tdf In CurrentDb.TableDefs
        ' check if table is a linked table
        If Len(tdf.Connect) > 0 Then
            'If tdf.Connect = "ODBC;DSN=catalhoyuk;WSID=CATAL;DATABASE=catalhoyuk;" Then
            '    tdf.Connect = "ODBC;DSN=catal;WSID=CATAL;DATABASE=Catalhoyuk;"
            '    tdf.RefreshLink
            'ElseIf tdf.Connect = "ODBC;DSN=catal;WSID=CATAL;DATABASE=Catalhoyuk;" Then
            '    tdf.Connect = "ODBC;DSN=catalhoyuk;WSID=CATAL;DATABASE=catalhoyuk;"
            '    tdf.RefreshLink
            'End If
            
            'relink to OnSite usage
            tdf.Connect = "ODBC;DSN=Catalhoyuk;WSID=CATAL;DATABASE=Catalhoyuk;"
            tdf.RefreshLink
            
            'relink to OffSite usage
            'tdf.Connect = "ODBC;DSN=Catal;WSID=CATAL;DATABASE=Catalhoyuk;"
            'tdf.RefreshLink
            Debug.Print tdf.Name & ": " & tdf.Connect
        End If
    Next
 
End Sub
