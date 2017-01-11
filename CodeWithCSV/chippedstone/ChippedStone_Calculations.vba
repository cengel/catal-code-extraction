Option Compare Database
Option Explicit

'This module deals with calculations

Sub CalcWgtL(frm As Form)
'calc wgt/l FOR HEAVY RESIDUE
'weight field / volume field  * 100 / Percent

On Error GoTo err_calcwgtL
Dim result
    
    'must check if enough field filled in to undertake calc
    If frm![Weight] <> "" And frm![txtVolume] <> "" And frm![cboPercent] <> "" Then
        'yes correct fields there - calculate
        'must use the forms names for these fields
        result = frm![Weight] / frm![txtVolume] * 100 / frm![cboPercent]
        'pass back the result to the wgt/l field on the form
        frm![txtWgt/L] = result
    End If
Exit Sub

err_calcwgtL:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub CalcCountL(frm As Form)
'calc count/l FOR HEAVY RESIDUE
'count field / volume field  * 100 / Percent

On Error GoTo err_CalcCountL
Dim result
    
    'must check if enough field filled in to undertake calc
    If frm![Count] <> "" And frm![txtVolume] <> "" And frm![cboPercent] <> "" Then
        'yes correct fields there - calculate
        'must use the forms names for these fields
        result = frm![Count] / frm![txtVolume] * 100 / frm![cboPercent]
        'pass back the result to the count/l field on the form
        frm![txtCount/L] = result
    End If
Exit Sub

err_CalcCountL:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub CalcCountLDrySeive(frm As Form)
'calc count/l FOR DRY SEIVE
'count field / volume field

On Error GoTo err_CalcCountLDS
Dim result
    
    'must check if enough field filled in to undertake calc
    If frm![Count] <> "" And frm![txtVolume] Then
        'yes correct fields there - calculate
        'must use the forms names for these fields
        result = frm![Count] / frm![txtVolume]
        'pass back the result to the count/l field on the form
        frm![txtCount/L] = result
    End If
Exit Sub

err_CalcCountLDS:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub CalcWgtLDrySeive(frm As Form)
'calc wgt/l FOR DRY SEIVE
'weight field / volume field

On Error GoTo err_calcwgtLDS
Dim result
    
    'must check if enough field filled in to undertake calc
    If frm![Weight] <> "" And frm![txtVolume] <> "" Then
        'yes correct fields there - calculate
        'must use the forms names for these fields
        result = frm![Weight] / frm![txtVolume]
        'pass back the result to the wgt/l field on the form
        frm![txtWgt/L] = result
    End If
Exit Sub

err_calcwgtLDS:
    Call General_Error_Trap
    Exit Sub
End Sub


Sub DoCalcsOnHistoricData(dataset)

On Error GoTo err_deal

Dim mydb As DAO.Database, myrs As DAO.Recordset, result, sql

Set mydb = CurrentDb()
If dataset = "" Then
    sql = "ChippedStone_Basic_Data"
Else
    sql = dataset
End If

''Set myrs = mydb.OpenRecordset("ChippedStone_Basic_Data")
Set myrs = mydb.OpenRecordset(sql)

If Not myrs.EOF And Not myrs.BOF Then

    myrs.MoveFirst
    
    Do Until myrs.EOF
        myrs.Edit
        If myrs![RetrievalMethod] = "Heavy Residue" Then
            If myrs![Weight] <> "" And myrs![HRVolume] <> "" And myrs![HRSamplePercent] <> "" Then
                result = myrs![Weight] / myrs![HRVolume] * 100 / myrs![HRSamplePercent]
                myrs![Wgt/L] = result
            End If
            
            If myrs![Count] <> "" And myrs![HRVolume] <> "" And myrs![HRSamplePercent] <> "" Then
                result = myrs![Count] / myrs![HRVolume] * 100 / myrs![HRSamplePercent]
                'pass back the result to the count/l field on the form
                myrs![Count/L] = result
            End If
        
        ElseIf myrs![RetrievalMethod] = "Dry Sieve" Then
            If myrs![Weight] <> "" And myrs![HRVolume] <> "" Then
                result = myrs![Weight] / myrs![HRVolume]
                myrs![Wgt/L] = result
            End If
        
            If myrs![Count] <> "" And myrs![HRVolume] Then
                result = myrs![Count] / myrs![HRVolume]
                myrs![Count/L] = result
            End If
        End If
        myrs.Update
    myrs.MoveNext
    Loop

End If

myrs.Close
Set myrs = Nothing

mydb.Close
Set mydb = Nothing

Exit Sub

err_deal:
    Resume Next
End Sub

Sub GatherDataAndDoCalcs()
'NEW 2010 - some data in th count and wgt /l fields never got filled out as the underlying
'soil volume information was not available at the time of entry
'this procedure finds these records and attempts to rectify it by checking if required data now exists
On Error GoTo err_gather
Dim mydb As DAO.Database, myrs As DAO.Recordset, result, getVol
Set mydb = CurrentDb()
    
    'first of all deal with dry seive - get records where count/l abd wgt/l missing as soil vol missing
    Dim sql1
    sql1 = "SELECT * FROM ChippedStone_Basic_Data WHERE (((ChippedStone_Basic_Data.Weight) Is Not Null) AND ((ChippedStone_Basic_Data.RetrievalMethod)='dry sieve') AND ((ChippedStone_Basic_Data.HRVolume) Is Null) AND ((ChippedStone_Basic_Data.Unit) Is Not Null)) OR (((ChippedStone_Basic_Data.RetrievalMethod)='dry sieve') AND ((ChippedStone_Basic_Data.HRVolume) Is Null) AND ((ChippedStone_Basic_Data.Count) Is Not Null) AND ((ChippedStone_Basic_Data.Unit) Is Not Null));"
    
    Set myrs = mydb.OpenRecordset(sql1)

    If Not myrs.EOF And Not myrs.BOF Then

        myrs.MoveFirst
    
        Do Until myrs.EOF
            
            getVol = DLookup("[Dry sieve volume]", "[Exca: Unit Sheet with relationships]", "[Unit Number] = " & myrs![Unit])
            If Not IsNull(getVol) Then
                myrs.Edit
                    myrs![HRVolume] = getVol
                myrs.Update
            End If
            myrs.MoveNext
        Loop

    End If

    myrs.Close
    Set myrs = Nothing
    
    'now deal with flot number records - get records where count/l or wgt/l and soil vol missing
    Dim sql2
    sql2 = "SELECT * FROM ChippedStone_Basic_Data WHERE (((ChippedStone_Basic_Data.Weight) Is Not Null) AND ((ChippedStone_Basic_Data.RetrievalMethod)='heavy residue') AND ((ChippedStone_Basic_Data.HRVolume) Is Null) AND ((ChippedStone_Basic_Data.FlotNum) Is Not Null)) OR (((ChippedStone_Basic_Data.RetrievalMethod)='heavy residue') AND ((ChippedStone_Basic_Data.HRVolume) Is Null) AND ((ChippedStone_Basic_Data.Count) Is Not Null) AND ((ChippedStone_Basic_Data.FlotNum) Is Not Null));"

    Set myrs = mydb.OpenRecordset(sql2)

    If Not myrs.EOF And Not myrs.BOF Then

        myrs.MoveFirst
    
        Do Until myrs.EOF
            getVol = DLookup("[Soil Volume]", "[view_ArchaeoBotany_Flot_Log]", "[Flot Number] = " & myrs![FlotNum])
            
            If Not IsNull(getVol) Then
                myrs.Edit
                    myrs![HRVolume] = getVol
                myrs.Update
            End If
            myrs.MoveNext
        Loop

    End If

    myrs.Close
    Set myrs = Nothing
    
    

mydb.Close
Set mydb = Nothing

'now redo calcs on basic data table
Call DoCalcsOnHistoricData("SELECT * FROM ChippedStone_Basic_Data WHERE (((ChippedStone_Basic_Data.[Wgt/L]) Is Null) AND ((ChippedStone_Basic_Data.RetrievalMethod)='heavy residue' Or (ChippedStone_Basic_Data.RetrievalMethod)='dry seive')) OR (((ChippedStone_Basic_Data.RetrievalMethod)='heavy residue' Or (ChippedStone_Basic_Data.RetrievalMethod)='dry seive') AND ((ChippedStone_Basic_Data.[Count/L]) Is Null));")

Exit Sub

err_gather:
    Call General_Error_Trap
    Exit Sub
End Sub

