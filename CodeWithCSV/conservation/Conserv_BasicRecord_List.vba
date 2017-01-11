Option Compare Database


Private Sub find_Click()
'find conservation ref
On Error GoTo err_find_Click
Dim filterquery, filterrelated, filterreference, filtertreatment, filterlocation, filterunit, querywunits As String
Dim rs, rsMV As DAO.Recordset
Dim strSQL, strOUT, fullrefwunits, fullrefwunitsyear, fullrefwunitsid As String
Dim lnglen As Long
Dim bismultivalue As Boolean

        If Not IsNull(Me![queryfullconserv]) Then
            filterquery = filterquery & "[FullConservation_Ref] like '*" & Me![queryfullconserv] & "*' AND "
        End If
        If Not IsNull(Me![querylocation]) Then
            filterquery = filterquery & "[Location] like '*" & Me![querylocation] & "*' AND "
        End If
        If Not IsNull(Me![queryunit]) Then
        strSQL = "SELECT [ConservationRef_Year], [ConservationRef_ID] FROM [Conservation_ConservRef_RelatedTo] WHERE [ExcavationIDNumber] = " & Trim(Me.queryunit.Value)
        Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)

        bismultivalue = (rs(0).Type > 100)

        Do While Not rs.EOF
            If bismultivalue Then
                Set rsMV = rs(0).Value
                Do While Not rsMV.EOF
                    If Not IsNull(rsMV(0)) Then
                        If (CInt(rsMV(0)) < 2000) Then
                            fullrefwunitsyear = Str(CInt(rsMV(0)) - 1900)
                        Else
                            fullrefwunitsyear = Str(CInt(rsMV(0)) - 2000)
                        End If
                        If (CInt(rsMV(1)) < 10) Then
                            fullrefwunitsid = "00" & rsMV(1)
                        ElseIf (CInt(rsMV(1)) < 100) Then
                            fullrefwunitsid = "0" & rsMV(1)
                                Else
                            fullrefwunitsid = rsMV(1)
                        End If
                        strOUT = strOUT & "[FullConservation_Ref] like '*" & Trim(fullrefwunitsyear) & "." & Trim(fullrefwunitsid) & "*' OR "
                    End If
                    rsMV.MoveNext
                Loop
                Set rsMV = Nothing
            ElseIf Not IsNull(rs(0)) Then
                If (CInt(rs(0)) < 2000) Then
                    fullrefwunitsyear = Str(CInt(rs(0)) - 1900)
                Else
                    If (CInt(rs(0)) < 2010) Then
                        fullrefwunitsyear = "0" & Trim(Str(CInt(rs(0)) - 2000))
                    Else
                        fullrefwunitsyear = Str(CInt(rs(0)) - 2000)
                    End If
                End If
                If (CInt(rs(1)) < 10) Then
                    fullrefwunitsid = "00" & rs(1)
                ElseIf (CInt(rs(1)) < 100) Then
                    fullrefwunitsid = "0" & rs(1)
                Else
                    fullrefwunitsid = rs(1)
                End If
                strOUT = strOUT & "[FullConservation_Ref] like '*" & Trim(fullrefwunitsyear) & "." & Trim(fullrefwunitsid) & "*' OR "
            End If
            rs.MoveNext
        Loop
        rs.Close

        lnglen = Len(strOUT) - 4
        If lnglen > 0 Then
            querywunits = "(" & Left(strOUT, lnglen) & ")"
        End If
            filterquery = filterquery & querywunits & " AND "
        End If
        If Not IsNull(Me![queryrelatedid]) Then
            filterquery = filterquery & "[RelatedToID] = " & Me![queryrelatedid] & " AND "
        End If
        If Not IsNull(Me![queryFindType]) Then
            filterquery = filterquery & "[Find Type] like '*" & Me![queryFindType] & "*' AND "
        End If
        If Not IsNull(Me![querytreatment]) Then
            filterquery = filterquery & "[Treatment] like '*" & Me![querytreatment] & "*' AND "
        End If

If filterquery <> "" Then
    filterquery = Left(filterquery, Len(filterquery) - 5)
    DoCmd.ApplyFilter , filterquery
Else
End If
If filterlocation <> "" Then
    DoCmd.ApplyFilter , filterquery
Else
End If

Exit Sub

err_find_Click:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub findunit_AfterUpdate()
    Me.Form.Filter = "[field_concat_units] like '*" & Me.findunit.Value & "*'"
    Me.Form.FilterOn = True
End Sub

Private Sub GoTo_Click()
On Error GoTo Err_GoTo_Click

            stLinkCriteria = "[FullConservation_Ref] = '" & Me![FullConservation_Ref] & "'"
            DoCmd.OpenForm "Conserv: Basic Record", acNormal, , stLinkCriteria

Exit_GoTo_Click:
    Exit Sub

Err_GoTo_Click:
    MsgBox Err.Description
    Resume Exit_GoTo_Click
    
End Sub

