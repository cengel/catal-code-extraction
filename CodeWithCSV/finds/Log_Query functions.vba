Option Compare Database
Option Explicit
Dim dbSource As Database
Dim rstUnitlist As Recordset

Private Sub Command9_Click()

Dim stDocName As String
    Dim SQLstring, areapart, datepart, unitpart, materialpart, XFpart, excavpart, priopart, caption As String

 
    'Create SQL statement
    
    If Me.Area <> "" Then
    areapart = " [Area] ='" & Me.[Area] & "'"
    Else
    areapart = " [Area] Like '*' "
    End If
    
    If Me.Date <> "" Then
    datepart = " And [date] =#" & Me.[Date] & "#"
    Else
    datepart = " And [date] Like '*' "
    End If
    
    
    If Me.Unit <> "" And Not Me.unit_ex Then
        unitpart = " And [Unit] =" & Me.[Unit]
        ElseIf Me.[Unit] <> "" And Me.unit_ex Then
        unitpart = " And Not [Unit] =" & Me.[Unit]
        Else
        unitpart = ""
    End If
    
    If Me.Material <> "" And Not Me.mat_ex Then
        materialpart = " And [Material] ='" & Me.[Material] & "'"
        ElseIf Me.Material <> "" And Me.mat_ex Then
        materialpart = " And Not [Material] ='" & Me.[Material] & "'"
        Else
        materialpart = ""
    End If
   
    If Me.X_Find <> "" And Not Me.x_ex Then
        XFpart = " And [Find no] ='" & Me.[X Find] & "'"
        ElseIf Me.[X Find] <> "" And Me.x_ex Then
        XFpart = " And Not [Find no] ='" & Me.[X Find] & "'"
        Else
        XFpart = ""
    End If
    
    If Me.Excav <> "" Then
        excavpart = " And [Excavator] ='" & Me.[Excav] & "'"
        Else
        excavpart = ""
    End If
    
    If Me.Priority Then
        priopart = " And [Priority] = -1"
        Else
        priopart = ""
    End If
    
    
    SQLstring = "SELECT * FROM [Log: Details] WHERE " & areapart & datepart & unitpart & materialpart & XFpart & excavpart & priopart
    'Debug.Print SQLstring
        
    'stDocName = "Exca: Unit list"
    
    DoCmd.OpenForm "Log: Details List", acFormDS
    Forms![Log: Details List].caption = "List of units"
    '--------format header
    caption = "List of units"
    
    If Me.Area <> "" Then
    caption = caption & " [Area] =" & Me.[Area]
    End If
    If Me.Date <> "" Then
    caption = caption & " and [date] =" & Me.[Date]
    End If
    Forms![Log: Details List].caption = caption
    '----------end header
    
    Forms![Log: Details List].RecordSource = SQLstring


End Sub

Private Sub Data_Category_BeforeUpdate(Cancel As Integer)

End Sub



Private Sub Location_BeforeUpdate(Cancel As Integer)

End Sub


Private Sub Command10_Click()
On Error GoTo Err_Command10_Click

    Dim stDocName As String
    Dim SQLstring, areapart, datepart, situpart, locpart, descripart, matpart, depopart, basalpart As String

 
    'Create select statement
    
    'construct SQL statement
    If Me.Area <> "" Then
    areapart = " [Area] ='" & Me.[Area] & "'"
    Else
    areapart = " [Area] Like '*' "
    End If
    
    If Me.Date <> "" Then
    datepart = " And [date] =#" & Me.[Date] & "#"
    Else
    datepart = " And [date] Like '*' "
    End If
    
    '---------------------------------------------------------------
   
        
    SQLstring = "SELECT * FROM [Log: Sheets] WHERE " & areapart & datepart
    'Debug.Print SQLstring
    
    stDocName = "Log: Search Results"
    DoCmd.OpenForm stDocName
    Forms![Log: Search Results].RecordSource = SQLstring

Exit_Command10_Click:
    Exit Sub

Err_Command10_Click:
    MsgBox Err.Description
    Resume Exit_Command10_Click
    
End Sub
Private Sub Close_Click()
On Error GoTo Err_close_Click


    DoCmd.Close

Exit_close_Click:
    Exit Sub

Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
    
End Sub

Private Sub log_sheet_Click()
On Error GoTo Err_log_sheet_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Log: Daily Log Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_log_sheet_Click:
    Exit Sub

Err_log_sheet_Click:
    MsgBox Err.Description
    Resume Exit_log_sheet_Click
    
End Sub
