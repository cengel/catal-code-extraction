Option Compare Database
Option Explicit

Private Sub Ctl1_mm_fraction_NotInList(NewData As String, Response As Integer)
'*************************************************************
' Add a new value to the 1mm fraction list - prompting user 1st
' SAJ
'**************************************************************
On Error GoTo err_1mmfraction

Dim retVal
retVal = MsgBox("This Fraction value has not been used in this field before, are you sure you wish to add it to the list?", vbInformation + vbYesNo, "New Fraction value")
If retVal = vbYes Then
    'allow value, as this is distinct query based list we must save the record
    'first but need to turn off limittolist first to be able to do so an alternative
    'way to do this would be to dlookup on entry when not limited
    'to list but this method is quicker (but messier) as not require DB lookup 1st
    Response = acDataErrContinue
    Me![1 mm fraction].LimitToList = False 'turn off limit to list so record can be saved
    DoCmd.RunCommand acCmdSaveRecord 'save rec
    Me![1 mm fraction].Requery 'requery combo to get new value in list
    Me![1 mm fraction].LimitToList = True 'put back on limit to list
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub

err_1mmfraction:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Ctl4_mm_fraction_NotInList(NewData As String, Response As Integer)
'*************************************************************
' Add a new value to the 4mm fraction - prompting user 1st
' SAJ
'**************************************************************
On Error GoTo err_4mmfraction

Dim retVal
retVal = MsgBox("This Fraction value has not been used in this field before, are you sure you wish to add it to the list?", vbInformation + vbYesNo, "New Fraction value")
If retVal = vbYes Then
    'allow value, as this is distinct query based list we must save the record
    'first but need to turn off limittolist first to be able to do so an alternative
    'way to do this would be to dlookup on entry when not limited
    'to list but this method is quicker (but messier) as not require DB lookup 1st
    Response = acDataErrContinue
    Me![4 mm fraction].LimitToList = False 'turn off limit to list so record can be saved
    DoCmd.RunCommand acCmdSaveRecord 'save rec
    Me![4 mm fraction].Requery 'requery combo to get new value in list
    Me![4 mm fraction].LimitToList = True 'put back on limit to list
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub

err_4mmfraction:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()
'*****************************************************************************************
'some fields have been replaced after 2004, toggle between fields to hide/show depending
'on the year
' SAJ
'*****************************************************************************************
On Error GoTo err_current

'new fields for 2006
    If Me![YearScanned] >= 2006 Then
        Me![1 mm culm nodes].Enabled = False
        Me![1 mm reed culm node].Enabled = True
        Me![1 mm cereal culm node].Enabled = True
    Else
        Me![1 mm culm nodes].Enabled = True
        Me![1 mm reed culm node].Enabled = False
        Me![1 mm cereal culm node].Enabled = False
    End If
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

