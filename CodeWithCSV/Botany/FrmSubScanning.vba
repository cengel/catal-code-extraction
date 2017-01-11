Option Compare Database
Option Explicit

Private Sub Ctl1_mm_random_split_NotInList(NewData As String, Response As Integer)
'*************************************************************
' Add a new value to the 1mm random split list - prompting user 1st
' SAJ
'**************************************************************
On Error GoTo err_1mmsplit

Dim retVal
retVal = MsgBox("This Random Split value has not been used on this screen before, are you sure you wish to add it to the list?", vbInformation + vbYesNo, "New Random Split value")
If retVal = vbYes Then
    'allow value, as this is distinct query based list we must save the record
    'first but need to turn off limittolist first to be able to do so an alternative
    'way to do this would be to dlookup on entry when not limited
    'to list but this method is quicker (but messier) as not require DB lookup 1st
    Response = acDataErrContinue
    Me![1 mm random split].LimitToList = False 'turn off limit to list so record can be saved
    DoCmd.RunCommand acCmdSaveRecord 'save rec
    Me![1 mm random split].Requery 'requery combo to get new value in list
    Me![1 mm random split].LimitToList = True 'put back on limit to list
Else
    'no leave it so they can edit it
    Response = acDataErrContinue
End If
Exit Sub

err_1mmsplit:
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

If Me![YearScanned] > 2004 Then
    Me![1 mm abundance scale used].Visible = False
    Me![1 mm non-random 5 ml counts].Visible = False
    
    If Me![YearScanned] = 2005 Then 'some 1mm vol vals fro 2005 so show for that year
        Me![1 mm flot volume (ml)].Visible = True
    Else
        Me![1 mm flot volume (ml)].Visible = False
    End If
    
    Me![1 mm random split].Visible = True
    Me![1 mm random subsample vol].Visible = True
    
Else
    'show old fields
    Me![1 mm abundance scale used].Visible = True
    Me![1 mm non-random 5 ml counts].Visible = True
    Me![1 mm flot volume (ml)].Visible = True
    Me![1 mm random split].Visible = False
    Me![1 mm random subsample vol].Visible = False

End If

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

