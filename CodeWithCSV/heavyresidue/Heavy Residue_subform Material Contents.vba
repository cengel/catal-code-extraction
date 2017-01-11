Option Compare Database
Option Explicit
Dim current4percent
Dim current2percent
Dim current1percent


Sub SetUpFields()
'set up view of fields - pot and clay ball only need 4mm entry, stone only need 4mm presence/abscence
On Error GoTo err_setupfield

    'If LCase(Me![Material]) = "stone" Or LCase(Me![Material]) = "worked stone" Then
    '    Me![chkStonePresent].Visible = True
    '    Me![4 % sorted].TabStop = False
    '    Me![4 Weight].TabStop = False
    'Else
    '    Me![chkStonePresent].Visible = False
    '    Me![4 % sorted].TabStop = True
    '    Me![4 Weight].TabStop = True
    'End If
    
    'If LCase(Me![Material]) = "pottery" Or LCase(Me![Material]) = "clay ball" Or LCase(Me![Material]) = "stone" Or LCase(Me![Material]) = "worked stone" Then
    ' MR added bone 4 %, elseif for bone diagnostic for Slobo 27/7/2006
    'this change above was not in the version brought back to cambridge end 2007, rescued from current download dir on site server in 2007 by SAJ
    'Changed bone-behaviour - is now enabled for 2mm an 1mm; DL 2016
    If LCase(Me![Material]) = "pottery" Or LCase(Me![Material]) = "clay ball" Or LCase(Me![Material]) = "stone" Or LCase(Me![Material]) = "worked stone" Then
        Me![2 % sorted].Enabled = False
        Me![2 % sorted].Locked = True
        Me![2 % sorted].TabStop = False
       ' Me![2 % sorted].BackColor = Me.Section(0).BackColor
        Me![2 Weight].Enabled = False
        Me![2 Weight].Locked = True
        Me![2 Weight].TabStop = False
       ' Me![2 Weight].BackColor = Me.Section(0).BackColor
        Me![1 % sorted].Enabled = False
        Me![1 % sorted].Locked = True
        Me![1 % sorted].TabStop = False
       ' Me![1 % sorted].BackColor = Me.Section(0).BackColor
        Me![1 Weight].Enabled = False
        Me![1 Weight].Locked = True
        Me![1 Weight].TabStop = False
       ' Me![1 Weight].BackColor = Me.Section(0).BackColor
    
    'this a new part of MR's change rescued 2007
    ElseIf LCase(Me![Material]) = "bone diagnostic" Then
        Me![4 % sorted].Enabled = False
        Me![4 % sorted].Locked = True
        Me![4 % sorted].TabStop = False
        Me![4 Weight].Enabled = False
        Me![4 Weight].Locked = True
        Me![4 Weight].TabStop = False
    Else
        'the 4 mm fields added to this bit due to elseif above
        Me![4 % sorted].Enabled = True
        Me![4 % sorted].Locked = False
        Me![4 % sorted].TabStop = True
        Me![4 Weight].Enabled = True
        Me![4 Weight].Locked = False
        Me![4 Weight].TabStop = True
    
        'original
        Me![2 % sorted].Enabled = True
        Me![2 % sorted].Locked = False
        Me![2 % sorted].TabStop = True
       ' Me![2 % sorted].BackColor = 16777215
        Me![2 Weight].Enabled = True
        Me![2 Weight].Locked = False
        Me![2 Weight].TabStop = True
       ' Me![2 Weight].BackColor = 16777215
        Me![1 % sorted].Enabled = True
        Me![1 % sorted].Locked = False
        Me![1 % sorted].TabStop = True
       ' Me![1 % sorted].BackColor = 16777215
        Me![1 Weight].Enabled = True
        Me![1 Weight].Locked = False
        Me![1 Weight].TabStop = True
       ' Me![1 Weight].BackColor = 16777215
    End If

    'new season 2007 - there are default values % sorted for diff material types
    'but if the a % sorted is typed in Slobo wants this carried down the list for any type of material
    'the after update routine on the sorted fields stores the value in a global to use
    'july 2008 after saj left 2007 and begining of 2008 Betsa reported that after changing the 100%
    'value when you enter a new HR record you get stuck in a null value/insert error loop that
    'means you have to keep pressing ok (you can carry on but its annoying.
    'I think this is because these 3 lines trigger a new materials record to be created before the
    'unit,sample and flot fields have been filled out so...
    If Not IsNull(Me![Unit]) Then
        If current4percent <> "" Then Me![4 % sorted] = current4percent
        If current2percent <> "" Then Me![2 % sorted] = current2percent
        If current1percent <> "" Then Me![1 % sorted] = current1percent
        'MsgBox "should now be: " & Me![2 % sorted]
    Else
        'new main HR record so blank these values
        current4percent = ""
        current2percent = ""
        current1percent = ""
    End If
Exit Sub

err_setupfield:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Ctl1___sorted_AfterUpdate()
'new season 2007 - if % sorted is changed in one record hold onto the value and make the the value for this column
'for every new entry
On Error GoTo err_ct11

    'form var updated with value
    current1percent = Me![1 % sorted]
    
    'make sure its not 0
    If Me![1 % sorted] = 0 Then
        MsgBox "Invalid entry - 0 not allowed. Update cancelled", vbCritical, "Invalid Data"
        Me![1 % sorted] = Me![1 % sorted].OldValue
    End If
Exit Sub

err_ct11:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Ctl2___sorted_AfterUpdate()
'new season 2007 - if % sorted is changed in one record hold onto the value and make the the value for this column
'for every new entry
On Error GoTo err_ct12

    'form var updated with value
    current2percent = Me![2 % sorted]
    
    'make sure its not 0
    If Me![2 % sorted] = 0 Then
        MsgBox "Invalid entry - 0 not allowed. Update cancelled", vbCritical, "Invalid Data"
        Me![2 % sorted] = Me![2 % sorted].OldValue
    End If
Exit Sub

err_ct12:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Ctl4___sorted_AfterUpdate()
'new season 2007 - if % sorted is changed in one record hold onto the value and make the the value for this column
'for every new entry
On Error GoTo err_ct14

    'form var updated with value
    current4percent = Me![4 % sorted]
    
    'make sure its not 0
    If Me![4 % sorted] = 0 Then
        MsgBox "Invalid entry - 0 not allowed. Update cancelled", vbCritical, "Invalid Data"
        Me![4 % sorted] = Me![4 % sorted].OldValue
    End If

Exit Sub

err_ct14:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_BUpd

Me![LastUpdated] = Now()

Exit Sub

err_BUpd:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Form_Current()
'season 2006 - new request that stone and worked stone categories only have 4mm tick box
'plus pottery and clay ball only 4mm entry
On Error GoTo err_current

    Call SetUpFields
Exit Sub

err_current:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Material_AfterUpdate()

'20/7/06 Slobo you asked me to set the default value for ALL 2% sorted to 50%
'and all 1% sorted to 25% - SAJ :)
'note I also changed the hard coded default value on those fields - probably its
'necessary to have this stuff here now? Check season 2007

'Me![2 % sorted] = 25
'Me![1 % sorted] = 12.5

Me![2 % sorted] = 50
Me![1 % sorted] = 25

If Me![Material] = "Plant" Then
Me![2 % sorted] = 50
End If

If Me![Material] = "Eggshell" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Flint" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Worked Stone" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Pottery" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Clay Ball" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Clay Figurines" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Beads" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Metal" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

If Me![Material] = "Ochre" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

'SAJ new season 2006
If Me![Material] = "Bone Diagnostic" Then
Me![2 % sorted] = 50
Me![1 % sorted] = 25
End If

'saj new season 2006
Call SetUpFields

End Sub


Private Sub Material_Change()
'season 2008 - this code is triggered and is out of date! When % is changed and therefore
'carried to a new record the choice of the material is actually a ON CHANGE event rather than
'after update which meant the carried % value was being overwritten by the values here.

'so to keep % categories in one place call after update - saj 2/08/2008
Call Material_AfterUpdate

'Me![2 % sorted] = 25
'Me![1 % sorted] = 12.5


'If Me![Material] = "Plant" Then
'Me![2 % sorted] = 50
'End If

'If Me![Material] = "Eggshell" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Flint" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Worked Stone" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Pottery" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Clay Ball" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Clay Figurines" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Beads" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Metal" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'If Me![Material] = "Ochre" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If

'SAJ new season 2006
'If Me![Material] = "Bone Diagnostic" Then
'Me![2 % sorted] = 50
'Me![1 % sorted] = 25
'End If
End Sub

