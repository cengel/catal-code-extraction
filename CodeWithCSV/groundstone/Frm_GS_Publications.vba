Option Compare Database
Option Explicit

Private Sub cboFind_AfterUpdate()
'find the sample number
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
    
        If Me.Filter <> "" Then
            If Me.Filter <> "[GID] = '" & Me![cboFind].Column(1) & "'" Then
                MsgBox "This form was opened to only show publication records relating to " & Me.Filter & ". This action has removed the filter and all records are available to view.", vbInformation, "Filter Removed"
                Me.FilterOn = False
            End If
        End If
        DoCmd.GoToControl Me![PublicationID].Name
        DoCmd.FindRecord Me![cboFind]
   
    End If

Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Close_Click()
On Error GoTo err_close

    DoCmd.Close acForm, Me.Name

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdAddNew_Click()
'add a new publication record
On Error GoTo err_cmdAddNew
Dim retVal, getsample, getGID, getUnit, getLetter, getNum, sql

If Me![txtGID] <> "" Then
    retVal = MsgBox("Do you want to add another publication record for this GID (" & Me![txtGID] & ")?", vbYesNo + vbQuestion, "New sample")
    If retVal = vbYes Then
            getGID = Me![txtGID]
            getUnit = Me![txtUnit]
            getLetter = Me![txtLetter]
            getNum = Me![txtNum]
            Me.AllowAdditions = True
            DoCmd.RunCommand acCmdRecordsGoToNew
            Me![txtGID] = getGID
            Me![txtUnit] = getUnit
            Me![txtLetter] = getLetter
            Me![txtNum] = getNum
            Me.AllowAdditions = False
            
            'it should be that this basic record is already marked as sampled but just in case mark is so
            'as a catchall for any previous mismatches between tables
            sql = "UPDATE [GroundStone 1: Basic Data] SET [Published] = True WHERE [GID] = '" & Me![txtGID] & "';"
            DoCmd.RunSQL sql
            Exit Sub
    End If
End If

MsgBox "You now be asked for the GID fields related to this publication, you must enter them all", vbInformation, "Data Required"

getUnit = InputBox("Please enter the Unit number related to the publication:", "Unit number")
If getUnit = "" Then
    MsgBox "You cannot enter a new record without a Unit number", vbCritical, "Action Cancelled"
    Exit Sub
End If

getLetter = InputBox("Please enter the Letter (X or K) related to the object published:", "Letter")
If getLetter = "" Then
    MsgBox "You cannot enter a new record without the id letter", vbCritical, "Action Cancelled"
    Exit Sub
End If

getNum = InputBox("Please enter the object number related to the publication:", "Object number")
If getNum = "" Then
    MsgBox "You cannot enter a new record without a number", vbCritical, "Action Cancelled"
    Exit Sub
End If

Me.AllowAdditions = True
DoCmd.RunCommand acCmdRecordsGoToNew
Me![txtGID] = getUnit & "." & getLetter & getNum
Me![txtUnit] = getUnit
Me![txtLetter] = getLetter
Me![txtNum] = getNum
Me.AllowAdditions = False

'now mark the basic record for this GID as sampled
sql = "UPDATE [GroundStone 1: Basic Data] SET [Published] = True WHERE [GID] = '" & Me![txtGID] & "';"
DoCmd.RunSQL sql
Exit Sub

err_cmdAddNew:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdGoFirst_Click()
On Error GoTo Err_gofirst_Click


    DoCmd.GoToRecord , , acFirst

    Exit Sub

Err_gofirst_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoLast_Click()
On Error GoTo Err_goLast_Click


    DoCmd.GoToRecord , , acLast

    Exit Sub

Err_goLast_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoNext_Click()
On Error GoTo Err_goNext_Click


    DoCmd.GoToRecord , , acNext

    Exit Sub

Err_goNext_Click:
    Call General_Error_Trap
    
End Sub

Private Sub cmdGoPrev_Click()
On Error GoTo Err_goPrev_Click


    DoCmd.GoToRecord , , acPrevious

    Exit Sub

Err_goPrev_Click:
    Call General_Error_Trap
    
End Sub

