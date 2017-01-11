Option Compare Database   'Use database order for string comparisons
Option Explicit
'**************************************************
' This form is new in v9.2 - SAJ
'**************************************************




Private Sub cmdAddNew_Click()
'v9.2 SAJ - add a new record
On Error GoTo err_cmdAddNew_Click

    DoCmd.RunCommand acCmdRecordsGoToNew

Exit Sub

err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cboFind_Click()
On Error GoTo err_cboFind
    
    If Me![cboFind] <> "" Then
        DoCmd.GoToControl "txtUnitNumber"
        DoCmd.FindRecord Me![cboFind]
   
    End If
Exit Sub

err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdReNumber_Click()
'new season 2007 - saj
'auto re-num of a unit in all of its sub tables
'PROBLEMS with transactions so set up all keys (couldn't check on creation so is some
'data cleaning required) - only couple of tables must be done manually
On Error GoTo err_cmdReNumber

    Dim retval, findUnit, sql, response, msg
    retval = InputBox("Please enter the new number for Unit " & Me![txtUnitNumber] & "?", "Enter new unit number")
    If retval <> "" Then
        'check valid entry
        If Not IsNumeric(retval) Then
            MsgBox "Invalid Unit number, please try again", vbExclamation, "Action Cancelled"
            Exit Sub
        End If
        
        'now check if new unit already exists
        findUnit = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit number] = " & retval)
        If Not IsNull(findUnit) Then
            MsgBox "Sorry but the unit number " & retval & " already exists. You must delete it first before you can alter " & Me![txtUnitNumber], vbExclamation, "Unit already exists"
            Exit Sub
        Else
            'new unit number does not exist so allow alteration
            msg = "Are you quite sure that you want to renumber Unit " & Me![txtUnitNumber] & " to " & retval & "?"
            response = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Unit Re-Number")
            If response = vbYes Then
                'TRANSACTION FAILING AS TWO CALLS TO SAME TABLE _ SECOND FAILS - TRIED TRANS IN TRANS BUT FAIL
                
                'On Error Resume Next
                Dim mydb As DAO.Database, wrkdefault As Workspace, wrkdefault1 As Workspace
                'Dim mydb1 As DAO.Database
                'Set wrkdefault = DBEngine.Workspaces(0)
                Set mydb = CurrentDb
                'Set mydb1 = CurrentDb
        
                '' Start of outer transaction.
                'wrkdefault.BeginTrans
        
                'change unit number
                'sql = "UPDATE [Exca: Unit Sheet] SET [Exca: Unit Sheet].[Unit Number] = " & retVal & " WHERE [Exca: Unit Sheet].[Unit Number]=" & Me![txtUnitNumber] & ";"
                Call RenumARecord("Exca: Unit Sheet", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                'these should all be done by cascade - but do here to
                Call RenumARecord("Exca: Units in Features", "Unit", retval, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: Units in Spaces", "Unit", retval, Me![txtUnitNumber], False, mydb)
                ' no longer a table 2009 Call RenumARecord("Exca: Units in Buildings", "Unit", retVal, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: Unit Interpretive Categories", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: Unit Data Categories", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: Dimensions", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
            
                If LCase(Me![txtCategory]) = "skeleton" Then
     'begin begin trans1
     '
     '              Set wrkdefault1 = DBEngine.Workspaces(0)
     '               wrkdefault1.BeginTrans
     '
     '               Call RenumARecord("Exca: skeletons same as", "skell_unit", retVal, Me![txtUnitNumber], False, mydb1)
     '
     '               If Err.Number = 0 Then
     '
     '                   wrkdefault1.CommitTrans
     '                   MsgBox "Renumbering has been successful"
     '
     '               Else
     '                   wrkdefault1.Rollback
     '                   MsgBox "A problem has occured and the re-numbering has been cancelled. The error message is: " & Err.Description
     '               End If
              
                    Call RenumARecord("Exca: Skeleton data", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
     ' end begin trans1
                    Call RenumARecord("Exca: skeletons same as", "skell_unit", retval, Me![txtUnitNumber], False, mydb)
                    Call RenumARecord("Exca: skeletons same as", "to_unit", retval, Me![txtUnitNumber], False, mydb)


                ElseIf LCase(Me![txtCategory]) = "cut" Then
                    Call RenumARecord("Exca: descriptions cut", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                Else
                    Call RenumARecord("Exca: descriptions layer", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                End If
           
               Call RenumARecord("Exca: stratigraphy", "Unit", retval, Me![txtUnitNumber], False, mydb)
               Call RenumARecord("Exca: stratigraphy", "to_Units", retval, Me![txtUnitNumber], True, mydb)
       
                Call RenumARecord("Exca: graphics list", "Unit", retval, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: samples", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
                Call RenumARecord("Exca: X-Finds: Basic data", "Unit Number", retval, Me![txtUnitNumber], False, mydb)
        
                ''Call RenumARecord("Exca: Unit Sheet", "Unit Number", retVal, Me![txtUnitNumber], False, mydb)
    
                'If Err.Number = 0 Then
                '    wrkdefault.CommitTrans
                '    MsgBox "Renumbering has been successful"
                '    Me.Requery
                '    Me![cboFind].Requery
                'Else
                '    wrkdefault.Rollback
                '    MsgBox "A problem has occured and the re-numbering has been cancelled. The error message is: " & Err.Description
                'End If
                'mydb1.close
                'Set mydb1 = Nothing
                mydb.Close
                Set mydb = Nothing
                'wrkdefault1.close
                'Set wrkdefault1 = Nothing
                'wrkdefault.close
                'Set wrkdefault = Nothing
                MsgBox "Renumbering has been successful"
            Else
                MsgBox "Re-numbering cancelled", vbInformation, "Action Cancelled"
    
            End If
        End If

    Else
        MsgBox "No unit number entered, action cancelled"
    End If

Exit Sub

err_cmdReNumber:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Excavation_Click()
'v9.2 SAJ - close the form
    DoCmd.Close acForm, Me.Name
End Sub



Private Sub cmdEdit_Click()
'v9.2 SAJ - check if user can edit this record
' The level value is used in 3 potential places on the Space Sheet - certain level
' Uncertain level start and uncertain level end so must check all these values to see
' if this level is used before allow edit
' At present not offering global edits - this can be extended to offer this if required

On Error GoTo Err_cmdEdit_Click

Dim checkValidAction, checkValidAction2, checkValidAction3, retval

    'check space sheet level field
    checkValidAction = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "Level", "edit")
    
    If checkValidAction = "ok" Then
        ''check space sheet uncertain level start field
        checkValidAction2 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelStart", "edit")
        
        If checkValidAction2 = "ok" Then
        'check space sheet uncertain level end field
            checkValidAction3 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelEnd", "edit")
        
            If checkValidAction3 = "ok" Then
                'edit action can go ahead - at present simply offer an input box for this
                retval = InputBox("No records refer to this Level (" & Me![txtLevel] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Level that you wish to replace this entry with:", "Enter edited Level")
                If retval <> "" Then
                    Me![txtLevel] = retval
                End If
                
            ElseIf checkValidAction3 = "fail" Then
                MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
            Else
                MsgBox checkValidAction3, vbExclamation, "Action Report"
            End If
        ElseIf checkValidAction2 = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            MsgBox checkValidAction2, vbExclamation, "Action Report"
        End If
        
    ElseIf checkValidAction = "fail" Then
        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
    Else
        MsgBox checkValidAction, vbExclamation, "Action Report"
    End If
    
Exit_cmdEdit_Click:
    Exit Sub

Err_cmdEdit_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub cmdDelete_Click()
'v9.2 SAJ - allow deletion of record with warning

On Error GoTo Err_cmdDelete_Click

Dim UFeature, USpace, UBuilding, UIntCat, UDataCat, UDim, UCatSpecific, UStrat, USkelSame, UGrap, USamp, UXfind
Dim retval, msg, msg1

retval = MsgBox("You have selected to delete Unit number: " & Me![txtUnitNumber] & ". The system will now check what additional data exists for this Unit and will prompt you again before deleting it." & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbCritical + vbYesNo, "Confirm Action")
If retval = vbYes Then
    
    'check units in feature
    UFeature = AdminDeletionCheck("Exca: Units in Features", "Unit", Me![txtUnitNumber], "Related to Feature", "In_Feature")
    
    'check units in spaces
    USpace = AdminDeletionCheck("Exca: Units in Spaces", "Unit", Me![txtUnitNumber], "Related to Space", "In_Space")
    
    'check units in building
    UBuilding = AdminDeletionCheck("Exca: Units in Buildings", "Unit", Me![txtUnitNumber], "Related to Building", "In_Building")
    
    'check interp cats
    UIntCat = AdminDeletionCheck("Exca: Unit Interpretive Categories", "Unit Number", Me![txtUnitNumber], "Interpretive Categories", "Interpretive Category")
    
    'check data cats
    UDataCat = AdminDeletionCheck("Exca: Unit Data Categories", "Unit Number", Me![txtUnitNumber], "Data Categories", "Data Category")
    
    'check dimensions
    UDim = AdminDeletionCheck("Exca: Dimensions", "Unit Number", Me![txtUnitNumber], "Dimensions", "Length")
    
    If LCase(Me![txtCategory]) = "skeleton" Then
        'check dimensions
        UCatSpecific = AdminDeletionCheck("Exca: skeleton data", "Unit Number", Me![txtUnitNumber], "Skeleton", "Target A - X")
        USkelSame = AdminDeletionCheck("Exca: skeletons same as", "skell_unit", Me![txtUnitNumber], "Skeleton", "To_Unit")
        USkelSame = USkelSame & AdminDeletionCheck("Exca: skeletons same as", "to_unit", Me![txtUnitNumber], "Skeleton related", "To_Unit")

    ElseIf LCase(Me![txtCategory]) = "cut" Then
        UCatSpecific = AdminDeletionCheck("Exca: descriptions cut", "Unit Number", Me![txtUnitNumber], "Cut Description", "Shape")
    Else
        UCatSpecific = AdminDeletionCheck("Exca: descriptions layer", "Unit Number", Me![txtUnitNumber], "Description", "Consistency")
    End If
    
    'check stratigraphy
    UStrat = AdminDeletionCheck("Exca: stratigraphy", "Unit", Me![txtUnitNumber], "Stratigraphy", "To_Units")
    UStrat = UStrat & AdminDeletionCheck("Exca: stratigraphy", "To_Units", Me![txtUnitNumber], "Stratigraphy", "Unit")
    
    
    UGrap = AdminDeletionCheck("Exca: graphics list", "Unit", Me![txtUnitNumber], "Graphics", "Type")
    USamp = AdminDeletionCheck("Exca: samples", "Unit number", Me![txtUnitNumber], "Samples", "Sample Number")
    UXfind = AdminDeletionCheck("Exca: X-Finds: Basic data", "Unit number", Me![txtUnitNumber], "X Finds", "GID Number")

    If UFeature <> "" Then msg = msg & UFeature & "; "
    If USpace <> "" Then msg = msg & USpace & "; "
    If UBuilding <> "" Then msg = msg & UBuilding & "; "
    If UIntCat <> "" Then msg = msg & UIntCat & "; "
    If UDataCat <> "" Then msg = msg & UDataCat & "; "
    If UDim <> "" Then msg = msg & UDim & "; "
    If UCatSpecific <> "" Then msg = msg & UCatSpecific & "; "
    If UStrat <> "" Then msg = msg & UStrat & "; "
    
    If LCase(Me![txtCategory]) = "skeleton" Then
        If USkelSame <> "" Then msg = msg & USkelSame & "; "
    End If
    If UGrap <> "" Then msg = msg & UGrap & "; "
    If USamp <> "" Then msg = msg & USamp & "; "
    If UXfind <> "" Then msg = msg & UXfind & "; "
    
    If msg = "" Then
        msg = "This Unit can safely be deleted."
    Else
        msg1 = "This Unit has the following relationships that will also be removed by the deletion - " & Chr(13) & Chr(13)
        msg = msg1 & msg
    End If
    
    msg = msg & Chr(13) & Chr(13) & "Are you quite sure that you want to permanently delete Unit " & Me![txtUnitNumber] & "?"
    retval = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retval = vbYes Then
        MsgBox "This can take a while and looks like it has hung, just let it run until a msg comes up"
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        
        ' Start of outer transaction.
        wrkdefault.BeginTrans
        
        If UFeature <> "" Then Call DeleteARecord("Exca: Units in Features", "Unit", Me![txtUnitNumber], False, mydb)
        If USpace <> "" Then Call DeleteARecord("Exca: Units in Spaces", "Unit", Me![txtUnitNumber], False, mydb)
        ' no longer a table 2009 If UBuilding <> "" Then Call DeleteARecord("Exca: Units in Buildings", "Unit", Me![txtUnitNumber], False, mydb)
        If UIntCat <> "" Then Call DeleteARecord("Exca: Unit Interpretive Categories", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UDataCat <> "" Then Call DeleteARecord("Exca: Unit Data Categories", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UDim <> "" Then Call DeleteARecord("Exca: Dimensions", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UCatSpecific <> "" Then
            
            If LCase(Me![txtCategory]) = "skeleton" Then
                Call DeleteARecord("Exca: Skeleton data", "Unit Number", Me![txtUnitNumber], False, mydb)
                Call DeleteARecord("Exca: skeletons same as", "skell_unit", Me![txtUnitNumber], False, mydb)
                Call DeleteARecord("Exca: skeletons same as", "to_unit", Me![txtUnitNumber], False, mydb)

            ElseIf LCase(Me![txtCategory]) = "cut" Then
                Call DeleteARecord("Exca: descriptions cut", "Unit Number", Me![txtUnitNumber], False, mydb)
            Else
                Call DeleteARecord("Exca: descriptions layer", "Unit Number", Me![txtUnitNumber], False, mydb)
            End If
        End If
        
        If UStrat <> "" Then
            Call DeleteARecord("Exca: stratigraphy", "Unit", Me![txtUnitNumber], False, mydb)
            Call DeleteARecord("Exca: stratigraphy", "to_Units", Me![txtUnitNumber], True, mydb)
        End If
    
        If UGrap <> "" Then Call DeleteARecord("Exca: graphics list", "Unit", Me![txtUnitNumber], False, mydb)
        If USamp <> "" Then Call DeleteARecord("Exca: samples", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UXfind <> "" Then Call DeleteARecord("Exca: X-Finds: Basic data", "Unit Number", Me![txtUnitNumber], False, mydb)
        
        Call DeleteARecord("Exca: Unit Sheet", "Unit Number", Me![txtUnitNumber], False, mydb)
    
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Deletion has been successful"
            Me.Requery
            Me![cboFind].Requery
        Else
            wrkdefault.Rollback
            'MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description

            msg = "A problem has occured and the deletion has been cancelled. " & Chr(13) & Chr(13)
            msg = msg & "SHAHINA this often fails if there is Plan/Section info, Skeleton Sameas data and Stratigraphy data present. You have delete from these tables manually first:"
            msg = msg & Chr(13) & Chr(13) & "Exca: Graphics list - all references to this unit in unit/feature number field" & Chr(13)
            msg = msg & Chr(13) & Chr(13) & "Exca: Stratigraphy - all references to this unit in both unit and to_units fields" & Chr(13)
            msg = msg & Chr(13) & Chr(13) & "(if it is a skeleton) Exca: Skeleton same as  - all references to this unit in skell_unit and to_unit fields" & Chr(13)
            msg = msg & Chr(13) & Chr(13) & "then come back here and try again...sorry...system error follows: " & Err.Description
            MsgBox msg
        End If

        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    Else
        MsgBox "Deletion cancelled", vbInformation, "Action Cancelled"
    
    End If
End If
    
    
Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
    
End Sub

Private Sub Form_Delete(Cancel As Integer)
Call cmdDelete_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'v9.2 SAJ - only adminstrators are allowed in here
On Error GoTo err_Form_Open

    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub

err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
