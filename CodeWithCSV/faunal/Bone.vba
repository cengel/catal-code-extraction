Option Compare Database
Option Explicit

Private Sub cmdBackup_Click()
'saj new to help faunal team take backups in case of central failure
On Error GoTo err_backup

Dim db As Database, nwname, msg, retVal, thepath

'thepath = left(currentdb.Name,

msg = "This facility allows you to take a local backup of the Faunal data. " & Chr(13) & Chr(13)
'msg = msg & "At present this database can only be saved into the same directory as the current file which is: " & Chr(13) & Chr(13) & CurrentDb.Name & Chr(13) & Chr(13)
msg = msg & "At present this database can only be saved into the same directory as the current file which is: " & Chr(13) & Chr(13) & "C:\Documents and Settings\All Users\Desktop\new database files" & Chr(13) & Chr(13)
msg = msg & "You can name the database whatever you like by overtyping the default entry in the next box."
MsgBox msg, vbInformation, "Backup Utility"

nwname = InputBox("Backup Database Name:", "Database Name", "Catal_Fauna_Data_Backup_" & Format(Date, "ddmmmyy") & ".mdb")

If nwname <> "" Then
    If InStr(nwname, ".mdb") = 0 Then
        nwname = nwname & ".mdb"
    End If

    nwname = "C:\Documents and Settings\All Users\Desktop\new database files\" & nwname
    
    If Dir(nwname) <> "" Then
        retVal = MsgBox("A database of this name already exists, this process will overwrite this file. Proceed anyway?", vbCritical + vbYesNo, "Overwrite Warning")
        If retVal = vbNo Then
            Exit Sub
        Else
            Kill nwname 'kill any of same name
        End If
    End If

    'create database of new name
    Set db = Workspaces(0).CreateDatabase(nwname, dbLangGeneral)
    'DoCmd.TransferDatabase acExport, "Microsoft Access", nwname, acTable, "Fauna_Bone_Faunal_Unit_Description", "Fauna_Bone_Faunal_Unit_Description", False
    
    'DoCmd.CopyObject nwname, "Fauna_Bone_Faunal_Unit_Description", acTable, "Fauna_Bone_Faunal_Unit_Description"

    'now copy each table
    DoCmd.Hourglass True
    Me![txtMsg].Visible = True
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Basic_Faunal_Data"
    DoCmd.RunSQL "SELECT Fauna_Bone_Basic_Faunal_Data.* INTO Fauna_Bone_Basic_Faunal_Data IN '" & nwname & "' FROM Fauna_Bone_Basic_Faunal_Data;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Faunal_Unit_Description"
    DoCmd.RunSQL "SELECT Fauna_Bone_Faunal_Unit_Description.* INTO Fauna_Bone_Faunal_Unit_Description IN '" & nwname & "' FROM Fauna_Bone_Faunal_Unit_Description;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Short_Faunal_Description"
    DoCmd.RunSQL "SELECT Fauna_Bone_Short_Faunal_Description.* INTO Fauna_Bone_Short_Faunal_Description IN '" & nwname & "' FROM Fauna_Bone_Short_Faunal_Description;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Cranial"
    DoCmd.RunSQL "SELECT Fauna_Bone_Cranial.* INTO Fauna_Bone_Cranial IN '" & nwname & "' FROM Fauna_Bone_Cranial;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Postcranial"
    DoCmd.RunSQL "SELECT Fauna_Bone_Postcranial.* INTO Fauna_Bone_Postcranial IN '" & nwname & "' FROM Fauna_Bone_Postcranial;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Measurements"
    DoCmd.RunSQL "SELECT Fauna_Bone_Measurements.* INTO Fauna_Bone_Measurements IN '" & nwname & "' FROM Fauna_Bone_Measurements;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Modification"
    DoCmd.RunSQL "SELECT Fauna_Bone_Modification.* INTO Fauna_Bone_Modification IN '" & nwname & "' FROM Fauna_Bone_Modification;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Artifacts"
    DoCmd.RunSQL "SELECT Fauna_Bone_Artifacts.* INTO Fauna_Bone_Artifacts IN '" & nwname & "' FROM Fauna_Bone_Artifacts;"
    Me![txtMsg] = "Backing up table:  Fauna_Bone_Contact"
    DoCmd.RunSQL "SELECT Fauna_Bone_Contact.* INTO Fauna_Bone_Contact IN '" & nwname & "' FROM Fauna_Bone_Contact;"
    Me![txtMsg] = "Backup completed to: " & db.Name
    
    retVal = MsgBox("Do you want to also backup all the code tables (there are 106 of them)?", vbQuestion + vbYesNo, "Backup Code Tables this time?")
    If retVal = vbYes Then
        Dim I, mydb As DAO.Database
        Set mydb = CurrentDb
        Dim tmptable As TableDef
        For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
            Set tmptable = mydb.TableDefs(I)
             
            If InStr(LCase(tmptable.Name), "code") <> 0 Then
                Me![txtMsg] = "Backing up table:  " & tmptable.Name
                DoCmd.RunSQL "SELECT [" & tmptable.Name & "].* INTO [" & tmptable.Name & "] IN '" & nwname & "' FROM [" & tmptable.Name & "];"
            End If
            
        Next I
        Set tmptable = Nothing
        Set mydb = Nothing
    End If
    db.Close
    Set db = Nothing
    
    DoCmd.Hourglass False
    MsgBox "Backup completed to " & nwname & " on " & Now()
    Me![txtMsg].Visible = False
Else
    MsgBox "Sorry this facility cannot run without a backup file name being entered. Please try again.", vbInformation, "Operation Cancelled"
End If
Exit Sub

err_backup:
    DoCmd.Hourglass False
    Me![txtMsg] = "Back up failed"
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
'new for version 1
On Error GoTo err_cmdQuit
    DoCmd.Quit acQuitSaveAll
Exit Sub

err_cmdQuit:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdQuit2_Click()
Call cmdQuit_Click

End Sub

Private Sub cranial_button_Click()
' This used to call macro Bone.cranial button
' Had to be translated to ensure form opened with property settings = can't add new records
' Season 2006
On Error GoTo err_cranial

    DoCmd.OpenForm "Fauna_Bone_Cranial", acNormal, , , acFormPropertySettings

Exit Sub

err_cranial:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub postcranial_button_Click()
' This used to call macro Bone.postcranial button
' Had to be translated to ensure form opened with property settings = can't add new records
' Season 2006
On Error GoTo err_postcranial

    DoCmd.OpenForm "Fauna_Bone_PostCranial", acNormal, , , acFormPropertySettings

Exit Sub

err_postcranial:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Unit_Description_Click()
'translated from macro Bone.Faunal Unit Description Button -saj
'season 2006
On Error GoTo err_unitdes

    DoCmd.OpenForm "Fauna_Bone_Faunal_Unit_Description", acNormal
    'DoCmd.Close acForm, Me.Name '2006 leave open now
    
Exit Sub

err_unitdes:
    Call General_Error_Trap
    Exit Sub
End Sub
