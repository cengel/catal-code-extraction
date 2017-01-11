Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo err_close
    
    DoCmd.Close acForm, Me.Name

Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errDel

Dim Response
Response = MsgBox("Do you really want to delete this inclusion group for Unit: " & Me!Unit & " , Ware code: " & Me![WARE CODE] & ", Surface Treatment: " & Me!SurfaceTreatment & "?", vbCritical + vbYesNo, "Confirm Deletion")
If Response = vbYes Then
    Dim sql
    sql = "Delete from Ceramics_Body_Sherd_inclusionsdetermined where inclusion_group_id = " & Me!InclusionGroupID & ";"
    DoCmd.RunSQL sql
    
    sql = "Delete from Ceramics_Body_Sherd_inclusion_group where inclusiongroupid = " & Me!InclusionGroupID & ";"
    DoCmd.RunSQL sql
    
    MsgBox "Deletion successful"
    DoCmd.Close acForm, Me.Name
    
End If

Exit Sub

errDel:
    Call General_Error_Trap
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open

    If Not IsNull(Me.OpenArgs) Then
        'this means a new inclusion group so must get unit and warecode from openargs
        Dim args, getUnit, getWarecode, getsurfaceT
        Dim firstamp, secondamp
        args = Me.OpenArgs
        firstamp = InStr(args, "&")
        secondamp = InStr(firstamp + 1, args, "&")
        
        getUnit = Left(args, InStr(args, "&") - 1)
        getWarecode = Mid(args, firstamp + 1, (secondamp - 1) - firstamp)
        getsurfaceT = Right(args, Len(args) - secondamp)
        
        DoCmd.RunCommand acCmdRecordsGoToNew
        
        Me![Unit].Locked = False
        Me![WARE CODE].Locked = False
        Me![SurfaceTreatment].Locked = False
        
        Me![Unit] = getUnit
        Me![WARE CODE] = getWarecode
        Me![SurfaceTreatment] = getsurfaceT
        
        Me![Unit].Locked = True
        Me![WARE CODE].Locked = True
        Me![SurfaceTreatment].Locked = True
    End If

Exit Sub

err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
