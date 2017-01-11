Option Compare Database

Private Sub cmdDLowerDisease_Click()
On Error GoTo Err_cmdUnitDesc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Deciduous_Lower_Teeth_Disease"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit_cmdUnitDesc_Click:
    Exit Sub

Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDesc_Click
End Sub

Private Sub cmdDUpperDisease_Click()
On Error GoTo Err_cmdUnitDesc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Deciduous_Upper_Teeth_Disease"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit_cmdUnitDesc_Click:
    Exit Sub

Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDesc_Click
End Sub

Private Sub cmdMNI_Click()
'new 2 sept 09 - 2009 season
On Error GoTo Err_cmdMNI

    'DoCmd.OpenForm "FRM_MNI"
    DoCmd.OpenForm "Frm_Pop_Choose_MNI"

    Exit Sub

Err_cmdMNI:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdPLowerDisease_Click()
On Error GoTo Err_cmdUnitDesc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Permanent_Lower_Teeth_Disease_v2"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit_cmdUnitDesc_Click:
    Exit Sub

Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDesc_Click
End Sub

Private Sub cmdPUpperDisease_Click()
On Error GoTo Err_cmdUnitDesc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Permanent_Upper_Teeth_Disease_v7"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit_cmdUnitDesc_Click:
    Exit Sub

Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDesc_Click
End Sub

Private Sub cmdUnitDesc_Click()
On Error GoTo Err_cmdUnitDesc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_UnitDescription"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit_cmdUnitDesc_Click:
    Exit Sub

Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDesc_Click
    
End Sub

Private Sub cmdUnitDescription_Click()
On Error GoTo Err_cmdUnitDescription_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdUnitDescription_Click:
    Exit Sub

Err_cmdUnitDescription_Click:
    MsgBox Err.Description
    Resume Exit_cmdUnitDescription_Click
    
End Sub
Private Sub cmdQuit_Click()
On Error GoTo Err_cmdQuit_Click


    DoCmd.Quit

Exit_cmdQuit_Click:
    Exit Sub

Err_cmdQuit_Click:
    MsgBox Err.Description
    Resume Exit_cmdQuit_Click
    
End Sub
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub

Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
    
End Sub
Private Sub CmdOpenNeonateFrm_Click()
On Error GoTo Err_CmdOpenNeonateFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    'unchanged start season 2007
    stDocName = "FRM_simons NEONATAL FORM"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenNeonateFrm_Click:
    Exit Sub

Err_CmdOpenNeonateFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenNeonateFrm_Click
    
End Sub
Private Sub CmdOpenJuvenileFrm_Click()
On Error GoTo Err_CmdOpenJuvenileFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons juvenile form"
    ''season 2007, saj
    stDocName = "FRM_Juvenile"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenJuvenileFrm_Click:
    Exit Sub

Err_CmdOpenJuvenileFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenJuvenileFrm_Click
    
End Sub
Private Sub CmdOpenAdultFrm_Click()
On Error GoTo Err_CmdOpenAdultFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_Simons adult form"
    ''season 2007, saj
    stDocName = "FRM_Adult"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenAdultFrm_Click:
    Exit Sub

Err_CmdOpenAdultFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAdultFrm_Click
    
End Sub
Private Sub CmdOpenDecidTeethFrm_Click()
On Error GoTo Err_CmdOpenDecidTeethFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons DECIDUOUS TEETH"
    ''season 2007, saj
    stDocName = "FRM_Deciduous_Teeth"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenDecidTeethFrm_Click:
    Exit Sub

Err_CmdOpenDecidTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenDecidTeethFrm_Click
    
End Sub
Private Sub CmdOpenPermTeethFrm_Click()
On Error GoTo Err_CmdOpenPermTeethFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons PERMANENT TEETH"
    ''season 2007, saj
    stDocName = "FRM_Permanent_Teeth"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenPermTeethFrm_Click:
    Exit Sub

Err_CmdOpenPermTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenPermTeethFrm_Click
    
End Sub
Private Sub CmdOpenMeasFrm_Click()
On Error GoTo Err_CmdOpenMeasFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Measurement form version 2"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmdOpenMeasFrm_Click:
    Exit Sub

Err_CmdOpenMeasFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMeasFrm_Click
    
End Sub
