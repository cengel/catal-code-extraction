Option Compare Database
Option Explicit

Private Sub Artefact_Class_AfterUpdate()
'The artefact class determines which subform appears, and if the user
'changes the class any existing data related to the old class must be removed.
'if Tool - show subform Frm_subform_Tools
'if Core - show subform Frm_subform_Cores/Debitage
'if uniden - show no subform
'SAJ
On Error GoTo err_Artefact

Dim retVal, sql

If Me![Artefact Class] <> "" And Me![ArtefactClassSubform].SourceObject <> "" Then
    
    If Not IsNull(Me![ArtefactClassSubform]![GID]) Then
        'data for old class exists
        
        If Me![Artefact Class].OldValue = "Core/Debitage" Then
            'was core
            retVal = MsgBox("If you change the artefact class you will lose all the Core specific data, are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![Artefact Class] = Me![Artefact Class].OldValue
                Exit Sub
            Else
                sql = "DELETE FROM [GroundStone 5: Cores/Debitage] WHERE [GID]='" & Me![GID] & "';"
                DoCmd.RunSQL sql
            End If
        ElseIf Me![Artefact Class].OldValue = "Tool" Then
            'was tools
            retVal = MsgBox("If you change the artefact class you will lose all the Tool specific data, are you sure?", vbQuestion + vbYesNo, "Confirm Action")
            If retVal = vbNo Then
                Me![Artefact Class] = Me![Artefact Class].OldValue
                Exit Sub
            Else
                sql = "DELETE FROM [GroundStone 4: Tools] WHERE [GID]='" & Me![GID] & "';"
                DoCmd.RunSQL sql
            End If
        End If
    End If
End If

If Me![Artefact Class] <> "" Then
    If Me![Artefact Class].Column(1) = 2 Then
        'user has selected TOOL
        Me![txtArtefactClassLBL].ControlSource = "='The artefact class is Tool'"
        Me![ArtefactClassSubform].SourceObject = "Frm_subform_Tools"
        Me![ArtefactClassSubform].Form![txtGID] = Me![txtGID]
        Me![ArtefactClassSubform].Height = "4620"
        Forms![Frm_GS_Main]![subfrmWorkedOrUnworked].Height = "9800"
    ElseIf Me![Artefact Class].Column(1) = 1 Then
        'user has selected CORE
        Me![txtArtefactClassLBL].ControlSource = "='The artefact class is Cores/Debitage'"
        Me![ArtefactClassSubform].SourceObject = "Frm_subform_Cores/Debitage"
        Me![ArtefactClassSubform].Form![txtGID] = Me![txtGID]
        Me![ArtefactClassSubform].Height = "4620"
        Forms![Frm_GS_Main]![subfrmWorkedOrUnworked].Height = "9800"
    Else
        'user selected unidentifiable
        Me![txtArtefactClassLBL].ControlSource = ""
        Me![ArtefactClassSubform].SourceObject = ""
        Me![ArtefactClassSubform].Height = "0"
        Forms![Frm_GS_Main]![subfrmWorkedOrUnworked].Height = "4900"
    End If
    
    'now update the artefact type and subtype lists
    Dim rowsrc
    Me![Artefact Type] = Null
    rowsrc = "SELECT [GroundStone List of Values: Artefact Type].[Tool Types], "
    rowsrc = rowsrc & "[GroundStone List of Values: Artefact Type].[Code], "
    rowsrc = rowsrc & "[GroundStone List of Values: Artefact Type].[ClassID], "
    rowsrc = rowsrc & "[GroundStone List of Values: Artefact Class].Class "
    rowsrc = rowsrc & "FROM [GroundStone List of Values: Artefact Type] LEFT JOIN [GroundStone List of Values: Artefact Class] ON [GroundStone List of Values: Artefact Type].ClassID = [GroundStone List of Values: Artefact Class].ClassID "
    rowsrc = rowsrc & "WHERE [GroundStone List of Values: Artefact Class].Class = '" & Me![Artefact Class] & "'"
    rowsrc = rowsrc & " ORDER BY [GroundStone List of Values: Artefact Type].[Tool Types];"
    Me![Artefact Type].RowSource = rowsrc
    
    Me![Artefact SubType] = Null
    rowsrc = "SELECT [Groundstone List of Values: Artefact SubType].[Tool Subtype], "
    rowsrc = rowsrc & "[Groundstone List of Values: Artefact SubType].Code, "
    rowsrc = rowsrc & "[Groundstone List of Values: Artefact SubType].TypeCode, "
    rowsrc = rowsrc & "[GroundStone List of Values: Artefact Type].[Tool Types]  "
    rowsrc = rowsrc & " FROM [Groundstone List of Values: Artefact SubType] LEFT JOIN [GroundStone List of Values: Artefact Type] ON [Groundstone List of Values: Artefact SubType].TypeCode = [GroundStone List of Values: Artefact Type].Code"
    rowsrc = rowsrc & " WHERE [Tool Types] = '" & Me![Artefact Type] & "'"
    rowsrc = rowsrc & " ORDER BY [Groundstone List of Values: Artefact SubType].[Tool Subtype];"
    Me![Artefact SubType].RowSource = rowsrc
End If
Exit Sub

err_Artefact:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub Artefact_Type_AfterUpdate()
'
On Error GoTo err_arteType

    'now update the artefact subtype list

    Dim rowsrc
    Me![Artefact SubType] = Null
    rowsrc = "SELECT [Groundstone List of Values: Artefact SubType].[Tool Subtype], "
    rowsrc = rowsrc & "[Groundstone List of Values: Artefact SubType].Code, "
    rowsrc = rowsrc & "[Groundstone List of Values: Artefact SubType].TypeCode, "
    rowsrc = rowsrc & "[GroundStone List of Values: Artefact Type].[Tool Types]  "
    rowsrc = rowsrc & " FROM [Groundstone List of Values: Artefact SubType] LEFT JOIN [GroundStone List of Values: Artefact Type] ON [Groundstone List of Values: Artefact SubType].TypeCode = [GroundStone List of Values: Artefact Type].Code"
    rowsrc = rowsrc & " WHERE [Tool Types] = '" & Me![Artefact Type] & "'"
    rowsrc = rowsrc & " ORDER BY [Groundstone List of Values: Artefact SubType].[Tool Subtype];"
    Me![Artefact SubType].RowSource = rowsrc

Exit Sub

err_arteType:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Current()

If Me![Artefact Class].Column(1) = 2 Then
    'Me![Form Groundstone: Tools].Visible = True
    Me![txtArtefactClassLBL].ControlSource = "='The artefact class is Tool'"
    Me![ArtefactClassSubform].SourceObject = "Frm_subform_Tools"
    Me![ArtefactClassSubform].Height = "4620"
ElseIf Me![Artefact Class].Column(1) = 1 Then
    'Me![Form Groundstone: Tools].Visible = False
    Me![txtArtefactClassLBL].ControlSource = "='The artefact class is Cores/Debitage'"
    Me![ArtefactClassSubform].SourceObject = "Frm_subform_Cores/Debitage"
    Me![ArtefactClassSubform].Height = "4620"
Else
    'unidentifiable
    Me![txtArtefactClassLBL].ControlSource = ""
    Me![ArtefactClassSubform].SourceObject = ""
    Me![ArtefactClassSubform].Height = 0
End If
    
End Sub

Private Sub Raw_Material_Group_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_Raw_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [Groundstone List of Values: RawMaterialGroup]([RawMaterialGroup]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_Raw_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Raw_Material_Texture_NotInList(NewData As String, Response As Integer)
'Allow more values to be added if necessary
On Error GoTo err_RawText_NotInList

Dim retVal, sql, inputname

retVal = MsgBox("This value is not in the list, are you sure you want to add it?", vbQuestion + vbYesNo, "Value not in list")
If retVal = vbYes Then
    Response = acDataErrAdded
    sql = "INSERT INTO [Groundstone List of Values: RawMaterialTexture]([RawMaterialTexture]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    Response = acDataErrContinue
End If

   
Exit Sub

err_RawText_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub
