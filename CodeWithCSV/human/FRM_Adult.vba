Option Compare Database
Option Explicit

Function LoadXmlFile(Path As String) As MSXML2.DOMDocument60
    LoadXmlFile = New MSXML2.DOMDocument60

    With LoadXmlFile
        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .Load (Path)
    End With
End Function

Private Sub var(p1 As Object)
    Throw New NotImplementedException
End Sub


    
Public Sub DisplayNode(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
ByVal Indent As Integer)

Dim xNode As MSXML2.IXMLDOMNode

Indent = Indent + 1

For Each xNode In Nodes

    If xNode.nodeType = MSXML2.DOMNodeType.NODE_ELEMENT Then
        Log = Log & (Space$(Indent) & xNode.parentNode.nodeName & _
                        ":" & xNode.nodeValue)
    End If

    If xNode.hasChildNodes Then
        'DisplayNode(xNode.childNodes, Indent)
    End If
Next xNode

End Sub



Private Sub cboFind_AfterUpdate()
'find skeleton record - SAJ
On Error GoTo err_cboFind

    If Me![cboFind] <> "" Then
        Me.Filter = "[UnitNumber] = " & Me![cboFind] & " AND [Individual Number] = " & Me!cboFind.Column(1)
        Me.FilterOn = True
    End If
Exit Sub

err_cboFind:
    MsgBox Err.Description
    Exit Sub
End Sub


Private Sub cmdAll_Click()
'take off any filter - saj
On Error GoTo err_all

    Me.FilterOn = False
    Me.Filter = ""
Exit Sub

err_all:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub CmdOpenPermTeethFrm_Click()
On Error GoTo Err_CmdOpenPermTeethFrm_Click

    Call DoRecordCheck("HR_Teeth development measurement", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth development score", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Teeth wear", Me![txtUnit], Me![txtIndivid], "UnitNumber")
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    ''stDocName = "FRM_simons PERMANENT TEETH"
    stDocName = "FRM_Permanent_Teeth"
    
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenPermTeethFrm_Click:
    Exit Sub

Err_CmdOpenPermTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenPermTeethFrm_Click
    
End Sub
Private Sub CmdOpenMeasFrm_Click()
On Error GoTo Err_CmdOpenMeasFrm_Click

    Call DoRecordCheck("HR_Measurements version 2", Me![txtUnit], Me![txtIndivid], "UnitNumber")

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Measurement form version 2"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenMeasFrm_Click:
    Exit Sub

Err_CmdOpenMeasFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMeasFrm_Click
    
End Sub
Private Sub CmdOpenUnitDescFrm_Click()
On Error GoTo Err_CmdOpenUnitDescFrm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
    
Exit_CmdOpenUnitDescFrm_Click:
    Exit Sub

Err_CmdOpenUnitDescFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenUnitDescFrm_Click
    
End Sub
Private Sub CmdOpenMainMenuFrm_Click()
Call ReturnToMenu(Me)
    
End Sub
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click

    Call DoRecordCheck("HR_ageing and sexing", Me![txtUnit], Me![txtIndivid], "Unit Number")

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , "[Unit Number] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub

Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'saj new season 2007 - form jumping to middle on open - make sure it keeps the top at the top
On Error GoTo err_open

    DoCmd.GoToControl "FRM_SUBFORM_Adult_skull"
    'docmd.GoToControl ""

Exit Sub

err_open:
    General_Error_Trap
    Exit Sub
End Sub

Private Sub openSkeletonView_Click()
'Handles Button1.Click
On Error GoTo err_openSkeletonView

Dim xmldom As MSXML2.DOMDocument60
Set xmldom = New DOMDocument60
Dim xml_nodes_collection As MSXML2.IXMLDOMSelection
Dim xml_node_attribute As MSXML2.IXMLDOMAttribute
Dim root, unittag, node As MSXML2.IXMLDOMElement
Dim unreferenced As String
Dim referenced As String
Dim selectedbone, complete_colour, fragment_colour As String
Dim wrappingsvg As String
Dim file
Dim filebuffer, fileReader As String
Dim replacecheck
Dim fso
Dim Unit
Dim bones As DAO.Recordset
Dim bonesfiltered As DAO.Recordset
Dim mydb

Dim Path

Path = sketchpath & "units\skeletons\" & "S" & Me![txtUnit] & "_" & Me![txtIndivid] & ".jpg"
   
If Dir(Path) = "" Then
    'directory not exist
    MsgBox "The skeleton of this unit has not been modelled in yet. It will be created now, but the process will take some seconds. Try to open the skeleton view again in ca. 10 seconds.", vbInformation, "No Sketch available to view"

Set fso = CreateObject("Scripting.FileSystemObject")
Set root = xmldom.documentElement
Set unittag = xmldom.documentElement

complete_colour = "#666666"
fragment_colour = "#aaaaaa"

xmldom.async = False
xmldom.SetProperty "ProhibitDTD", False

xmldom.SetProperty "SelectionNamespaces", "xmlns='http://www.w3.org/2000/svg'"
xmldom.resolveExternals = False
xmldom.validateOnParse = False
xmldom.SetProperty "SelectionLanguage", "XPath"

xmldom.Load (sketchpath & "\units\skeletons\prototype\skeleton_prototype.svg")

Set unittag = xmldom.selectSingleNode("//*[local-name()='tspan' and @id = 'unit']")
unittag.Text = "S" & Me![txtUnit] & ".B" & Me![txtIndivid]

Set mydb = CurrentDb
Set bones = mydb.OpenRecordset("HR_Adult_Cranial_Data", dbOpenSnapshot)

bones.Filter = "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

Set bonesfiltered = bones.OpenRecordset

If Not (bonesfiltered.BOF And bonesfiltered.EOF) Then
    bonesfiltered.MoveFirst
    Do Until bonesfiltered.EOF 'Or counter = 1000
        
        If Not IsNull(bonesfiltered![Occipital_left_lateral]) Or Not IsNull(bonesfiltered![Frontal]) Or Not IsNull(bonesfiltered![Occipital_right_lateral]) Or Not IsNull(bonesfiltered![Occipital_squamous]) Or Not IsNull(bonesfiltered![Occipital_basilar]) Or Not IsNull(bonesfiltered![Parietal_left]) Or Not IsNull(bonesfiltered![Parietal_right]) Or Not IsNull(bonesfiltered![Temporal_left_petrous]) Or Not IsNull(bonesfiltered![Temporal_right_petrous]) Or Not IsNull(bonesfiltered![Temporal_left_squamous]) Or Not IsNull(bonesfiltered![Temporal_right_squamous]) Or Not IsNull(bonesfiltered![Sphenoid_body]) Or Not IsNull(bonesfiltered![Sphenoid_left_wing]) Or Not IsNull(bonesfiltered![Sphenoid_right_wing]) Then
        selectedbone = "frontalbone"
            'Set xml_nodes_collection = xmldom.selectNodes("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            selectedbone = "frontalbone"
            If bonesfiltered![Occipital_left_lateral] = 1 Or bonesfiltered![Frontal] = 1 Or bonesfiltered![Occipital_right_lateral] = 1 Or bonesfiltered![Occipital_squamous] = 1 Or bonesfiltered![Occipital_basilar] = 1 Or bonesfiltered![Parietal_left] = 1 Or bonesfiltered![Parietal_right] = 1 Or bonesfiltered![Temporal_left_petrous] = 1 Or bonesfiltered![Temporal_right_petrous] = 1 Or bonesfiltered![Temporal_left_squamous] = 1 Or bonesfiltered![Temporal_right_squamous] = 1 Or bonesfiltered![Sphenoid_body] = 1 Or bonesfiltered![Sphenoid_left_wing] = 1 Or bonesfiltered![Sphenoid_right_wing] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Zygomatic_left]) Or Not IsNull(bonesfiltered![Zygomatic_right]) Or Not IsNull(bonesfiltered![Nasal_left]) Or Not IsNull(bonesfiltered![Nasal_right]) Or Not IsNull(bonesfiltered![Palatine_left]) Or Not IsNull(bonesfiltered![Palatine_right]) Or Not IsNull(bonesfiltered![Small_face_bones]) Or Not IsNull(bonesfiltered![Maxilla_left]) Or Not IsNull(bonesfiltered![Maxilla_right]) Then
        selectedbone = "facialbone"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Zygomatic_left] = 1 Or bonesfiltered![Zygomatic_right] = 1 Or bonesfiltered![Nasal_left] = 1 Or bonesfiltered![Nasal_right] = 1 Or bonesfiltered![Palatine_left] = 1 Or bonesfiltered![Palatine_right] = 1 Or bonesfiltered![Small_face_bones] = 1 Or bonesfiltered![Maxilla_left] = 1 Or bonesfiltered![Maxilla_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
         
        If Not IsNull(bonesfiltered![Mandible_right]) Or Not IsNull(bonesfiltered![Mandible_left]) Or bonesfiltered![Hyoid] <> 0 Or bonesfiltered![Thyroid/cricoid_cartilage] <> 0 Then
        selectedbone = "mandible"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Mandible_right] = 1 Or bonesfiltered![Mandible_left] = 1 Or bonesfiltered![Hyoid] = 0 Or bonesfiltered![Thyroid/cricoid_cartilage] = 0 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        bonesfiltered.MoveNext
    Loop
Else
End If

Set bones = Nothing

Set mydb = CurrentDb
Set bones = mydb.OpenRecordset("HR_Adult_shoulder_hip", dbOpenSnapshot)

bones.Filter = "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

Set bonesfiltered = bones.OpenRecordset

If Not (bonesfiltered.BOF And bonesfiltered.EOF) Then
    bonesfiltered.MoveFirst
    Do Until bonesfiltered.EOF 'Or counter = 1000
        
        If Not IsNull(bonesfiltered![Clavicle_left_proximal]) Or Not IsNull(bonesfiltered![Clavicle_left_shaft]) Or Not IsNull(bonesfiltered![Clavicle_left_distal]) Then
        selectedbone = "clavicle_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Clavicle_left_proximal] = 1 Or bonesfiltered![Clavicle_left_shaft] = 1 Or bonesfiltered![Clavicle_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Clavicle_right_proximal]) Or Not IsNull(bonesfiltered![Clavicle_right_shaft]) Or Not IsNull(bonesfiltered![Clavicle_right_distal]) Then
        selectedbone = "clavicle_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
           If Not bonesfiltered![Clavicle_right_proximal] = 1 Or bonesfiltered![Clavicle_right_shaft] = 1 Or bonesfiltered![Clavicle_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Scapula_left]) Then
        selectedbone = "scapula_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        selectedbone = "acromionprocess_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        selectedbone = "coracoidprocess_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        If Not IsNull(bonesfiltered![Scapula_right]) Then
        selectedbone = "scapula_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        selectedbone = "acromionprocess_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        selectedbone = "coracoidprocess_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scapula_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Hipbone_ilium_left]) Or Not IsNull(bonesfiltered![Hipbone_ischium_left]) Or Not IsNull(bonesfiltered![Hipbone_pubis_left]) Then
        selectedbone = "ilium_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_left] = 1 Or bonesfiltered![Hipbone_ischium_left] = 1 Or bonesfiltered![Hipbone_pubis_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "iliaccrest_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_left] = 1 Or bonesfiltered![Hipbone_ischium_left] = 1 Or bonesfiltered![Hipbone_pubis_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "acetabulum_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_left] = 1 Or bonesfiltered![Hipbone_ischium_left] = 1 Or bonesfiltered![Hipbone_pubis_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='polygon']")
            If bonesfiltered![Hipbone_ilium_left] = 1 Or bonesfiltered![Hipbone_ischium_left] = 1 Or bonesfiltered![Hipbone_pubis_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Hipbone_ilium_right]) Or Not IsNull(bonesfiltered![Hipbone_ischium_right]) Or Not IsNull(bonesfiltered![Hipbone_pubis_right]) Then
        selectedbone = "ilium_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_right] = 1 Or bonesfiltered![Hipbone_ischium_right] = 1 Or bonesfiltered![Hipbone_pubis_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        selectedbone = "iliaccrest_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_right] = 1 Or bonesfiltered![Hipbone_ischium_right] = 1 Or bonesfiltered![Hipbone_pubis_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        selectedbone = "acetabulum_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hipbone_ilium_right] = 1 Or bonesfiltered![Hipbone_ischium_right] = 1 Or bonesfiltered![Hipbone_pubis_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='polygon']")
            If bonesfiltered![Hipbone_ilium_right] = 1 Or bonesfiltered![Hipbone_ischium_right] = 1 Or bonesfiltered![Hipbone_pubis_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        bonesfiltered.MoveNext
    Loop
Else
End If

Set bones = Nothing

Set mydb = CurrentDb
Set bones = mydb.OpenRecordset("HR_Adult_Arm_Data", dbOpenSnapshot)

bones.Filter = "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

Set bonesfiltered = bones.OpenRecordset

If Not (bonesfiltered.BOF And bonesfiltered.EOF) Then
    bonesfiltered.MoveFirst
    Do Until bonesfiltered.EOF 'Or counter = 1000
        
        If Not IsNull(bonesfiltered![Humerus_left_proximal]) Or Not IsNull(bonesfiltered![Humerus_left_shaft]) Or Not IsNull(bonesfiltered![Humerus_left_distal]) Then
        selectedbone = "humerus_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Humerus_left_proximal] = 1 Or bonesfiltered![Humerus_left_shaft] = 1 Or bonesfiltered![Humerus_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Humerus_right_proximal]) Or Not IsNull(bonesfiltered![Humerus_right_shaft]) Or Not IsNull(bonesfiltered![Humerus_right_distal]) Then
        selectedbone = "humerus_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Humerus_right_proximal] = 1 Or bonesfiltered![Humerus_right_shaft] = 1 Or bonesfiltered![Humerus_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ulna_left_proximal]) Or Not IsNull(bonesfiltered![Ulna_left_shaft]) Or Not IsNull(bonesfiltered![Ulna_left_distal]) Then
        selectedbone = "ulna_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ulna_left_proximal] = 1 Or bonesfiltered![Ulna_left_shaft] = 1 Or bonesfiltered![Ulna_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ulna_right_proximal]) Or Not IsNull(bonesfiltered![Ulna_right_shaft]) Or Not IsNull(bonesfiltered![Ulna_right_distal]) Then
        selectedbone = "ulna_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ulna_right_proximal] = 1 Or bonesfiltered![Ulna_right_shaft] = 1 Or bonesfiltered![Ulna_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Radius_left_proximal]) Or Not IsNull(bonesfiltered![Radius_left_shaft]) Or Not IsNull(bonesfiltered![Radius_left_distal]) Then
        selectedbone = "radius_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Radius_left_proximal] = 1 Or bonesfiltered![Radius_left_shaft] = 1 Or bonesfiltered![Radius_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Radius_right_proximal]) Or Not IsNull(bonesfiltered![Radius_right_shaft]) Or Not IsNull(bonesfiltered![Radius_right_distal]) Then
        selectedbone = "radius_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Radius_right_proximal] = 1 Or bonesfiltered![Radius_right_shaft] = 1 Or bonesfiltered![Radius_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Scaphoid_left] <> 0 Then
        selectedbone = "scaphoid_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scaphoid_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Scaphoid_right] <> 0 Then
        selectedbone = "scaphoid_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Scaphoid_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Lunate_left] <> 0 Then
        selectedbone = "lunate_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Lunate_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Lunate_right] <> 0 Then
        selectedbone = "lunate_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Lunate_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Triquetral_left] <> 0 Then
        selectedbone = "triquetrum_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Triquetral_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Triquetral_right] <> 0 Then
        selectedbone = "triquetrum_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Triquetral_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Pisiform_left] <> 0 Then
        selectedbone = "pisiform_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Pisiform_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Pisiform_right] <> 0 Then
        selectedbone = "pisiform_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Pisiform_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Trapezium_left] <> 0 Then
        selectedbone = "trapezoid_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Trapezium_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Trapezoid_right] <> 0 Then
        selectedbone = "trapezoid_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Trapezoid_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Capitate_left] <> 0 Then
        selectedbone = "capitate_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Capitate_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Capitate_right] <> 0 Then
        selectedbone = "capitate_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Capitate_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Trapezoid_left] <> 0 Then
        selectedbone = "trapezium_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Trapezoid_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Trapezium_right] <> 0 Then
        selectedbone = "trapezium_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Trapezium_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Hamate_left] <> 0 Then
        selectedbone = "hamate_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hamate_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Hamate_right] <> 0 Then
        selectedbone = "hamate_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Hamate_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_1_left] <> 0 Then
        selectedbone = "metacarpal1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_1_right] <> 0 Then
        selectedbone = "metacarpal1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_2_left] <> 0 Then
        selectedbone = "metacarpal2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_2_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_2_right] <> 0 Then
        selectedbone = "metacarpal2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_2_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_3_left] <> 0 Then
        selectedbone = "metacarpal3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_3_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_3_right] <> 0 Then
        selectedbone = "metacarpal3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_3_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metacarpal_4_left] <> 0 Then
        selectedbone = "metacarpal4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_4_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_4_right] <> 0 Then
        selectedbone = "metacarpal4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_4_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_5_left] <> 0 Then
        selectedbone = "metacarpal5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_5_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metacarpal_5_right] <> 0 Then
        selectedbone = "metacarpal5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metacarpal_5_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Proximal_phalanx_1_left] <> 0 Then
        selectedbone = "prox.phalanx1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanx_1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Proximal_phalanx_1_right] <> 0 Then
        selectedbone = "prox.phalanx1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Proximal_phalanges_2-5_left]) Then
        selectedbone = "prox.phalanx2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalanx3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        selectedbone = "prox.phalanx4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalanx5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        End If

        If Not IsNull(bonesfiltered![Proximal_phalanges_2-5_right]) Then
        selectedbone = "prox.phalanx2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalanx3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        selectedbone = "prox.phalanx4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalanx5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        End If


        If bonesfiltered![Distal_phalanx_1_left] <> 0 Then
        selectedbone = "dist.phalanx1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Distal_phalanx_1_right] <> 0 Then
        selectedbone = "dist.phalanx1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If


        If Not IsNull(bonesfiltered![Distal_phalanges_2-5_left]) Then
        selectedbone = "dist.phalanx2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Distal_phalanges_2-5_right]) Then
        selectedbone = "dist.phalanx2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalanx5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If


        If Not IsNull(bonesfiltered![Middle_phalanges_2-5_left]) Then
        selectedbone = "mid.phalanx2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Middle_phalanges_2-5_right]) Then
        selectedbone = "mid.phalanx2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "mid.phalanx5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        bonesfiltered.MoveNext
    Loop
Else
End If


Set bones = Nothing

Set mydb = CurrentDb
Set bones = mydb.OpenRecordset("HR_Adult_Axial_Data", dbOpenSnapshot)

bones.Filter = "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

Set bonesfiltered = bones.OpenRecordset

If Not (bonesfiltered.BOF And bonesfiltered.EOF) Then
    bonesfiltered.MoveFirst
    Do Until bonesfiltered.EOF 'Or counter = 1000
        
        If Not IsNull(bonesfiltered![Sternum_manubrium]) Or Not IsNull(bonesfiltered![Sternum_body]) Then
        selectedbone = "sternum"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Sternum_manubrium] = 1 Or bonesfiltered![Sternum_body] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ribs_left_1]) Then
        selectedbone = "rib1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_1] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ribs_left_2]) Then
        selectedbone = "rib2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_2] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_3]) Then
        selectedbone = "rib3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_3] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
                
        If Not IsNull(bonesfiltered![Ribs_left_4]) Then
        selectedbone = "rib4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_4] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_5]) Then
        selectedbone = "rib5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_5] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_6]) Then
        selectedbone = "rib6b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_6] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_7]) Then
        selectedbone = "rib7b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_7] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_8]) Then
        selectedbone = "rib8b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_8] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_9]) Then
        selectedbone = "rib9b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_9] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_10]) Then
        selectedbone = "rib10b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
        'node.setAttribute "fill", complete_colour
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_11]) Then
        selectedbone = "rib11b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_11] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_left_12]) Then
        selectedbone = "rib12b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_left_12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ribs_right_1]) Then
        selectedbone = "rib1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_1] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Ribs_right_2]) Then
        selectedbone = "rib2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_2] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_3]) Then
        selectedbone = "rib3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_3] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
                
        If Not IsNull(bonesfiltered![Ribs_right_4]) Then
        selectedbone = "rib4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_4] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_5]) Then
        selectedbone = "rib5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_5] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_6]) Then
        selectedbone = "rib6a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_6] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_7]) Then
        selectedbone = "rib7a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_7] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_8]) Then
        selectedbone = "rib8a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_8] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_9]) Then
        selectedbone = "rib9a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_9] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_10]) Then
        selectedbone = "rib10a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
        'node.setAttribute "fill", complete_colour
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_11]) Then
        selectedbone = "rib11a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_11] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Ribs_right_12]) Then
        selectedbone = "rib12a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Ribs_right_12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![C6]) Then
        selectedbone = "vertebra4"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![C6] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![C7]) Then
        selectedbone = "vertebra5"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![C7] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![T1]) Or Not IsNull(bonesfiltered![T2]) Or Not IsNull(bonesfiltered![T3]) Or Not IsNull(bonesfiltered![T4]) Or Not IsNull(bonesfiltered![T5]) Or Not IsNull(bonesfiltered![T6]) Or Not IsNull(bonesfiltered![T7]) Or Not IsNull(bonesfiltered![T8]) Or Not IsNull(bonesfiltered![T9]) Or Not IsNull(bonesfiltered![T10]) Or Not IsNull(bonesfiltered![T11]) Or Not IsNull(bonesfiltered![T12]) Then
        selectedbone = "vertebra6"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![T1] = 1 Or bonesfiltered![T2] = 1 Or bonesfiltered![T3] = 1 Or bonesfiltered![T4] = 1 Or bonesfiltered![T5] = 1 Or bonesfiltered![T6] = 1 Or bonesfiltered![T7] = 1 Or bonesfiltered![T8] = 1 Or bonesfiltered![T9] = 1 Or bonesfiltered![T10] = 1 Or bonesfiltered![T11] = 1 Or bonesfiltered![T12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If

        selectedbone = "vertebra7"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![T1] = 1 Or bonesfiltered![T2] = 1 Or bonesfiltered![T3] = 1 Or bonesfiltered![T4] = 1 Or bonesfiltered![T5] = 1 Or bonesfiltered![T6] = 1 Or bonesfiltered![T7] = 1 Or bonesfiltered![T8] = 1 Or bonesfiltered![T9] = 1 Or bonesfiltered![T10] = 1 Or bonesfiltered![T11] = 1 Or bonesfiltered![T12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "vertebra18"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![T1] = 1 Or bonesfiltered![T2] = 1 Or bonesfiltered![T3] = 1 Or bonesfiltered![T4] = 1 Or bonesfiltered![T5] = 1 Or bonesfiltered![T6] = 1 Or bonesfiltered![T7] = 1 Or bonesfiltered![T8] = 1 Or bonesfiltered![T9] = 1 Or bonesfiltered![T10] = 1 Or bonesfiltered![T11] = 1 Or bonesfiltered![T12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "vertebra19"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![T1] = 1 Or bonesfiltered![T2] = 1 Or bonesfiltered![T3] = 1 Or bonesfiltered![T4] = 1 Or bonesfiltered![T5] = 1 Or bonesfiltered![T6] = 1 Or bonesfiltered![T7] = 1 Or bonesfiltered![T8] = 1 Or bonesfiltered![T9] = 1 Or bonesfiltered![T10] = 1 Or bonesfiltered![T11] = 1 Or bonesfiltered![T12] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![L1]) Then
        selectedbone = "vertebra20"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![L1] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![L2]) Then
        selectedbone = "vertebra21"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![L2] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![L3]) Then
        selectedbone = "vertebra22"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![L3] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![L4]) Then
        selectedbone = "vertebra23"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![L4] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![L5]) Then
        selectedbone = "vertebra24"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![L5] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Sacrum]) Then
        selectedbone = "sacrum"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Sacrum] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        bonesfiltered.MoveNext
    Loop
Else
End If

Set bones = Nothing

Set mydb = CurrentDb
Set bones = mydb.OpenRecordset("HR_Adult_Leg_Data", dbOpenSnapshot)

bones.Filter = "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]

Set bonesfiltered = bones.OpenRecordset

If Not (bonesfiltered.BOF And bonesfiltered.EOF) Then
    bonesfiltered.MoveFirst
    Do Until bonesfiltered.EOF 'Or counter = 1000
        
        If Not IsNull(bonesfiltered![Femur_left_proximal]) Or Not IsNull(bonesfiltered![Femur_left_shaft]) Or Not IsNull(bonesfiltered![Femur_left_distal]) Then
        selectedbone = "femur_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Femur_left_proximal] = 1 Or bonesfiltered![Femur_left_shaft] = 1 Or bonesfiltered![Femur_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Femur_right_proximal]) Or Not IsNull(bonesfiltered![Femur_right_shaft]) Or Not IsNull(bonesfiltered![Femur_right_distal]) Then
        selectedbone = "femur_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Femur_right_proximal] = 1 Or bonesfiltered![Femur_right_shaft] = 1 Or bonesfiltered![Femur_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Tibia_left_proximal]) Or Not IsNull(bonesfiltered![Tibia_left_shaft]) Or Not IsNull(bonesfiltered![Tibia_left_distal]) Then
        selectedbone = "tibia_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Tibia_left_proximal] = 1 Or bonesfiltered![Tibia_left_shaft] = 1 Or bonesfiltered![Tibia_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Tibia_right_proximal]) Or Not IsNull(bonesfiltered![Tibia_right_shaft]) Or Not IsNull(bonesfiltered![Tibia_right_distal]) Then
        selectedbone = "tibia_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Tibia_right_proximal] = 1 Or bonesfiltered![Tibia_right_shaft] = 1 Or bonesfiltered![Tibia_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Fibula_left_proximal]) Or Not IsNull(bonesfiltered![Fibula_left_shaft]) Or Not IsNull(bonesfiltered![Fibula_left_distal]) Then
        selectedbone = "fibula_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Fibula_left_proximal] = 1 Or bonesfiltered![Fibula_left_shaft] = 1 Or bonesfiltered![Fibula_left_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Fibula_right_proximal]) Or Not IsNull(bonesfiltered![Fibula_right_shaft]) Or Not IsNull(bonesfiltered![Fibula_right_distal]) Then
        selectedbone = "fibula_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Fibula_right_proximal] = 1 Or bonesfiltered![Fibula_right_shaft] = 1 Or bonesfiltered![Fibula_right_distal] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Patella_left]) Then
        selectedbone = "patella_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Patella_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Patella_right]) Then
        selectedbone = "patella_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Patella_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Calcaneus_right] <> 0 Then
        selectedbone = "calcaneus_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Calcaneus_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Calcaneus_left] <> 0 Then
        selectedbone = "calcaneus_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Calcaneus_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Talus_right] <> 0 Then
        selectedbone = "talus_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Talus_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Talus_left] <> 0 Then
        selectedbone = "talus_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Talus_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Navicular_right] <> 0 Then
        selectedbone = "navicular_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Navicular_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Navicular_left] <> 0 Then
        selectedbone = "navicular_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Navicular_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Cuboid_right] <> 0 Then
        selectedbone = "cuboid_a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuboid_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Cuboid_left] <> 0 Then
        selectedbone = "cuboid_b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuboid_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Cuneiform1_right] <> 0 Then
        selectedbone = "cuneiform1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Cuneiform1_left] <> 0 Then
        selectedbone = "cuneiform1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Cuneiform2_right] <> 0 Then
        selectedbone = "cuneiform2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform2_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Cuneiform2_left] <> 0 Then
        selectedbone = "cuneiform2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform2_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Cuneiform3_right] <> 0 Then
        selectedbone = "cuneiform3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform3_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Cuneiform3_left] <> 0 Then
        selectedbone = "cuneiform3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Cuneiform3_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metatarsal_1_right] <> 0 Then
        selectedbone = "metatarsal1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_1_left] <> 0 Then
        selectedbone = "metatarsal1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metatarsal_2_right] <> 0 Then
        selectedbone = "metatarsal2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_2_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_2_left] <> 0 Then
        selectedbone = "metatarsal2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_2_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_3_right] <> 0 Then
        selectedbone = "metatarsal3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_3_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_3_left] <> 0 Then
        selectedbone = "metatarsal3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_3_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metatarsal_4_right] <> 0 Then
        selectedbone = "metatarsal4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_4_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_4_left] <> 0 Then
        selectedbone = "metatarsal4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_4_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Metatarsal_5_right] <> 0 Then
        selectedbone = "metatarsal5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_5_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Metatarsal_5_left] <> 0 Then
        selectedbone = "metatarsal5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Metatarsal_5_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Proximal_phalanx_1_right] <> 0 Then
        selectedbone = "prox.phalange1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Proximal_phalanx_1_left] <> 0 Then
        selectedbone = "prox.phalange1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanx_1_left] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Middle_phalanges_2-5_right]) Then
        selectedbone = "dist.phalange2.2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange3.2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange4.2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange5.2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Middle_phalanges_2-5_left]) Then
        selectedbone = "dist.phalange2.2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange3.2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange4.2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange5.2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Middle_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If bonesfiltered![Distal_phalanx_1_right] <> 0 Then
        selectedbone = "dist.phalange1.1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange1.2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If bonesfiltered![Distal_phalanx_1_left] <> 0 Then
        selectedbone = "dist.phalange1.1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange1.2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanx_1_right] = 1 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If

        If Not IsNull(bonesfiltered![Distal_phalanges_2-5_right]) Then
        selectedbone = "dist.phalange2.1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange3.1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange4.1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange5.1a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Distal_phalanges_2-5_left]) Then
        selectedbone = "dist.phalange2.1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange3.1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange4.1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "dist.phalange5.1b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Distal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Proximal_phalanges_2-5_right]) Then
        selectedbone = "prox.phalange2b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange3b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange4b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange5b"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_right] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        If Not IsNull(bonesfiltered![Proximal_phalanges_2-5_left]) Then
        selectedbone = "prox.phalange2a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange3a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange4a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        
        selectedbone = "prox.phalange5a"
        Set node = xmldom.selectSingleNode("//*[local-name()='g' and @id = '" & selectedbone & "']/*[local-name()='path']")
            If bonesfiltered![Proximal_phalanges_2-5_left] < 4 Then
                node.setAttribute "fill", fragment_colour
            Else
                node.setAttribute "fill", complete_colour
            End If
        End If
        
        bonesfiltered.MoveNext
    Loop
Else
End If



xmldom.Save (sketchpath & "\units\skeletons\S" & Me![txtUnit] & "_" & Me![txtIndivid] & ".svg")

'ugly hack to reference schemas necessary for svg display [implementing schemas in MSXML does not work correctly]

unreferenced = Chr(34) & "Layer_1" & Chr(34) & " width=" & Chr(34) & "277.779px" & Chr(34)

referenced = Chr(34) & "Layer_1" & Chr(34) & " xmlns=" & Chr(34) & "http://www.w3.org/2000/svg" & Chr(34) & " xmlns:xlink=" & Chr(34) & "http://www.w3.org/1999/xlink" & Chr(34) & " width=" & Chr(34) & "277.779px" & Chr(34)
    'Set filebuffer = My.Computer.FileSystem.ReadAllText(CurrentProject.Path & "\skeleton_modified.svg")

    If fso.FileExists(sketchpath & "\units\skeletons\S" & Me![txtUnit] & "_" & Me![txtIndivid] & ".svg") Then
    Set file = fso.OpenTextFile(sketchpath & "\units\skeletons\S" & Me![txtUnit] & "_" & Me![txtIndivid] & ".svg", 1)

    While Not file.AtEndOfStream
        filebuffer = filebuffer & file.ReadLine & Chr(13) & Chr(10)
    Wend
    file.Close

    If InStr(filebuffer, unreferenced) > 0 Then
        fileReader = Replace(filebuffer, unreferenced, referenced)
        Set file = fso.OpenTextFile(sketchpath & "\units\skeletons\S" & Me![txtUnit] & "_" & Me![txtIndivid] & ".svg", 2)
        file.WriteLine fileReader
        file.Close
    Else

    End If
    'WebBrowser1.Navigate (Application.StartupPath & "/wrapsvg.html")
    'WebBrowser1.DocumentText = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & "?><!DOCTYPE html PUBLIC " & Chr(34) & "-//W3C//DTD XHTML 1.0 Transitional//EN" & Chr(34) & " " & Chr(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & Chr(34) & "><html xmlns=" & Chr(34) & "http://www.w3.org/1999/xhtml" & Chr(34) & "><head><!-- <meta content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & " http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " /> --><title>Skeleton</title></head><body>" & fileReader & "</body></html>"
    'Me![picturecontainer].Picture = CurrentProject.Path & "\skeleton_modified.svg"
Else
    
End If

Else
    DoCmd.OpenForm "frm_pop_graphic", acNormal, , , acFormReadOnly, , Me![txtUnit] & "_" & Me![txtIndivid]
End If
    
Exit Sub

err_openSkeletonView:
    MsgBox Err.Description
    Exit Sub


End Sub





