Option Compare Database
Option Explicit

Sub PostEx_BFD_BodyPortion(tblrun, Taxon, Element)
'this code will update the bodyportion field on the bfd dependant on values in
'taxon and element field.
'It can be run to do the whole table (tblrun = true) or for an individual record on entry
'which will carry in taxon and element
On Error GoTo err_postex1
Dim sql

    If tblrun = True Then
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Skull' WHERE (Taxon between 1 and 199) AND ((Element between 1 AND 24) OR (Element = 118));"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Axial' WHERE (Taxon between 1 and 199) AND (Element between 25 AND 35);"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Girdle' WHERE (Taxon between 1 and 199) AND ((Element between 36 AND 37) OR (Element between 66 AND 72));"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Upperlimb' WHERE (Taxon between 1 and 199) AND ((Element between 38 AND 40) OR (Element between 74 AND 78) OR (Element =115));"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Lowerlimb' WHERE (Taxon between 1 and 199) AND ((Element between 41 AND 64) OR (Element between 79 AND 113));"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Body Portion] = 'Unidentified' WHERE (Taxon between 1 and 199) AND (Element =116);"
        DoCmd.RunSQL sql

    Else
        'must have element and taxon to proceed
        If Taxon <> "" And Element <> "" Then
            If (Taxon >= 1 And Taxon <= 199) And ((Element = 118) Or (Element >= 1 And Element <= 24)) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Skull"

            ElseIf (Taxon >= 1 And Taxon <= 199) And (Element >= 25 And Element <= 35) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Axial"

            ElseIf (Taxon >= 1 And Taxon <= 199) And (Element = 36 Or Element = 37) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Girdle"

            ElseIf (Taxon >= 1 And Taxon <= 199) And ((Element >= 38 And Element <= 40) Or (Element >= 74 And Element <= 78) Or (Element = 115)) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Upperlimb"
         
            ElseIf (Taxon >= 1 And Taxon <= 199) And ((Element >= 41 And Element <= 64) Or (Element >= 79 And Element <= 113)) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Lowerlimb"

           

            ElseIf (Taxon >= 1 And Taxon <= 199) And (Element = 116) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Unidentified"
    
            ElseIf (Taxon >= 1 And Taxon <= 199) And (Element >= 66 And Element <= 72) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtBodyPortion] = "Girdle"

            End If
        End If
    End If
Exit Sub

err_postex1:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub PostEx_BFD_SizeClass(tblrun, Taxon)
'this code will update the size class field on the bfd dependant on values in
'taxon field.
'It can be run to do the whole table (tblrun = true) or for an individual record on entry
'which will carry in taxon and element
On Error GoTo err_postex2
Dim sql

    If tblrun = True Then
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Size Class] = '02 size' WHERE (Taxon = 2) OR (Taxon = 50) OR (Taxon = 61) OR (Taxon between 64 and 69) OR (Taxon between 75 and 85) OR (Taxon between 95 and 98) OR (Taxon = 106) OR (Taxon between 153 and 158 );"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Size Class] = '03 size' WHERE (Taxon = 3) OR (Taxon between 14 and 23) OR (Taxon = 30) OR (Taxon = 32) OR (Taxon = 51) OR (Taxon between 86 and 91);"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Size Class] = '05 size' WHERE (Taxon = 5) OR (Taxon = 34) OR (Taxon = 37) OR (Taxon = 40) OR (Taxon between 42 and 44)  OR (Taxon between 46 and 47)  OR (Taxon = 100);"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Basic_Faunal_Data] SET [Size Class] = '07 size' WHERE (Taxon between 7 and 8) OR (Taxon = 11) OR (Taxon = 13) OR (Taxon between 25 and 28) OR (Taxon = 31)  OR (Taxon = 33)  OR (Taxon = 41) OR (Taxon = 45);"
        DoCmd.RunSQL sql
    Else
    'must have taxon to proceed
        If Taxon <> "" Then
            'size class 2
            If Taxon = 2 Or Taxon = 50 Or Taxon = 61 Or (Taxon >= 64 And Taxon <= 69) Or (Taxon >= 75 And Taxon <= 85) Or (Taxon >= 95 And Taxon <= 98) Or Taxon = 106 Or (Taxon >= 153 And Taxon <= 158) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtSizeClass] = "02 size"
            'size class 3
            ElseIf Taxon = 3 Or (Taxon >= 14 And Taxon <= 23) Or (Taxon >= 14 And Taxon <= 23) Or Taxon = 30 Or Taxon = 32 Or Taxon = 51 Or (Taxon >= 86 And Taxon <= 91) Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtSizeClass] = "03 size"
            'size class 5
            ElseIf Taxon = 5 Or Taxon = 34 Or Taxon = 37 Or Taxon = 40 Or (Taxon >= 42 And Taxon <= 44) Or Taxon = 46 Or Taxon = 47 Or Taxon = 100 Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtSizeClass] = "05 size"
            'size class 7
            ElseIf Taxon = 7 Or Taxon = 8 Or Taxon = 11 Or Taxon = 13 Or (Taxon >= 25 And Taxon <= 28) Or Taxon = 31 Or Taxon = 33 Or Taxon = 41 Or Taxon = 45 Then
                Forms![Fauna_Bone_Basic_Faunal_Data]![txtSizeClass] = "07 size"

            End If
        End If
    End If
Exit Sub

err_postex2:
    Call General_Error_Trap
    Exit Sub
End Sub

Sub PostEx_PostCran_ElementPortion(tblrun, ProxDist)
'this code will update the element portion field on the post cran dependant on values in
'taxon field.
'It can be run to do the whole table (tblrun = true) or for an individual record on entry
'which will carry in proximal/distal value
On Error GoTo err_postex3
Dim sql

    If tblrun = True Then
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Complete' WHERE [Proximal/Distal] = 10;"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Proximal End' WHERE [Proximal/Distal] between 1 and 3;"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Shaft' WHERE [Proximal/Distal] between 4 and 6;"
        DoCmd.RunSQL sql
    
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Distal End' WHERE [Proximal/Distal] between 7 and 9;"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Vertebral Body' WHERE [Proximal/Distal] = 20;"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Vertebral Processes' WHERE [Proximal/Distal] between 21 and 22;"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Indeterminate' WHERE [Proximal/Distal] <1;"
        DoCmd.RunSQL sql
        
        sql = "UPDATE [Fauna_Bone_Postcranial] SET [Element Portion] = 'Body and Processes' WHERE [Proximal/Distal] between 23 and 24;"
        DoCmd.RunSQL sql
    Else
    'must have prox/distal to proceed
        If ProxDist <> "" Then
            'set element portion
            If ProxDist = 10 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Complete"
            ElseIf ProxDist >= 1 And ProxDist <= 3 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Proximal End"
            ElseIf ProxDist >= 4 And ProxDist <= 6 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Shaft"
            ElseIf ProxDist >= 7 And ProxDist <= 9 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Distal End"
            ElseIf ProxDist = 20 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Vertebral Body"
            ElseIf ProxDist >= 21 And ProxDist <= 22 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Vertebral Processes"
            ElseIf ProxDist < 1 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Indeterminate"
            ElseIf ProxDist >= 23 And ProxDist <= 24 Then
                Forms![Fauna_Bone_Postcranial]![txtElementPortion] = "Vertebral Processes"
            End If
        End If
    End If
Exit Sub

err_postex3:
    Call General_Error_Trap
    Exit Sub
End Sub

