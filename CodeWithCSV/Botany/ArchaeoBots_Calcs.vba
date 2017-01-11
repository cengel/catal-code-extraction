Option Compare Database
Option Explicit

Function Calc_WoodParenceDung_ml_per_litre(flot, fldVal, FractionVal)
'*************************************************************************************
'this function takes the Priority Sample table 4mm wood/parenc/dung field and the fraction
'field, gathers the basic data soil vol field and then calc the wood/parenc/dung per litre figure.
' Inputs:   Flot = Priority Sample flot number
'           fldVal = Priority Sample 4mm Wood or 4mm Parenc or 4mm dung field
'           FractionVal = Priority Sample 4mm Fraction Field
'
' Outputs: wood ml per litre or parenc ml per litre or dung ml per litre
' Called from: FRM_Priority - cmdAddtoPReport_Click
'              FRM_PriorityReport - cmdRecalc_Click
'
'SAJ season 2006, request from AB
'************************************************************************************
On Error GoTo err_catch

    'first check correct values passed in
    If (fldVal <> "" And Not IsNull(fldVal)) And (FractionVal <> "" And Not IsNull(FractionVal)) And (flot <> "" And Not IsNull(flot)) Then
        'values passed in ok now proceed
        'go off and get the soil vol field
        Dim firstpart, result, SoilVol
        SoilVol = DLookup("[Soil Volume]", "[Bot: Basic data]", "[Flot Number] = " & flot)
        If Not IsNull(SoilVol) Then
            firstpart = fldVal / FractionVal
            result = firstpart / SoilVol
            Calc_WoodParenceDung_ml_per_litre = result
        Else
            MsgBox "The calculation function Wood_ml_per_litre within Module ArchaeoBots_Calcs has been unable to gather the Soil Volume for this flot to undertake the calculation", vbExclamation, "Insufficient Information"
            Exit Function
        End If
    Else
        MsgBox "The calculation function Wood_ml_per_litre within Module ArchaeoBots_Calcs has not received the necessary values to undertake the calculation", vbExclamation, "Insufficient Information"
        Exit Function
    End If

Exit Function

err_catch:
    Call General_Error_Trap
    Exit Function
End Function

Function Calc_seedchaff_per_litre(flot)
'*****************************************************************************
'this function takes the Priority Sample table 4mm and 1mm fraction
'field, gathers the basic data soil vol field and then calc the seeds/chaff per litre figure.
' Inputs:   Flot = Priority Sample flot number
' all other values are gathered from the recordset
'
' Outputs: seeds/chaff per litre
' Called from: FRM_Priority - cmdAddtoPReport_Click
'              FRM_PriorityReport - cmdRecalc_Click
'
'SAJ season 2006, request from AB
'*****************************************************************************
On Error GoTo err_seedchaff

    'first check correct values passed in
    If (flot <> "" And Not IsNull(flot)) Then
        'values passed in ok now proceed
        'go off and get the soil vol field
        Dim firstpart, result, SoilVol, runningAdd, runningAdd2, FourmmCalcPart, OnemmCalcPart, AddBothCalcParts
        'get soil vol
        SoilVol = DLookup("[Soil Volume]", "[Bot: Basic data]", "[Flot Number] = " & flot)
        If Not IsNull(SoilVol) Then
            'first step add together all the 4mm values EXCEPT dung, parenc, wood, celtis
            Dim mydb As DAO.Database, myrs As DAO.Recordset
            Set mydb = CurrentDb
            Set myrs = mydb.OpenRecordset("SELECT * FROM [Bot: Priority Sample] WHERE [Flot Number] = " & flot & ";", dbOpenSnapshot)
            
            If Not (myrs.BOF And myrs.EOF) Then
                myrs.MoveFirst
                'get out 4mm fields required and add them together
                'if you need to add more simply cut and paste and edit fieldname
                If Not IsNull(myrs![4 mm barley grain]) Then runningAdd = myrs![4 mm barley grain]
                If Not IsNull(myrs![4 mm glume wheat grain]) Then runningAdd = runningAdd + myrs![4 mm glume wheat grain]
                If Not IsNull(myrs![4 mm glume wheat glume bases]) Then runningAdd = runningAdd + myrs![4 mm glume wheat glume bases]
                If Not IsNull(myrs![4 mm cereal indeterminate grain]) Then runningAdd = runningAdd + myrs![4 mm cereal indeterminate grain]
                If Not IsNull(myrs![4 mm nutshell]) Then runningAdd = runningAdd + myrs![4 mm nutshell]
                If Not IsNull(myrs![4 mm pea]) Then runningAdd = runningAdd + myrs![4 mm pea]
                If Not IsNull(myrs![4 mm culm node]) Then runningAdd = runningAdd + myrs![4 mm culm node]
                If Not IsNull(myrs![4 mm reed culm node]) Then runningAdd = runningAdd + myrs![4 mm reed culm node]
                If Not IsNull(myrs![4 mm cereal culm node]) Then runningAdd = runningAdd + myrs![4 mm cereal culm node]
                
                If Not IsNull(myrs![4 mm fraction]) Then
                    FourmmCalcPart = runningAdd / myrs![4 mm fraction]
                End If
                
                If Not IsNull(myrs![1 mm barley grain]) Then runningAdd2 = myrs![1 mm barley grain]
                If Not IsNull(myrs![1 mm barley rachis]) Then runningAdd2 = runningAdd2 + myrs![1 mm barley rachis]
                If Not IsNull(myrs![1 mm glume wheat grain]) Then runningAdd2 = runningAdd2 + myrs![1 mm glume wheat grain]
                If Not IsNull(myrs![1 mm glume wheat glume bases]) Then runningAdd2 = runningAdd2 + myrs![1 mm glume wheat glume bases]
                If Not IsNull(myrs![1 mm free-threshing wheat grain]) Then runningAdd2 = runningAdd2 + myrs![1 mm free-threshing wheat grain]
                If Not IsNull(myrs![1 mm free-threshing cereal rachis]) Then runningAdd2 = runningAdd2 + myrs![1 mm free-threshing cereal rachis]
                If Not IsNull(myrs![1 mm basal wheat rachis]) Then runningAdd2 = runningAdd2 + myrs![1 mm basal wheat rachis]
                If Not IsNull(myrs![1 mm cereal indeterminate grain]) Then runningAdd2 = runningAdd2 + myrs![1 mm cereal indeterminate grain]
                If Not IsNull(myrs![1 mm culm nodes]) Then runningAdd2 = runningAdd2 + myrs![1 mm culm nodes]
                If Not IsNull(myrs![1 mm reed culm node]) Then runningAdd2 = runningAdd2 + myrs![1 mm reed culm node]
                If Not IsNull(myrs![1 mm cereal culm node]) Then runningAdd2 = runningAdd2 + myrs![1 mm cereal culm node]
                If Not IsNull(myrs![1 mm lentil]) Then runningAdd2 = runningAdd2 + myrs![1 mm lentil]
                If Not IsNull(myrs![1 mm pea]) Then runningAdd2 = runningAdd2 + myrs![1 mm pea]
                If Not IsNull(myrs![1 mm chickpea]) Then runningAdd2 = runningAdd2 + myrs![1 mm chickpea]
                If Not IsNull(myrs![1 mm bitter vetch]) Then runningAdd2 = runningAdd2 + myrs![1 mm bitter vetch]
                If Not IsNull(myrs![1 mm pulse indeterminate]) Then runningAdd2 = runningAdd2 + myrs![1 mm pulse indeterminate]
                If Not IsNull(myrs![1 mm weed/wild seed]) Then runningAdd2 = runningAdd2 + myrs![1 mm weed/wild seed]
                If Not IsNull(myrs![1 mm Cyperaceae]) Then runningAdd2 = runningAdd2 + myrs![1 mm Cyperaceae]
                If Not IsNull(myrs![1 mm nutshell/fruitstone]) Then runningAdd2 = runningAdd2 + myrs![1 mm nutshell/fruitstone]
                If Not IsNull(myrs![1 mm fruitstone]) Then runningAdd2 = runningAdd2 + myrs![1 mm fruitstone]

                If Not IsNull(myrs![1 mm fraction]) Then
                    OnemmCalcPart = runningAdd2 / myrs![1 mm fraction]
                End If
                
                If Not IsNull(OnemmCalcPart) And Not IsNull(FourmmCalcPart) Then
                    AddBothCalcParts = OnemmCalcPart + FourmmCalcPart
                    
                    result = AddBothCalcParts / SoilVol
                    Calc_seedchaff_per_litre = result
                End If
            
            
            
            Else
                MsgBox "Flot number record cannot be found", vbCritical, "Record cannot be found"
                
            End If
            
            myrs.Close
            Set myrs = Nothing
            mydb.Close
            Set mydb = Nothing
        Else
            MsgBox "The calculation function Wood_ml_per_litre within Module ArchaeoBots_Calcs has been unable to gather the Soil Volume for this flot to undertake the calculation", vbExclamation, "Insufficient Information"
            Exit Function
        End If
    Else
        MsgBox "The calculation function Calc_seedchaff_per_litre within Module ArchaeoBots_Calcs has been unable to gather all the information for this flot to undertake the calculation", vbExclamation, "Insufficient Information"
        Exit Function
    End If

Exit Function

err_seedchaff:
    Call General_Error_Trap
    Exit Function
End Function

Function Calculate_AllPreviousPriorityRecords()
'*******************************************************************
' run once to calc historic data for new priority report table
' SAJ
'*******************************************************************
On Error GoTo err_cmdcalc
    
    Dim mydb As DAO.Database, myrs As DAO.Recordset
    Set mydb = CurrentDb
    Set myrs = mydb.OpenRecordset("SELECT * FROM [Bot: Priority Report]")
    
    If Not (myrs.EOF And myrs.BOF) Then
        myrs.MoveFirst
        Do Until myrs.EOF
            myrs.Edit
            Dim getFourmmFraction, getWood, getParenc, getDung, result1, result2, result3, result4
            getFourmmFraction = DLookup("[4 mm Fraction]", "[Bot: Priority Sample]", "[Flot Number] = " & myrs![Flot Number])
            If Not IsNull(getFourmmFraction) Then
                'calc the values required
                getWood = DLookup("[4 mm Wood]", "[Bot: Priority Sample]", "[Flot Number] = " & myrs![Flot Number])
                If Not IsNull(getWood) Then
                    result1 = Calc_WoodParenceDung_ml_per_litre(myrs![Flot Number], getWood, getFourmmFraction)
                    myrs![Wood_ml_Per_Litre] = Round(result1, 2)
                End If
    
                getParenc = DLookup("[4 mm Parenc]", "[Bot: Priority Sample]", "[Flot Number] = " & myrs![Flot Number])
                If Not IsNull(getParenc) Then
                    result2 = Calc_WoodParenceDung_ml_per_litre(myrs![Flot Number], getParenc, getFourmmFraction)
                   myrs![Parenc_ml_Per_Litre] = Round(result2, 2)
                End If
    
                getDung = DLookup("[4 mm Dung]", "[Bot: Priority Sample]", "[Flot Number] = " & myrs![Flot Number])
                If Not IsNull(getDung) Then
                    result3 = Calc_WoodParenceDung_ml_per_litre(myrs![Flot Number], getDung, getFourmmFraction)
                    myrs![Dung_ml_Per_Litre] = Round(result3, 2)
                End If
        
                result4 = Calc_seedchaff_per_litre(myrs![Flot Number])
                myrs![Seeds_Chaff_Per_Litre] = Round(result4, 2)
        
            Else
                MsgBox "The system cannot obtain the 4mm fraction value so cannot recalculate the fields", vbCritical, "Error Obtaining Fraction"
            End If
    myrs.Update
    myrs.MoveNext
    Loop
    End If
    
    
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Function

err_cmdcalc:
    Call General_Error_Trap
    Exit Function
End Function

