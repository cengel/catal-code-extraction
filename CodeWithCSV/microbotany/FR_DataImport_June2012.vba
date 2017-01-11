Option Compare Database

Sub insert_flips_data()
'sql = "INSERT INTO FR_Phytolith_sample_analysis_details ( SiteCode, SampleProcessYear, LabSampleNumber, SampleID, SingleOrMulti, PhytoName, DicotOrMonocot, PhytoCount, [PhytoN/Slide], [PhytoN/gm] )"
'sql = sql & "SELECT 'CH' AS Expr1, '09' AS Expr2, 2 AS Expr3, 'CH-09-2' AS Expr4, FR_Analysis_Details1.SorMforImport, "
'sql = sql & "FR_Analysis_Details1.PhytoName , FR_Analysis_Details1.DorMforImport, "
'sql = sql & "FR_Analysis_Details1.count02, FR_Analysis_Details1.[n/slide02], FR_Analysis_Details1.[n/gm02] "
'sql = sql & "FROM FR_Analysis_Details1;"

'DoCmd.RunSQL sql

Dim num, numtxt
'num = 3
'
'Do Until num > 143
'    numtxt = num
'    If num < 10 Then numtxt = "0" & numtxt
'
'    sql = "INSERT INTO FR_Phytolith_sample_analysis_details ( SiteCode, SampleProcessYear, LabSampleNumber, SampleID, SingleOrMulti, PhytoName, DicotOrMonocot, PhytoCount, [PhytoN/Slide], [PhytoN/gm] )"
'    sql = sql & " SELECT 'CH' AS Expr1, '09' AS Expr2, " & num & " AS Expr3, 'CH-09-" & num & "' AS Expr4, FR_Analysis_Details1.SorMforImport, "
'    sql = sql & "FR_Analysis_Details1.PhytoName , FR_Analysis_Details1.DorMforImport, "
'    sql = sql & "FR_Analysis_Details1.count" & numtxt & ", FR_Analysis_Details1.[n/slide" & numtxt & "], FR_Analysis_Details1.[n/gm" & numtxt & "] "
'    sql = sql & "FROM FR_Analysis_Details1;"
'
'    DoCmd.RunSQL sql
'    num = num + 1
'Loop

num = 56

Do Until num > 143
    numtxt = num
    If num < 10 Then numtxt = "0" & numtxt
    
    sql = "INSERT INTO FR_Phytolith_sample_analysis_details ( SiteCode, SampleProcessYear, LabSampleNumber, SampleID, SingleOrMulti, PhytoName, DicotOrMonocot, PhytoCount, [PhytoN/Slide], [PhytoN/gm] )"
    sql = sql & " SELECT 'CH' AS Expr1, '09' AS Expr2, " & num & " AS Expr3, 'CH-09-" & num & "' AS Expr4, FR_AnalysisDetails2.SorMforImport, "
    sql = sql & "FR_AnalysisDetails2.PhytoName , FR_AnalysisDetails2.DorMforImport, "
    sql = sql & "FR_AnalysisDetails2.count" & numtxt & ", FR_AnalysisDetails2.[n/slide" & numtxt & "], FR_AnalysisDetails2.[n/gm" & numtxt & "] "
    sql = sql & "FROM FR_AnalysisDetails2;"
    
    DoCmd.RunSQL sql
    num = num + 1
Loop


End Sub

