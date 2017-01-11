Option Compare Database

Private Sub Befehl0_Click()
On Error GoTo Err_Befehl0_Click


    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdFind

Exit_Befehl0_Click:
    Exit Sub

Err_Befehl0_Click:
    MsgBox Err.Description
    Resume Exit_Befehl0_Click
    
End Sub
Private Sub Befehl1_Click()
On Error GoTo Err_Befehl1_Click

    Dim stDocName As String

    stDocName = "Exca: Building Sheet"
    DoCmd.OpenReport stDocName, acNormal

Exit_Befehl1_Click:
    Exit Sub

Err_Befehl1_Click:
    MsgBox Err.Description
    Resume Exit_Befehl1_Click
    
End Sub


Private Sub Form_Load()

Dim prt As Printer

For Each prt In Printers
    Me.cbo_Printer.AddItem prt.DeviceName
Next prt

End Sub

Private Sub print_bulk_Click()
On Error GoTo Err_print_bulk_Click

Dim retval, retvalprint
Dim msg
Dim checknum
Dim prt As Printer

If Not IsNull(Me.cbo_Printer.Value) Or Not IsNull(Me.Path) Then

If Not IsNull(Me.cbo_Printer.Value) Then
msg = "Do you want to send all units in between " & Me!unit_from & " and " & Me!unit_to & " to the printer (" & Me.cbo_Printer.Value & ")?"

retval = MsgBox(msg, vbInformation + vbYesNo, "print bulk")
        
            If retval = vbNo Then
                MsgBox "Ok, units will not be printed!", vbExclamation, "noprinting!"
            Else
                Application.Printer = Application.Printers(Me.cbo_Printer.Value)
                
                For I = unit_from To unit_to
                    checknum = DLookup("[Category]", "[Exca: Unit Sheet]", "[Unit Number] = " & I)
                    If Not IsNull(checknum) Then
                        If LCase(checknum) = "layer" Or LCase(checknum) = "cluster" Then
                            reportname = "R_Unit_Sheet_layercluster"
                        ElseIf LCase(checknum) = "cut" Then
                            reportname = "R_Unit_Sheet_cut"
                        ElseIf LCase(checknum) = "skeleton" Then
                            reportname = "R_Unit_Sheet_skeleton"
                        End If
                        
                        If reportname <> "" Then
                            DoCmd.OpenReport reportname, acPreview, , "[unit number] = " & I, acHidden
                            Set Reports(reportname).Printer = Application.Printer
                            DoCmd.OpenReport reportname, acViewNormal, , "[unit number] = " & I
                            DoCmd.Close acReport, reportname
                        Else
                        
                        End If
                    Else
                        Debug.Print "Unit " & I & " does not contain enough information (category)!", vbExclamation, "nocategory!"
                    End If
                    checknum = ""
                    reportname = ""
                
                Next I
            
            End If
ElseIf Not IsNull(Me.Path) Then

msg = "Do you want to export all units in between " & Me!unit_from & " and " & Me!unit_to & " as pdfs to " & Me.Path & "?"

retval = MsgBox(msg, vbInformation + vbYesNo, "pdf bulk")
        
            If retval = vbNo Then
                MsgBox "Ok, units will not be exported!", vbExclamation, "nopdf!"
            Else
                
                For I = unit_from To unit_to
                    checknum = DLookup("[Category]", "[Exca: Unit Sheet]", "[Unit Number] = " & I)
                    If Not IsNull(checknum) Then
                        If LCase(checknum) = "layer" Or LCase(checknum) = "cluster" Then
                            reportname = "R_Unit_Sheet_layercluster"
                        ElseIf LCase(checknum) = "cut" Then
                            reportname = "R_Unit_Sheet_cut"
                        ElseIf LCase(checknum) = "skeleton" Then
                            reportname = "R_Unit_Sheet_skeleton"
                        End If
                        
                        If reportname <> "" Then
                            DoCmd.OpenReport reportname, acViewPreview, , "[unit number] = " & I
                            DoCmd.OutputTo acOutputReport, "", acFormatPDF, Path & "\U" & I & ".pdf", False
                            DoCmd.Close acReport, reportname
                        Else
                        
                        End If
                    Else
                        Debug.Print "Unit " & I & " does not contain enough information (category)!", vbExclamation, "nocategory!"
                    End If
                    checknum = ""
                    reportname = ""
                
                Next I
            
            End If


End If

Else
    MsgBox "You have to select a printer or enter a pathname first!", vbExclamation, "noprinter!"
End If


Exit Sub

Err_print_bulk_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
