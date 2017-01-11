Option Compare Database
Option Explicit
' Copyright Lebans Holdings 1999 Ltd.


Private Sub cmdLoadPicture_Click()
' You must supply the reference to an Image Control
' when you call this Function. The FileName is Optional.
' If not supplied the File Dialog Window is called.
fLoadPicture Me.JGSForm.Form.Image1, , True
' To pass a FileName including the path
' call the Function like:
'fLoadPicture Me.Image1 , "C:\test.jpg"
' Set ScrollBars back to 0,0
' Scroll the Form back to X:0,Y:0
ScrollToHome Me.JGSForm.Form.Image1
End Sub

Private Sub CmdClip_Click()
With Me.JGSForm.Form.Image1
    If .ImageWidth <= Me.JGSForm.Form.Width - 200 Then
        .Width = .ImageWidth
    Else
        .Width = Me.JGSForm.Form.Width - 200
    End If
    
    If .ImageHeight <= Me.JGSForm.Form.Detail.Height - 200 Then
        .Height = .ImageHeight
    Else
        .Height = Me.JGSForm.Form.Detail.Height - 200
    End If
    
    .SizeMode = acOLESizeClip '0
End With

' Force ScrollBars back to Top and Left
ScrollToHome Me.JGSForm.Form.Image1
End Sub

Private Sub cmdSave_Click()
' Save Enhanced Metafile to disk
Dim blRet As Boolean
blRet = fSaveImagetoDisk(Me.JGSForm.Form.Image1)
End Sub

Private Sub CmdStretch_Click()
With Me.JGSForm.Form.Image1
    .Width = Me.JGSForm.Form.Width - 200
    .Height = Me.JGSForm.Form.Detail.Height - 200
    .SizeMode = acOLESizeStretch '3
End With
End Sub

Private Sub CmdZoom_Click()
With Me.JGSForm.Form.Image1
    .Width = Me.JGSForm.Form.Width - 200
    .Height = Me.JGSForm.Form.Detail.Height - 200
    .SizeMode = acOLESizeZoom '1
End With
End Sub

Private Sub Command46_Click()
'close this form
On Error GoTo err_cmd46

    DoCmd.Close acForm, Me.Name
Exit Sub

err_cmd46:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub Form_Activate()
'DoCmd.MoveSize 5, 5, 8500, 7000
DoCmd.MoveSize 5, 5, 12500, 10000
End Sub

Private Sub Form_Load()
'DoCmd.MoveSize 5, 5, 8500, 7000
DoCmd.MoveSize 5, 5, 12500, 10000
End Sub

Private Sub CmdBig_Click()

Dim intWidth As Integer
Dim intHeight As Integer

With Me.JGSForm.Form.Image1
    intWidth = .Width * 1.05
    intHeight = .Height * 1.05

    If intWidth < .Parent.Width Then
        .Width = intWidth
    Else
        MsgBox "Sorry, that is as Big as you can go!", vbOKOnly, "Maximum Zoom"
        Exit Sub
    End If
    
    If intHeight < .Parent.Detail.Height Then
        .Height = intHeight
    Else
        MsgBox "Sorry, that is as Big as you can go!", vbOKOnly, "Maximum Zoom"
        Exit Sub
    End If
     ' Set size mode
    .SizeMode = acOLESizeZoom
End With
' Allow Access time to Repaint Screen
' since we have Autorepeat set to TRUE for
' this Command Button
DoEvents
   
End Sub

Private Sub CmdSmall_Click()

Dim intWidth As Integer
Dim intHeight As Integer


With Me.JGSForm.Form.Image1
    intWidth = .Width * 0.95
    intHeight = .Height * 0.95

    If intWidth > 200 Then
        .Width = intWidth
    Else
        MsgBox "Sorry, that is as small as you can go!", vbOKOnly, "Minimum Zoom"
        Exit Sub
    End If
    
    If intHeight > 200 Then
        .Height = intHeight
    Else
        MsgBox "Sorry, that is as small as you can go!", vbOKOnly, "Minimum Zoom"
        Exit Sub
    End If

    .SizeMode = acOLESizeZoom
End With
' Allow Access time to Repaint Screen
' since we have Autorepeat set to TRUE for
' this Command Button
DoEvents
End Sub

Private Sub Form_Open(Cancel As Integer)
'**********************************************************************
' Display images within the image box - update image control with correct
' image for this record based on name passed in. This is only a prototype
' so at present path is hardcoded.
' SAJ july 2007
'**********************************************************************
On Error GoTo err_open

Dim Path, FileName, fname, newfile

'using global constanst Declared in globals-shared
'path = "\\catal\Site_Sketches\"
'path = sketchpath

If Me.OpenArgs <> "" Then
    Path = sketchpath2015 & "buildings\sketches\"
    FileName = Me.OpenArgs
    Path = Path & "B" & FileName & "*" & ".jpg"
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        newfile = fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "buildings\sketches\" & newfile

    Me![txtImagePath] = Path
    
    If Dir(Path) = "" Then
            'directory not exist
            MsgBox "The sketch cannot be found, it may not have been scanned in yet. The database is looking for: " & Path & " please check it exists."
            DoCmd.Close acForm, Me.Name
    Else
        'Me.Picture = path
        ' You must supply the reference to an Image Control
        ' when you call this Function. The FileName is Optional.
        ' If not supplied the File Dialog Window is called.
        fLoadPicture Me.JGSForm.Form.Image1, Me![txtImagePath], True
        ' To pass a FileName including the path
        ' call the Function like:
        'fLoadPicture Me.Image1 , "C:\test.jpg"
        ' Set ScrollBars back to 0,0
        ' Scroll the Form back to X:0,Y:0
        ScrollToHome Me.JGSForm.Form.Image1
    End If
Else
    MsgBox "No image name was passed in to this form when it was opened, system does not know which image to display. Please open from Unit sheet only", vbInformation, "No image to display"
End If
Exit Sub

err_open:
    If Err.Number = 2220 Then
        'this is the error thrown if file not found
        If Dir(Path) = "" Then
            'directory not exist
            MsgBox "The image cannot be found. The database is looking for: " & Path & " please check it exists."
        Else
            MsgBox "The image file cannot be found - check the file exists"
            'DoCmd.GoToControl "txtSketch"
        End If
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub
