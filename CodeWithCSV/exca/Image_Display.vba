Option Compare Database
Option Explicit

Private Sub Form_Current()
'**********************************************************************
' Display images within the image box - update image control with correct
' path of image as user move through records
' This will only work for locally networked machines
' SAJ v9.1
'**********************************************************************
On Error GoTo err_Current

Dim newStr, newstr2, zeronum, FileName

'path string are coming from Portfolio with : instead of \
newStr = Replace(Me![Path], ":", "\")
'plus the directory structure of the machine they were catalogued on
' ImageLocationOnSite is global constant set in the module Globals-shared
'newstr2 = Replace(newStr, "besiktas\", ImageLocationOnSite)
''line below doesn't seem necessary at present, replaced with newstr2 = newstr
''newstr2 = Replace(newStr, "catal\", ImageLocationOnSite)
newstr2 = newStr

'this was 2008 solution as had to add certain number of zeros to id number to create
'the filename eg: 64192 becomes p0000064192.jpg
'zeronum = 10 - (Len(Me![Record_ID]))
'FileName = "p"
'        'Response.Write zeronum
'        Do While zeronum > 0
'            FileName = FileName & "0"
'            zeronum = zeronum - 1
'            'Response.Write filename
'        Loop
'
'newstr2 = newstr2 & "\" & FileName & Me![Record_ID] & ".jpg"
    
'2009 solution now portfolio is putting previews in long tree of subdirectories which
'need unpicking from filename
Dim dirpath, breakid

zeronum = 10 - (Len(Me![Record_ID]))
FileName = "p"
        'Response.Write zeronum
        Do While zeronum > 0
            FileName = FileName & "0"
            zeronum = zeronum - 1
            'Response.Write filename
        Loop
                    
FileName = FileName & Me![Record_ID]

breakid = Left(FileName, 3)
breakid = Mid(FileName, 2, 2) 'chop off leading p
dirpath = breakid & "\"

breakid = Mid(FileName, 4, 2)
dirpath = dirpath & breakid & "\"

breakid = Mid(FileName, 6, 2)
dirpath = dirpath & breakid & "\"

breakid = Mid(FileName, 8, 2)
dirpath = dirpath & breakid & "\"

'breakid = Mid(FileName, 10, 2)
'dirpath = dirpath & breakid & "\"

'MsgBox newstr2
'2008
'newstr2 = newstr2 & "\" & FileName & Me![Record_ID] & ".jpg"
newstr2 = newstr2 & "\" & dirpath & FileName & ".jpg"
Me!txtFullPath = newstr2
Me!Image145.Picture = newstr2

Exit Sub
err_Current:
    If Err.Number = 2220 Then
        'this is the error thrown if file not found
        'first check if dir exists
        If Dir(newstr2) = "" Then
            'directory not exist
            MsgBox "The directory where images are supposed to be stored cannot be found. Please contact the database administrator"
        Else
            MsgBox "The image file cannot be found - check the file exists"
            DoCmd.GoToControl "txtSketch"
        End If
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
''msgbox Me.OpenArgs
'new 2007 - year passed in with openargs determines data source as currently 2007 catalog is seperate
If Me.OpenArgs <> "" Then
    'If Me.OpenArgs = 2007 Then
    '    Me.RecordSource = "Select * from view_Portfolio_2007Previews WHERE " & Me.Filter
    'Else
    '    Me.RecordSource = "Select * from view_Portfolio_Upto2007Previews WHERE " & Me.Filter
    'End If
    '2008 one catalog
    Me.RecordSource = "Select * from view_Portfolio_Previews_2008 WHERE " & Me.Filter
    ''NEW LATE AUGUST 2009 - due to overwork of unit sheet OnCurrent it no longer checks if there
    ''are images there  but now allows user to press button and picks up here if any exist
    If Me.RecordsetClone.RecordCount <= 0 Then
        MsgBox "No images have been found in the Portfolio catalogue for this entity", vbInformation, "No images to display"
        DoCmd.Close acForm, Me.Name
    End If
    
End If

End Sub
