Option Compare Database
Option Explicit

Private Sub Form_Current()
'**********************************************************************
' Display images within the image box - update image control with correct
' path of image as user move through records
' This will only work for locally networked machines
' SAJ
'**********************************************************************
On Error GoTo err_current

Dim newStr, newstr2, zeronum, FileName

'path string are coming from Portfolio with : instead of \
'obtain from view
Dim mydb, rspath
Set mydb = CurrentDb
Set rspath = mydb.OpenRecordset("view_Portfolio_Preview_Path", dbOpenSnapshot)
If rspath.BOF And rspath.EOF Then
    'problem getting path this will not work
    MsgBox "Path of previews cannot be located. Sorry the images cannot be viewed", vbInformation, "Preview Path Missing"
    rspath.Close
    Set rspath = Nothing
    mydb.Close
    Set mydb = Nothing
    DoCmd.Close acForm, Me.Name
Else
    rspath.MoveFirst
    If rspath!Path = "" Then
        'problem getting path this will not work
        MsgBox "Path of previews cannot be located. Sorry the images cannot be viewed", vbInformation, "Preview Path Missing"
        rspath.Close
        Set rspath = Nothing
        mydb.Close
        Set mydb = Nothing
        DoCmd.Close acForm, Me.Name
    Else
    
        newStr = Replace(rspath![Path], ":", "\")
        rspath.Close
        Set rspath = Nothing
        mydb.Close
        Set mydb = Nothing
        
        'plus the directory structure of the machine they were catalogued on
        ' ImageLocationOnSite is global constant set in the module Globals-shared
        'newstr2 = Replace(newStr, "besiktas\", ImageLocationOnSite)
        ''line below doesn't seem necessary at present, replaced with newstr2 = newstr
        ''newstr2 = Replace(newStr, "catal\", ImageLocationOnSite)
        newstr2 = newStr

        'MsgBox newstr2
        '2009 solution now portfolio is putting previews in long tree of subdirectories which
        'need unpicking from filename
        Dim dirpath, breakid
    
        zeronum = 10 - (Len(Me![record_id]))
        FileName = "p"
                'Response.Write zeronum
                Do While zeronum > 0
                    FileName = FileName & "0"
                    zeronum = zeronum - 1
                    'Response.Write filename
                Loop
                        
        FileName = FileName & Me![record_id]
    
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
        newstr2 = newstr2 & "\" & dirpath & FileName & ".jpg"
        Me!txtFullPath = newstr2
        Me!Image145.Picture = newstr2


        'Me!Image145.Picture = newstr2
    End If
End If
Exit Sub
err_current:
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
