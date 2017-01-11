Option Compare Database
Option Explicit

Private Sub cmdGoDB_Click()
On Error GoTo err_handler

If Me![FullConservation_Ref] <> "" Then
    
    'old code does not work with runtime as not take into account mdw
     'Dim appAccess As Access.Application
     Dim appAccess As New Access.Application
    
     Dim dbpath
     dbpath = Replace(CurrentDb.Name, "Finds Register Central.mdb", "Conservation Central Database.mdb")
     '' Open database in Microsoft Access window.
     appAccess.OpenCurrentDatabase dbpath, False 'false = shared
     '' Open given form.
     
    Set appAccess = getobject(dbpath)
    'having it as an object reference allows use of the openform command
    'to form - it will hold focus till shut
    'appAccess.DoCmd.OpenForm "Conserv: Basic Record", acNormal, , "FullConservation_Ref = '00.088'", , acDialog
    ''appAccess.DoCmd.OpenForm "'Conserv: Basic Record'", acNormal, , , , acDialog
    'conservation form opens straight away so just need to make vis
    appAccess.Visible = True
    appAccess.DoCmd.Close acForm, "Conserv: Basic Record"
    appAccess.DoCmd.OpenForm "Conserv: Basic Record", acNormal, , "FullConservation_Ref = '" & Me![FullConservation_Ref] & "'", , acDialog
    'once shut the code will continue - so must close connection to DB
    appAccess.CloseCurrentDatabase
    'and quit the shell that it is runningin
    appAccess.Quit
    Set appAccess = Nothing
    '--
End If
Exit Sub

err_handler:

    If Err.Number = -2147467259 Or Err.Number = 432 Then
       ' the user has entered the incorrect password - must shut the underlying shell
'       If KillProcess(shellobj, 0) Then
'         'MsgBox "App was terminated"
'       End If
       
       'Set shellobj = Null
        If Err.Number = -2147467259 Then MsgBox "The system cannot open the scheme database - you must enter a valid password", vbCritical, "Scheme cannot be opened"
        If Err.Number = 432 Then MsgBox "The system cannot open the scheme database. The path to the database entered into the Scheme Administration screen cannot be found", vbCritical, "Scheme cannot be opened"
        DoCmd.Hourglass False
    Else
        DoCmd.Hourglass False
    
        MsgBox "An error has occurred in General Procedures - OpenASchemeAtGivenForm(). The error is:" & Chr(13) & Chr(13) & Err.Number & " -- " & Err.Description
    
        appAccess.CloseCurrentDatabase
        'and quit the shell that it is runningin
        appAccess.Quit
        Set appAccess = Nothing
    End If
    Exit Sub

End Sub
