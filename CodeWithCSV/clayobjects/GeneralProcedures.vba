Option Compare Database
Option Explicit


Function StartUp()
'*****************************************************************************
' All actions necessary to start the system as smoothly as possible
'
' SAJ
'*****************************************************************************
On Error GoTo err_startup

'DoCmd.RunCommand acCmdWindowHide 'hide the DB window from prying eyes

DoCmd.OpenForm "ClayObjects_Login", acNormal, , , acFormEdit, acDialog

'you can hide the warning messages that Access popups up when
'you do sql tasks in the background - however the negative side to
'this is that you hide all these types of message which you may not
'want to do - the options you have are:
'   DoCmd.SetWarnings False 'turns off macro msgs
'   Application.SetOption "Confirm Record Changes", False
'   Application.SetOption "Confirm Document Deletions", False
'    Application.SetOption "Confirm Action Queries", False  'this will hide behind the scences sql actions
'you could of course turn this on an off around each statement - I'm undecided at present

'now the tables are all ok find out the current version
'SetCurrentVersion

'**** open move from marked place above
DoCmd.OpenForm "ClayObjects: Main", acNormal, , , acFormReadOnly 'open main menu

'refresh the main menu so the version number appears
'Forms![Excavation].Refresh

Exit Function

err_startup:
    Call General_Error_Trap
    'now should the system quit out here?
    'to be decided
End Function

