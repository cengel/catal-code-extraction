Option Compare Database
Option Explicit

Public Const DBName = "Human Remains Central Database"

Public VersionNumber

Public GeneralPermissions

'Public Const ImageLocationOnSite = "i:\images\"
'on site 2006 this wasn't needed as link simply to \\catal\ worked ok
Public Const ImageLocationOnSite = "H:\Catalhoyuk\images\"
Public Const sketchpath = "\\catal\recordsketches\"

'Public Const ImageLocationOnWeb = "http://www.catalhoyuk.com/siteimages/getphoto.asp"
Public Const ImageLocationOnWeb = "http://www.catalhoyuk.com/database/database_new/test/getphoto.asp"
'Public Const ImageLocationOnWeb = "http://localhost/catalweb/getphoto.asp"

Public Const VersionNumberLocal = "5.3" 'NEW 2010 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared

