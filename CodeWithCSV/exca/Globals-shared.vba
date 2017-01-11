Option Compare Database
Option Explicit

Public Const DBName = "Excavation Central Database"

Public VersionNumber
Public Const VersionNumberLocal = "18.6" 'NEW 2009 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared

Public GeneralPermissions

'Public Const ImageLocationOnSite = "i:\images\"
'on site 2006 this wasn't needed as link simply to \\catal\ worked ok
Public Const ImageLocationOnSite = "H:\Catalhoyuk\images\"



'Public Const ImageLocationOnWeb = "http://www.catalhoyuk.com/siteimages/getphoto.asp"
Public Const ImageLocationOnWeb = "http://www.catalhoyuk.com/database/database_new/test/getphoto.asp"
'Public Const ImageLocationOnWeb = "http://localhost/catalweb/getphoto.asp"

Public spString 'var to hold call to sp used in Delete_Category_SubTable_Entry() on Unit Sheet

Public Const sketchpath = "\\catal\Site_Sketches\"
Public Const sketchpath2015 = "\\catal\recordsketches\"

Public ThisYear
