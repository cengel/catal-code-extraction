Option Compare Database
Option Explicit

'Public Const DBName = "Ceramics_NEW"

Public Const DBName = "Ceramics_2013"

Public VersionNumber
Public Const VersionNumberLocal = "3.2" 'NEW 2009 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared

Public GeneralPermissions

Public spString 'var to hold call to sp used in Delete_Category_SubTable_Entry() on Unit Sheet

