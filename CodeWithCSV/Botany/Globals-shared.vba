Option Compare Database
Option Explicit
'renamed
'Public Const DBName = "ArchaeoBotanyDatabase"

Public Const DBName = "Botany Central Database"

Public VersionNumber
Public Const VersionNumberLocal = "6" 'NEW 2009 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared

Public GeneralPermissions

Public spString 'var to hold call to sp used in DeleteSampleRecord()
