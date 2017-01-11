Option Compare Database
Option Explicit

Public Const DBName = "Finds Register Central"

Public VersionNumber
Public Const VersionNumberLocal = "8.3" 'NEW 2009 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared

Public GeneralPermissions

Public spString 'var to hold call to sp used in DeleteCrateRecord()

Public logon

Public SecondCrate 'used by crate comparison tool - R_crate_comparison

Public CrateLetterFlag 'used to create a conditional query in Location Panel which will only show crates for a particular team to move inbetween

