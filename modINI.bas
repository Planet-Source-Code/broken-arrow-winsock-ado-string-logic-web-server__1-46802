Attribute VB_Name = "modINI"
Option Explicit

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetINIString(INISection As String, INIKey As String, Optional INIFile As String, Optional DefaultValue As String = "") As String
If INIFile = "" Then INIFile = CheckPath(App.Path) & App.Title & ".INI"
Dim Buffer As String * 255
GetPrivateProfileString INISection, INIKey, DefaultValue, Buffer, Len(Buffer), INIFile
GetINIString = Buffer
End Function

Public Function GetINIInt(INISection As String, INIKey As String, Optional INIFile As String, Optional DefaultValue As Long = 0) As Long
If INIFile = "" Then INIFile = CheckPath(App.Path) & App.Title & ".INI"
GetINIInt = GetPrivateProfileInt(INISection, INIKey, DefaultValue, INIFile)
End Function

Public Sub SetINI(INISection As String, INIKey As String, Value As String, Optional INIFile As String)
If INIFile = "" Then INIFile = CheckPath(App.Path) & App.Title & ".INI"
WritePrivateProfileString INISection, INIKey, Value, INIFile
End Sub

Private Function CheckPath(Path As String) As String
If Right(Path, 1) <> "\" Then CheckPath = Path & "\" Else CheckPath = Path
End Function
