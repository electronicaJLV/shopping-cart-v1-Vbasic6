Attribute VB_Name = "ModIniFile"
'declarations for working with Ini files

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
    "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
 
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
    "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

 
'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 2055, v As Long
  v = GetPrivateProfileString(Section, Key, "", RetVal, 2055, Filename)
  ReadIni = Left(RetVal, v)
'new version:
''    Dim RetVal As String, v As Long
''    Dim retLen As Long
''    Do
''        retLen = retLen + 260 ' arbitrary - can be set higher/lower if desired
''        RetVal = Space(retLen)
''        v = GetPrivateProfileString(Section, Key, "", RetVal, retLen, Filename)
''    Loop Until v < retLen - 1
''
''    ReadIni = Left$(RetVal, v)
End Function
 
'reads an Ini section
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim RetVal As String * 2055, v As Long
  v = GetPrivateProfileSection(Section, RetVal, 2055, Filename)
  ReadIniSection = Left(RetVal, v - 1)
End Function
 
'writes an Ini string
Public Sub WriteIni(Filename As String, Section As String, Key As String, Value As String)
  WritePrivateProfileString Section, Key, Value, Filename
End Sub
 
'writes an Ini section
Public Sub WriteIniSection(Filename As String, Section As String, Value As String)
  WritePrivateProfileSection Section, Value, Filename
End Sub

