Attribute VB_Name = "modReadWriteIniFiles"
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
 
'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
  ReadIni = Left(RetVal, v)
'new version:
''Public Function ReadIni(Filename As String, Section As String, Key As String) As String
''
''    Dim RetVal As String, v As Long
''    Dim retLen As Long
''    Do
''        retLen = retLen + 260 ' arbitrary - can be set higher/lower if desired
''        RetVal = Space(retLen)
''        v = GetPrivateProfileString(Section, Key, "", RetVal, retLen, Filename)
''    Loop Until v < retLen - 1
''
''    ReadIni = Left$(RetVal, v)
''
''End Function
End Function
 
'reads an Ini section
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileSection(Section, RetVal, 255, Filename)
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
''Private Sub Form_Load()
''
''Dim File As String, OFLen As Double, Str As String
''
''  File = "C:\temp.ini"
''  OFLen = FileLen(File)
''
''    'write a few example sections:
''  WriteIniSection File, "Test1", ""
''  WriteIniSection File, "Test2", "Here shoud be found some text"
''
''    'write a few ini strings
''  WriteIni File, "Test3", "Ini1", "This is ini 1"
''  WriteIni File, "Test1", "Ini2", "This is ini 2"
''
''    'test message - inform we've written the data
''  MsgBox Format((FileLen(File) - OFLen) / 1024, "0.00") & " KB data written to " & Chr(34) & File & Chr(34)
''
''    'read the ini file
''  Str = Str & "Test2 section: " & vbTab & ReadIniSection(File, "Test2") & vbCrLf
''  Str = Str & "Test1 section: " & vbTab & ReadIniSection(File, "Test1") & vbCrLf
''  Str = Str & "Ini1 string: " & vbTab & ReadIni(File, "Test3", "Ini1") & vbCrLf
''  Str = Str & "Ini2 string: " & vbTab & ReadIni(File, "Test1", "Ini2") & vbCrLf
''
''    'show the data from the file
''  MsgBox Str
''
''End Sub
