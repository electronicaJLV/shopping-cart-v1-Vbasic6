Attribute VB_Name = "modSetCursor"
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private IDC_HAND As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Sub MANO()
SetCursor LoadCursor(0, 32649&)
End Sub

Public Sub PROHIBIR()
SetCursor LoadCursor(0, 32648&)
End Sub

