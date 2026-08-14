Attribute VB_Name = "modDBHandler"
Public conn As New ADODB.Connection
Public Tabla As New ADODB.Recordset 'Recordset Object
Public Tabla2 As New ADODB.Recordset 'Recordset Object
Public sqlStr As String ' String variable to store sql command
Public Const IDC_HAND = 32649&
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ABRIR_DB()
clave = "11234659"
'On Error GoTo errortrap
conn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & App.Path & "\data\datos.mdb;Password=" & clave & ";"
conn.Open
Tabla.CursorLocation = adUseClient
Tabla.Open sqlStr, conn, adOpenStatic, adLockOptimistic
Exit Sub
errortrap:
MsgBox "Algo ocurrión en la conexión con la base de datos.", vbCritical, "Electónica JLV"
End
End Sub

Public Sub CERRAR_DB()
Tabla.Close
: conn.Close
End Sub





