Attribute VB_Name = "modDatabase"
Public conn As New ADODB.Connection
Public conn2 As New ADODB.Connection
Public Tabla As New ADODB.Recordset 'Recordset Object
Public Tabla2 As New ADODB.Recordset 'Recordset Object
Public sqlStr As String ' String variable to store sql command
Public sqlStr2 As String ' String variable to store sql command´
Public TipoBaseDatos As String

'MS DATA ACTIVEX 2.8 OBJECT --> References
'MySQL ODBC 5.2 UNICODE Driver
'MySQL ODBC 3.51 Driver
'AccessConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=nwind.mdb;Uid=usuario;Pwd=clave;"
'                 Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\data.mdb;
'2013             Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\data.accdb;Uid=usuario;Pwd=clave;


Public Sub ABRIR_DB()
TipoBaseDatos = "access"
TipoBaseDatos = "access"

'On Error GoTo errortrap
Select Case TipoBaseDatos
    Case Is = "access"
        dbnombre = App.Path & "\data\datos.mdb"
        clave = "11234659"
        conn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & dbnombre & ";Password=" & clave & ";"
    Case Is = "mysql"
        servidor = "localhost"
        dbnombre = "services"
        usuario = "root"
        clave = "123456789"
        conn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & servidor & ";Database=" & dbnombre & ";UID=" & usuario & ";PWD=" & clave & ";"
End Select

conn.Open
Tabla.CursorLocation = adUseClient
Tabla.Open sqlStr, conn, adOpenStatic, adLockOptimistic
Exit Sub
errortrap:
a = Err
Select Case a
    Case -2147467259
        m = "No se puede conectar con el servidor."
    Case Else
        m = "Algo ocurrión en la conexión con la base de datos."
End Select

MsgBox m, vbCritical, "Electónica JLV"
CERRAR_DB
End
End Sub

Public Sub CERRAR_DB()
On Error Resume Next
Tabla.Close
conn.Close
End Sub

Public Sub ABRIR_DB2()

On Error GoTo errortrap
Select Case TipoBaseDatos
    Case Is = "access"
        dbnombre = App.Path & "\data\datos.mdb"
        clave = "11234659"
        conn2.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & dbnombre & ";Password=" & clave & ";"
    Case Is = "mysql"
        servidor = "localhost"
        dbnombre = "services"
        usuario = "root"
        clave = "123456789"
        conn2.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & servidor & ";Database=" & dbnombre & ";UID=" & usuario & ";PWD=" & clave & ";"
End Select

conn2.Open
Tabla2.CursorLocation = adUseClient
Tabla2.Open sqlStr2, conn2, adOpenStatic, adLockOptimistic
Exit Sub
errortrap:
MsgBox "Algo ocurrión en la conexión con la base de datos.", vbCritical, "Electónica JLV"
End
End Sub

Public Sub CERRAR_DB2()
Tabla2.Close
conn2.Close
End Sub


