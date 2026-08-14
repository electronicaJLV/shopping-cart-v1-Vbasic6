Attribute VB_Name = "modDatos"
Public Empresa As String
Public Telefono As String
Public Correo As String
Public Rif As String
Public Direccion As String
Public Calle As String
Public Municipio As String
Public Estado As String
Public Ciudad As String
'+++++++++++++++++++++++++++++
Public ValorDolar As Currency
Public PrecioUno As String
Public PrecioDos As String
Public FormatoDinero As String
Public FormatoCantidad As String
Public UsarAnimaciones As Boolean
Public MiMaquina As Boolean
Public ConstPantalla As Integer

Sub CargarDatosdeEmpresa()
ConstPantalla = Screen.TwipsPerPixelX
a = LCase(ReadIni(App.Path & "\data\lista.ini", "general", "Usar Animaciones"))
If a = "si" Then
    UsarAnimaciones = True
Else
    UsarAnimaciones = False
End If

FormatoDinero = "###,###,##0.00"
FormatoCantidad = "###,##0"
PrecioDos = "Precio Bs"
PrecioUno = "Precio $"

sqlStr = "SELECT * FROM Empresa"
ABRIR_DB
If Tabla.RecordCount > 0 Then
    ValorDolar = Tabla("Dolar-precio").Value
    Empresa = Tabla("nombre").Value & ""
    Direccion = Tabla("Direccion").Value & ""
    Telefono = Tabla("Telefono").Value & ""
    Correo = Tabla("Correo").Value & ""
    Rif = Tabla("Rif").Value & ""
    Direccion = Tabla("Direccion").Value & ""
    Calle = Tabla("Calle").Value & ""
    Municipio = Tabla("Municipio").Value & ""
    Estado = Tabla("Estado").Value & ""
    Ciudad = Tabla("Ciudad").Value & ""
End If
CERRAR_DB
End Sub


