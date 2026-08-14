VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCambiodelDolar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   3660
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   7230
   Icon            =   "frmCambiodelDolar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCambiodelDolar.frx":E65A
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdAceptar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   3600
      Picture         =   "frmCambiodelDolar.frx":3102A0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2289
      TabIndex        =   6
      Text            =   "000.000.000,00"
      Top             =   1200
      Width           =   2652
   End
   Begin VB.PictureBox cmdCancelar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   5280
      Picture         =   "frmCambiodelDolar.frx":314876
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1176
      Left            =   240
      Picture         =   "frmCambiodelDolar.frx":318E4C
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1932
      Begin VB.PictureBox picUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   600
         Picture         =   "frmCambiodelDolar.frx":31A928
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox mask_invert 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   0
         Picture         =   "frmCambiodelDolar.frx":31C404
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox mask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   600
         Picture         =   "frmCambiodelDolar.frx":31DEDE
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox picAll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   960
         ScaleHeight     =   44
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin ComctlLib.ImageList BotonesImages 
         Left            =   1320
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   98
         ImageHeight     =   92
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   30
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":31F9B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":32646A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":32CF1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":33E5BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":34FC60
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3692C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":37F6F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":395B26
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":39C5D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3A308A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3A9B3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3B05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3B70A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3C8742
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3D9DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3E0896
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3F1F38
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3F89EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":3FF49C
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":410B3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":4221E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":433882
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":444F24
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":4565C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":467C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":47CBAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":491AEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":4A6A2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":4BB970
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCambiodelDolar.frx":4CD012
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblArticulo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el nuevo valor del dólar:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   1
      Left            =   489
      TabIndex        =   8
      Top             =   360
      Width           =   6252
   End
End
Attribute VB_Name = "frmCambiodelDolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DatosTemp(20) As String

Private Sub cmbGrupo_Click()
cmbGrupo.BackColor = vbWhite
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbGrupo_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub FListInventario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub Form_Load()
sqlStr = "SELECT * FROM Empresa"
ABRIR_DB
If Tabla.RecordCount > 0 Then
     ValorDolar = Tabla("Dolar-precio").Value
End If
CERRAR_DB
txtDato.Text = Format(ValorDolar, FormatoDinero)
cmdAceptar.Tag = "ok"
cmdCancelar.Tag = "no"
cmdAceptar.ToolTipText = "Aceptar"
cmdCancelar.ToolTipText = "Cancelar"
ActivarBoton Me, "ini", Me.hDC, cmdAceptar
ActivarBoton Me, "ini", Me.hDC, cmdCancelar
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", Me.hDC, cmdCancelar
End Sub

Private Sub cmdCancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", Me.hDC, cmdCancelar
Unload Me
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", Me.hDC, cmdAceptar
End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", Me.hDC, cmdAceptar
Dim D As Currency, B As Currency
Me.MousePointer = 11
sqlStr = "SELECT * FROM Empresa"
ABRIR_DB
If Tabla.RecordCount > 0 Then
    Tabla("Dolar-precio").Value = -100000
    Tabla.Update
    Tabla("Dolar-precio").Value = txtDato.Text
    ValorDolar = txtDato.Text
    Tabla.Update
End If
CERRAR_DB
sqlStr = "SELECT * FROM Inventario"
ABRIR_DB
i = 0
While Not Tabla.EOF
    D = Tabla("Precio $").Value
    B = D * ValorDolar
    Tabla("Precio Bs").Value = -100000
    Tabla.Update
    Tabla("Precio Bs").Value = B
    Tabla.Update
    Tabla.MoveNext
Wend
CERRAR_DB
Me.MousePointer = 0
    MsgBox "Cambio realizado con éxito.", vbInformation, "Valor del dólar"
    Unload Me
End Sub

Private Sub txtDato_GotFocus()
txtDato.SelStart = Len(txtDato.Text)
End Sub

Private Sub txtDato_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 32 Then KeyCode = 0
        If KeyCode = 37 Then KeyCode = 0
        If KeyCode = 38 Then KeyCode = 0
        If KeyCode = 39 Then KeyCode = 0
        If KeyCode = 40 Then KeyCode = 0
End Sub

Private Function Numerizar(k As String, decimales As Boolean) As String
Dim m As Currency
texto = txtDato.Text
If k = Chr(8) Then texto = Mid(texto, 1, Len(texto) - 1): k = ""  'backspace
texto = texto & k
If Len(texto) > 13 Then texto = Mid(texto, 1, 13)                 'limita a 13 digitos
texto = Replace(texto, ",", "")
texto = Replace(texto, ".", "")
texto = texto & T
If texto = "" Then texto = "0"
m = texto
texto = Trim(Str(m))                                         'elmina los ceros adelante
D = -3
If decimales = True Then
If Len(texto) < 4 Then texto = Right(("000" & texto), 3)
texto = Left(texto, Len(texto) - 2) & "," & Right(texto, 2)   'coma decimal
D = 0
End If

If Len(texto) > 6 + D Then texto = Mid(texto, 1, (Len(texto) - (6 + D))) & "." & Right(texto, 6 + D)    'miles
If Len(texto) > 10 + D Then texto = Mid(texto, 1, (Len(texto) - (10 + D))) & "." & Right(texto, 10 + D)    'millones

txtDato.Text = texto
txtDato.SelLength = 0
txtDato.SelStart = Len(txtDato.Text)

End Function

Private Sub txtDato_KeyPress(KeyAscii As Integer)
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
        Numerizar Chr(KeyAscii), True: KeyAscii = 0
End Sub
