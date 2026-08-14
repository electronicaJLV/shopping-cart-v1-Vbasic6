VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   1812
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":E65A
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdCambiodelDolar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   552
      Left            =   4800
      Picture         =   "frmMain.frx":3102A0
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   7
      ToolTipText     =   "Valor del dólar"
      Top             =   480
      Width           =   588
   End
   Begin VB.PictureBox cmdInventario 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   732
      Left            =   2160
      Picture         =   "frmMain.frx":311D7A
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   6
      ToolTipText     =   "Inventario"
      Top             =   360
      Width           =   1164
   End
   Begin VB.PictureBox cmdClientes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   732
      Left            =   480
      Picture         =   "frmMain.frx":316350
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   5
      ToolTipText     =   "Clientes"
      Top             =   360
      Width           =   1164
   End
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1176
      Left            =   1080
      Picture         =   "frmMain.frx":31A926
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1932
      Begin VB.PictureBox picAll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   960
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   4
         Top             =   240
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
         Height          =   552
         Left            =   600
         Picture         =   "frmMain.frx":31C402
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   588
      End
      Begin VB.PictureBox mask_invert 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   552
         Left            =   0
         Picture         =   "frmMain.frx":31DEDC
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   588
      End
      Begin VB.PictureBox picUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   552
         Left            =   600
         Picture         =   "frmMain.frx":31F9B6
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   588
      End
      Begin ComctlLib.ImageList BotonesImages 
         Left            =   1320
         Top             =   600
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   98
         ImageHeight     =   92
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   32
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":321492
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":327F44
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":32E9F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":340098
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":35173A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":36AD9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3811CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":397600
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":39E0B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3A4B64
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3AB616
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3B20C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3B8B7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3CA21C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3DB8BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3E2370
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3F3A12
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3FA4C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":400F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":412618
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":423CBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":43535C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4469FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4580A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":469742
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":47E684
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4935C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4A8508
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4BD44A
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4CEAEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4D559E
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":4E6C40
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CargarDatosdeEmpresa
With Me
    .Caption = Empresa & " - Menú Principal"
    .Font.Size = 16
    .Font.Name = "Consolas"
End With

cmdClientes.Tag = "clientes"
cmdInventario.Tag = "inventario"
cmdCambiodelDolar.Tag = "minidolar"
ActivarBoton Me, "ini", Me.hDC, cmdClientes
ActivarBoton Me, "ini", Me.hDC, cmdInventario
ActivarBoton Me, "ini", Me.hDC, cmdCambiodelDolar
End Sub

Private Sub cmdInventario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdInventario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", Me.hDC, cmdInventario
End Sub

Private Sub cmdInventario_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", Me.hDC, cmdInventario
frmInventario.Show
End Sub

Private Sub cmdClientes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdClientes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", Me.hDC, cmdClientes
End Sub

Private Sub cmdClientes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", Me.hDC, cmdClientes
frmClientes.Show
End Sub

Private Sub cmdCambiodelDolar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdCambiodelDolar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", Me.hDC, cmdCambiodelDolar
End Sub

Private Sub cmdCambiodelDolar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", Me.hDC, cmdCambiodelDolar
frmCambiodelDolar.Show
End Sub

