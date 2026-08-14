VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmClientes 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7365
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10980
   Icon            =   "frmClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmClientes.frx":E65A
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1176
      Left            =   360
      Picture         =   "frmClientes.frx":3102A0
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   5520
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
         Picture         =   "frmClientes.frx":311D7C
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
         Picture         =   "frmClientes.frx":313858
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
         Picture         =   "frmClientes.frx":315332
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
               Picture         =   "frmClientes.frx":316E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":31D8BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":324370
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":335A12
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3470B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":360716
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":376B48
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":38CF7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":393A2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":39A4DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3A0F90
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3A7A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3AE4F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3BFB96
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3D1238
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3D7CEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3E938C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3EFE3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":3F68F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":407F92
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":419634
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":42ACD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":43C378
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":44DA1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":45F0BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":473FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":488F40
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":49DE82
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":4B2DC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":4C4466
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFrameListado 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      Picture         =   "frmClientes.frx":4CAF18
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11070
      TabIndex        =   5
      Top             =   0
      Width           =   1.66050e5
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3480
         Top             =   240
      End
      Begin VB.PictureBox cmdAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   360
         Picture         =   "frmClientes.frx":7CCB60
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.PictureBox cmdDel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   3480
         Picture         =   "frmClientes.frx":7D1136
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.PictureBox cmdEdit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   1920
         Picture         =   "frmClientes.frx":7D570C
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.PictureBox cmdFind 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   9960
         Picture         =   "frmClientes.frx":7D9CE2
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtFind 
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
         Height          =   528
         Left            =   6480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   3372
      End
      Begin MSFlexGridLib.MSFlexGrid FListCliente 
         Height          =   5052
         Left            =   315
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   8916
         _Version        =   393216
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picFrameDatos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      Picture         =   "frmClientes.frx":7DB7BC
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   865
      TabIndex        =   12
      Top             =   0
      Width           =   12972
      Begin VB.TextBox txtDato 
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
         Height          =   1092
         Index           =   4
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmClientes.frx":ADD402
         Top             =   2880
         Width           =   7332
      End
      Begin VB.TextBox txtDato 
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
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   2280
         Width           =   7332
      End
      Begin VB.TextBox txtDato 
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
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1680
         Width           =   7332
      End
      Begin VB.PictureBox cmdCancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   9120
         Picture         =   "frmClientes.frx":ADD42B
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1455
      End
      Begin VB.TextBox txtDato 
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
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1080
         Width           =   7332
      End
      Begin VB.TextBox txtDato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Text            =   "1234567890123"
         Top             =   360
         Width           =   2532
      End
      Begin VB.PictureBox cmdAceptar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   7440
         Picture         =   "frmClientes.frx":AE1A01
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono"
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
         Index           =   5
         Left            =   360
         TabIndex        =   23
         Top             =   1680
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Correo"
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
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   2280
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres y Ap."
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
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula/RIF"
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
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   2532
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DatosTemp(20) As String

Private Sub FListCliente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub Form_Load()
CargarDatosdeEmpresa
With Me

    .Caption = Empresa & " - Clientes"
    .Font.Name = "Consolas"
    .Font.Size = 14
    .Font.Bold = True
    .Width = 11505
    .Height = 7800
    FListCliente.Width = 720
    FListCliente.Height = 336.8
    FListCliente.Width = Me.TextWidth(Space(11 + 21 + 14 + 14 + 9 + 3))
    cmdAdd.Left = 24
    cmdEdit.Left = 128
    cmdDel.Left = 232
    cmdFind.Left = 688
    cmdAceptar.Left = 512
    cmdCancelar.Left = 632
    txtFind.Left = 456
    
    picFrameListado.Width = .Width / ConstPantalla
    picFrameListado.Height = .Height / ConstPantalla
    picFrameDatos.Width = .Width / ConstPantalla
    picFrameDatos.Height = .Height / ConstPantalla
    picFrameListado.Top = 0
    picFrameListado.Left = 0
    picFrameDatos.Top = 0
    picFrameDatos.Left = 0
    picFrameListado.ZOrder 0
    Me.Refresh
End With
cmdAdd.ToolTipText = "Nuevo"
cmdDel.ToolTipText = "Eliminar"
cmdEdit.ToolTipText = "Editar"
cmdFind.ToolTipText = "Buscar"
cmdAceptar.ToolTipText = "Aceptar"
cmdCancelar.ToolTipText = "Cancelar"
cmdAdd.Tag = "clienteadd"
cmdDel.Tag = "clientedel"
cmdEdit.Tag = "clienteedit"
cmdFind.Tag = "microfind"
cmdAceptar.Tag = "ok"
cmdCancelar.Tag = "no"
ActivarBoton Me, "ini", picFrameListado.hDC, cmdAdd
ActivarBoton Me, "ini", picFrameListado.hDC, cmdDel
ActivarBoton Me, "ini", picFrameListado.hDC, cmdEdit
ActivarBoton Me, "ini", picFrameListado.hDC, cmdFind
ActivarBoton Me, "ini", picFrameDatos.hDC, cmdAceptar
ActivarBoton Me, "ini", picFrameDatos.hDC, cmdCancelar
Cliente_Listar
End Sub

Private Sub cmdDel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
MANO
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "do", picFrameListado.hDC, cmdDel
End Sub

Private Sub cmdDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "up", picFrameListado.hDC, cmdDel
Cliente_Eliminar
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameListado.hDC, cmdAdd
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameListado.hDC, cmdAdd
Cliente_Nuevo
AnimarPanel
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameDatos.hDC, cmdCancelar
End Sub

Private Sub cmdCancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameDatos.hDC, cmdCancelar
AnimarPanel
FListCliente.RowSel = 0
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameDatos.hDC, cmdAceptar
End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameDatos.hDC, cmdAceptar
Cliente_Almacenar
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
MANO
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "do", picFrameListado.hDC, cmdEdit
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListCliente.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "up", picFrameListado.hDC, cmdEdit
Cliente_Editar
AnimarPanel
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameListado.hDC, cmdFind
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameListado.hDC, cmdFind

For i = 1 To FListCliente.Rows - 1
    a = LCase(FListCliente.TextMatrix(i, 1))
    a = Mid(a, 1, Len(txtFind.Text))
    If a = LCase(txtFind.Text) Then
    With FListCliente
        .Row = i
        .Col = 0
        .RowSel = i
        .ColSel = .Cols - 1
        .TopRow = i
        .HighLight = flexHighlightAlways
        End With
        Exit For
    End If
Next
End Sub

Private Sub Cliente_Listar()
Dim D As Currency, B As Currency
Me.MousePointer = 11
FListCliente.Clear
With FListCliente
    .Font.Size = Me.Font.Size
    .Font.Name = Me.Font.Name
    .Font.Bold = Me.Font.Bold
    .Cols = 5
    .Rows = 1
    .FixedCols = 0
    .MergeCells = 0
    .TabStop = False
    .BackColorFixed = &H404040
    .ForeColorFixed = &HC0C0C0
    .GridColor = vbBlack
    .Appearance = 1
    .AllowBigSelection = False
    .FixedCols = 0
    .RowSel = 0
    .AllowUserResizing = 0
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightWithFocus
    .FocusRect = 0
    .TabStop = False
    .ColWidth(0) = Me.TextWidth(Space(11)) * ConstPantalla
    .ColWidth(1) = Me.TextWidth(Space(22)) * ConstPantalla
    .ColWidth(2) = Me.TextWidth(Space(13)) * ConstPantalla
    .ColWidth(3) = Me.TextWidth(Space(21)) * ConstPantalla
    .ColWidth(4) = Me.TextWidth(Space(21)) * ConstPantalla
    .TextMatrix(0, 0) = "Cédula"
    .TextMatrix(0, 1) = "Nombres y Apellidos"
    .TextMatrix(0, 2) = "Teléfono"
    .TextMatrix(0, 3) = "Dirección"
    .TextMatrix(0, 4) = "Correo"
End With

sqlStr = "SELECT * FROM Cliente order by Nombre"
ABRIR_DB
i = 0
While Not Tabla.EOF
    FListCliente.AddItem Tabla("cedula").Value & vbTab & _
    Tabla("Nombre").Value & vbTab & _
    Tabla("Telefono").Value & vbTab & _
    Tabla("Direccion") & vbTab & _
    Tabla("Correo").Value
    i = i + 1
    FListCliente.RowData(i) = Tabla("referencia").Value
    Tabla.MoveNext
Wend
CERRAR_DB
FListCliente.RowSel = 0
Me.MousePointer = 0
End Sub

Private Sub Cliente_Nuevo()
For i = 0 To txtDato.UBound
    txtDato(i).Text = ""
    txtDato(i).BackColor = vbWhite
Next
FListCliente.RowSel = 0
txtDato(0).SetFocus
End Sub

Private Sub Cliente_Almacenar()
Dim NoActualizar As Boolean
Dim nuevo As Boolean
Dim errorEnCampos As Boolean
Dim m As Currency
For i = 0 To 1
    If txtDato(i).Text = "" Then
        errorEnCampos = True
        txtDato(i).BackColor = &HC0C0FF
    Else
        txtDato(i).BackColor = vbWhite
    End If
Next
If errorEnCampos = True Then
    MsgBox "Debe completar los datos señalados.", vbInformation, "Cliente"
    Exit Sub
End If
NoActualizar = True
For i = 0 To txtDato.UBound
   If DatosTemp(i) <> txtDato(i).Text Then
        NoActualizar = False
   End If
Next
If NoActualizar = True Then
    AnimarPanel
    FListCliente.RowSel = 0
    Exit Sub
End If

If (MsgBox("¿Está seguro de almacenar este cliente?", vbYesNo + vbQuestion, "Cliente")) = vbNo Then
    Exit Sub
End If
i = FListCliente.RowSel
txtDato(0).BackColor = 12632256
cual = FListCliente.RowData(i)
sqlStr = "SELECT * FROM Cliente WHERE referencia = " & cual
ABRIR_DB
If Tabla.RecordCount = 0 Then nuevo = True
    If nuevo = True Then Tabla.AddNew
    Tabla("cedula").Value = txtDato(0).Text
    Tabla("Nombre").Value = txtDato(1).Text
    Tabla("Telefono").Value = txtDato(2).Text
    Tabla("Correo").Value = txtDato(3).Text
    Tabla("Direccion").Value = txtDato(4).Text
    Tabla.Update
If nuevo = False Then
    FListCliente.TextMatrix(i, 0) = Tabla("cedula").Value
    FListCliente.TextMatrix(i, 1) = Tabla("Nombre").Value
    FListCliente.TextMatrix(i, 2) = Tabla("Telefono").Value
    FListCliente.TextMatrix(i, 3) = Tabla("Direccion").Value
    FListCliente.TextMatrix(i, 4) = Tabla("Correo").Value
End If
CERRAR_DB
If nuevo = True Then Cliente_Listar

For i = 0 To txtDato.UBound
    DatosTemp(i) = txtDato(i).Text
Next
FListCliente.RowSel = 0
AnimarPanel
End Sub

Private Sub Cliente_Editar()
i = FListCliente.RowSel
If i = 0 Then Exit Sub
txtDato(0).BackColor = 12632256
cual = FListCliente.RowData(i)
sqlStr = "SELECT * FROM Cliente WHERE referencia = " & cual
ABRIR_DB
If Tabla.RecordCount > 0 Then
    txtDato(0).Text = Tabla("cedula").Value & ""
    txtDato(1).Text = Tabla("Nombre").Value & ""
    txtDato(2).Text = Tabla("Telefono").Value
    txtDato(3).Text = Tabla("Correo").Value
    txtDato(4).Text = Tabla("Direccion").Value
End If
CERRAR_DB
For i = 0 To txtDato.UBound
    DatosTemp(i) = txtDato(i).Text
Next
End Sub

Private Sub Cliente_Eliminar()
i = FListCliente.RowSel
If i = 0 Then Exit Sub
a = (FListCliente.TextMatrix(i, 1))
If (MsgBox("Ha seleccionado: " & a & "." & vbCrLf & "Al eliminar este cliente no podrá recuperar los datos relacionados." & vbCrLf & vbCrLf & "¿Está seguro de eliminar este cliente?", vbYesNo + vbQuestion, "Cliente")) = vbNo Then
    Exit Sub
End If
cual = FListCliente.RowData(i)
sqlStr = "SELECT * FROM Cliente WHERE referencia =" & cual
ABRIR_DB
If Tabla.RecordCount > 0 Then
    Tabla.MoveFirst
    Tabla.Delete
    Tabla.Update
End If
CERRAR_DB
If FListCliente.Rows - FListCliente.FixedRows = 1 Then
  FListCliente.Rows = FListCliente.Rows - 1
Else
FListCliente.RemoveItem (i)
End If
FListCliente.RowSel = 0
End Sub

Private Sub Timer1_Timer()
If UsarAnimaciones = False Then
If picFrameListado.Visible = True Then
    picFrameListado.Visible = False
    picFrameDatos.Visible = True
Else
    picFrameListado.Visible = True
    picFrameDatos.Visible = False
End If
Timer1.Enabled = False
Exit Sub
End If

Static ACCION As Integer
If ACCION = 0 Then picFrameListado.Top = picFrameListado.Top - 20
If ACCION = 1 Then picFrameListado.Top = picFrameListado.Top + 20
If picFrameListado.Top <= -picFrameListado.Height Then
    Timer1.Enabled = False
    ACCION = 1
End If
If picFrameListado.Top = 0 Then
    Timer1.Enabled = False
    ACCION = 0
End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdFind_MouseUp 0, 0, 0, 0
End Sub

Public Sub AnimarPanel()
    Timer1.Enabled = True
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
    Select Case Index
        Case 2, 3, 4, 5, 6: txtDato(Index).SelStart = Len(txtDato(Index).Text)
        Case Else
    End Select

End Sub

Private Sub txtDato_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 2, 3, 4, 5, 6:
        If KeyCode = 32 Then KeyCode = 0
        If KeyCode = 37 Then KeyCode = 0
        If KeyCode = 38 Then KeyCode = 0
        If KeyCode = 39 Then KeyCode = 0
        If KeyCode = 40 Then KeyCode = 0
        Case Else
    End Select
'If KeyCode = 13 Then
'End If
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0: If txtDato(0).BackColor = 12632256 Then KeyAscii = 0
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
        Numerizar Index, Chr(KeyAscii), False: KeyAscii = 0
    Case Else
        txtDato(Index).BackColor = vbWhite
End Select
End Sub



Private Function Numerizar(Index As Integer, k As String, decimales As Boolean) As String

Dim m As Currency
texto = txtDato(Index).Text
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

txtDato(Index).Text = texto
txtDato(Index).SelLength = 0
txtDato(Index).SelStart = Len(txtDato(Index).Text)

End Function

