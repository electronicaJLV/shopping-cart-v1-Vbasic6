VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmInventario 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   7365
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11415
   Icon            =   "frmInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmInventario.frx":E65A
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1176
      Left            =   600
      Picture         =   "frmInventario.frx":3102A0
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   4680
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
         Picture         =   "frmInventario.frx":311D7C
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
         Picture         =   "frmInventario.frx":313858
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
         Picture         =   "frmInventario.frx":315332
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
               Picture         =   "frmInventario.frx":316E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":31D8BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":324370
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":335A12
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3470B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":360716
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":376B48
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":38CF7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":393A2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":39A4DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3A0F90
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3A7A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3AE4F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3BFB96
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3D1238
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3D7CEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3E938C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3EFE3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":3F68F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":407F92
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":419634
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":42ACD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":43C378
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":44DA1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":45F0BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":473FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":488F40
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":49DE82
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":4B2DC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmInventario.frx":4C4466
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
      Height          =   7575
      Left            =   0
      Picture         =   "frmInventario.frx":4CAF18
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   865
      TabIndex        =   5
      Top             =   0
      Width           =   12972
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
         Picture         =   "frmInventario.frx":7CCB60
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
         Picture         =   "frmInventario.frx":7D1136
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
         Picture         =   "frmInventario.frx":7D570C
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
         Left            =   10320
         Picture         =   "frmInventario.frx":7D9CE2
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
         Left            =   6840
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   3372
      End
      Begin MSFlexGridLib.MSFlexGrid FListInventario 
         Height          =   5052
         Left            =   312
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   8916
         _Version        =   393216
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio  del Dolar:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   5520
         TabIndex        =   33
         Top             =   6120
         Width           =   5535
      End
   End
   Begin VB.PictureBox picFrameDatos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   84960
      Left            =   0
      Picture         =   "frmInventario.frx":7DB7BC
      ScaleHeight     =   5664
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   849
      TabIndex        =   12
      Top             =   0
      Width           =   12732
      Begin VB.PictureBox cmdCancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   9480
         Picture         =   "frmInventario.frx":ADD402
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   32
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
         Height          =   1572
         Index           =   7
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "frmInventario.frx":AE19D8
         Top             =   4200
         Width           =   7332
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
         Index           =   6
         Left            =   7680
         TabIndex        =   21
         Text            =   "000.000.000,00"
         Top             =   3480
         Width           =   2652
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
         Index           =   5
         Left            =   7680
         TabIndex        =   20
         Text            =   "000.000.000,00"
         Top             =   2880
         Width           =   2652
      End
      Begin VB.ComboBox cmbGrupo 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   4212
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
         Index           =   4
         Left            =   3000
         TabIndex        =   19
         Text            =   "000.000.000,00"
         Top             =   3480
         Width           =   2652
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
         Index           =   3
         Left            =   3000
         TabIndex        =   18
         Text            =   "000.000.000,00"
         Top             =   2880
         Width           =   2652
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
         Index           =   2
         Left            =   3000
         TabIndex        =   17
         Text            =   "123456789012"
         Top             =   2280
         Width           =   2292
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
         TabIndex        =   16
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1680
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
         Left            =   7680
         Picture         =   "frmInventario.frx":AE1A01
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
         Caption         =   "Comentario"
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
         Index           =   8
         Left            =   360
         TabIndex        =   31
         Top             =   4200
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio $"
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
         Index           =   7
         Left            =   5040
         TabIndex        =   30
         Top             =   3480
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Bs"
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
         Index           =   6
         Left            =   5040
         TabIndex        =   29
         Top             =   2880
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo $"
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
         TabIndex        =   28
         Top             =   3480
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Bs"
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
         TabIndex        =   27
         Top             =   2880
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         TabIndex        =   26
         Top             =   2280
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   2532
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         TabIndex        =   23
         Top             =   360
         Width           =   2532
      End
   End
End
Attribute VB_Name = "frmInventario"
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
    Dim constante As Single
CargarDatosdeEmpresa
CargarGrupos

With Me
    .Caption = Empresa & " - Inventario"
    .Font.Name = "Consolas"
    .Font.Size = 14
    .Font.Bold = True
    .Width = 11505
    .Height = 7800
    FListInventario.Width = 720
    FListInventario.Height = 336.8
    FListInventario.Width = Me.TextWidth(Space(11 + 21 + 14 + 14 + 9 + 3))
    cmdAdd.Left = 24
    cmdEdit.Left = 128
    cmdDel.Left = 232
    cmdFind.Left = 688
    cmdAceptar.Left = 512
    cmdCancelar.Left = 632
    lblInfo.Left = 368
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
cmdAdd.Tag = "invadd"
cmdDel.Tag = "invdel"
cmdEdit.Tag = "invedit"
cmdFind.Tag = "microfind"
cmdAceptar.Tag = "ok"
cmdCancelar.Tag = "no"
ActivarBoton Me, "ini", picFrameListado.hDC, cmdAdd
ActivarBoton Me, "ini", picFrameListado.hDC, cmdDel
ActivarBoton Me, "ini", picFrameListado.hDC, cmdEdit
ActivarBoton Me, "ini", picFrameListado.hDC, cmdFind
ActivarBoton Me, "ini", picFrameDatos.hDC, cmdAceptar
ActivarBoton Me, "ini", picFrameDatos.hDC, cmdCancelar
Inventario_Listar
End Sub

Private Sub cmdDel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
MANO
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "do", picFrameListado.hDC, cmdDel
End Sub

Private Sub cmdDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "up", picFrameListado.hDC, cmdDel
Inventario_Eliminar
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameListado.hDC, cmdAdd
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameListado.hDC, cmdAdd
Inventario_Nuevo
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
FListInventario.RowSel = 0
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MANO
End Sub

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "do", picFrameDatos.hDC, cmdAceptar
End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActivarBoton Me, "up", picFrameDatos.hDC, cmdAceptar
Inventario_Almacenar
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
MANO
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "do", picFrameListado.hDC, cmdEdit
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FListInventario.RowSel = 0 Then PROHIBIR: Exit Sub
ActivarBoton Me, "up", picFrameListado.hDC, cmdEdit
Inventario_Editar
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

For i = 1 To FListInventario.Rows - 1
    a = LCase(FListInventario.TextMatrix(i, 1))
    a = Mid(a, 1, Len(txtFind.Text))
    If a = LCase(txtFind.Text) Then
    With FListInventario
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

Private Sub Inventario_Listar()
Dim D As Currency, B As Currency
Dim IDo As Currency, IBs As Currency
Me.MousePointer = 11
FListInventario.Clear
With FListInventario
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
    .ColWidth(1) = Me.TextWidth(Space(21)) * ConstPantalla
    .ColWidth(2) = Me.TextWidth(Space(14)) * ConstPantalla
    .ColWidth(3) = Me.TextWidth(Space(14)) * ConstPantalla
    .ColWidth(4) = Me.TextWidth(Space(9)) * ConstPantalla
    .TextMatrix(0, 0) = "Código"
    .TextMatrix(0, 1) = "Descripción"
    .TextMatrix(0, 2) = PrecioUno
    .TextMatrix(0, 3) = PrecioDos
    .TextMatrix(0, 4) = "Cantidad"
End With

sqlStr = "SELECT * FROM Inventario order by Descripcion"
ABRIR_DB
i = 0
While Not Tabla.EOF
    D = Tabla("Precio $").Value
    B = D * ValorDolar
    Tabla("Precio Bs").Value = -100000
    Tabla.Update
    Tabla("Precio Bs").Value = B
    Tabla.Update
    FListInventario.AddItem Tabla("codigodebarras").Value & vbTab & _
    Tabla("descripcion").Value & vbTab & _
    Format(Tabla(PrecioUno).Value, FormatoDinero) & vbTab & _
    Format(Tabla(PrecioDos).Value, FormatoDinero) & vbTab & _
    Format(Tabla("cantidad").Value, "###,##0")
    i = i + 1
    FListInventario.RowData(i) = Tabla("referencia").Value
    IBs = IBs + Tabla("Precio Bs").Value
    IDo = IDo + Tabla("Precio $").Value
    Tabla.MoveNext
Wend
CERRAR_DB
FListInventario.RowSel = 0
lblInfo.Caption = "Valor del Dólar: " & Format(ValorDolar, FormatoDinero) _
& vbCrLf & "Inventario en $: " & Format(IDo, FormatoDinero) _
& vbCrLf & "Inventario en Bs: " & Format(IBs, FormatoDinero)
Me.MousePointer = 0
End Sub

Private Sub Inventario_Nuevo()
For i = 0 To txtDato.UBound
    txtDato(i).Text = ""
    txtDato(i).BackColor = vbWhite
Next
For i = 3 To 6
    txtDato(i).Text = "0,00" 'precio y costo
Next
txtDato(2).Text = "0" 'cantidad
cmbGrupo.Text = ""
cmbGrupo.ListIndex = -1
cmbGrupo.BackColor = vbWhite
FListInventario.RowSel = 0
txtDato(0).SetFocus
End Sub

Private Sub Inventario_Almacenar()
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
    If cmbGrupo.Text = "" Then
        errorEnCampos = True
        cmbGrupo.BackColor = &HC0C0FF
    Else
        cmbGrupo.BackColor = vbWhite
    End If
If errorEnCampos = True Then
    MsgBox "Debe completar los datos señalados.", vbInformation, "Inventario"
    Exit Sub
End If
NoActualizar = True
For i = 0 To txtDato.UBound
   If DatosTemp(i) <> txtDato(i).Text Then
        NoActualizar = False
   End If
Next
   If DatosTemp(i) <> cmbGrupo.Text Then
        NoActualizar = False
   End If
If NoActualizar = True Then
    AnimarPanel
    FListInventario.RowSel = 0
    Exit Sub
End If

DatosTemp(i) = cmbGrupo.Text
If (MsgBox("¿Está seguro de almacenar este artículo?", vbYesNo + vbQuestion, "Inventario")) = vbNo Then
    AnimarPanel
    Exit Sub
End If
i = FListInventario.RowSel
txtDato(0).BackColor = 12632256
cual = FListInventario.RowData(i)
sqlStr = "SELECT * FROM Inventario WHERE referencia = " & cual
ABRIR_DB
If Tabla.RecordCount = 0 Then nuevo = True
    If nuevo = True Then Tabla.AddNew
    Tabla("codigodebarras").Value = txtDato(0).Text
    Tabla("descripcion").Value = txtDato(1).Text
    On Error Resume Next
    m = Format(txtDato(2).Text, FormatoDinero)
    Tabla("cantidad").Value = m
    m = Format(txtDato(3).Text, FormatoDinero)
    Tabla("Costo Bs").Value = m
    m = Format(txtDato(4).Text, FormatoDinero)
    Tabla("Costo $").Value = m
    m = Format(txtDato(5).Text, FormatoDinero)
    Tabla("Precio Bs").Value = m
    m = Format(txtDato(6).Text, FormatoDinero)
    Tabla("Precio $").Value = m
    Tabla("comentario").Value = txtDato(7).Text
    Tabla("grupo").Value = cmbGrupo.ItemData(cmbGrupo.ListIndex)
    Tabla("f_vencimiento").Value = "01-01-1999"
    Tabla("f_compra").Value = "01-01-1999"
    Tabla("servicio").Value = False
    Tabla("imagen").Value = "sin imágen"

    On Error GoTo 0
    Tabla.Update
If nuevo = False Then
    FListInventario.TextMatrix(i, 0) = Tabla("codigodebarras").Value
    FListInventario.TextMatrix(i, 1) = Tabla("descripcion").Value
    FListInventario.TextMatrix(i, 2) = Format(Tabla(PrecioUno).Value, FormatoDinero)
    FListInventario.TextMatrix(i, 3) = Format(Tabla(PrecioDos).Value, FormatoDinero)
    FListInventario.TextMatrix(i, 4) = Format(Tabla("cantidad").Value, FormatoCantidad)
End If
CERRAR_DB
If nuevo = True Then Inventario_Listar

For i = 0 To txtDato.UBound
    DatosTemp(i) = txtDato(i).Text
Next
DatosTemp(i) = cmbGrupo.Text
FListInventario.RowSel = 0
AnimarPanel
End Sub

Private Sub Inventario_Editar()
i = FListInventario.RowSel
If i = 0 Then Exit Sub
txtDato(0).BackColor = 12632256
cual = FListInventario.RowData(i)
sqlStr = "SELECT * FROM Inventario WHERE referencia = " & cual
ABRIR_DB
If Tabla.RecordCount > 0 Then
    txtDato(0).Text = Tabla("codigodebarras").Value & ""
    txtDato(1).Text = Tabla("descripcion").Value & ""
    txtDato(2).Text = Format("0" & Tabla("Cantidad").Value, FormatoCantidad)
    txtDato(3).Text = Format("0" & Tabla("Costo Bs").Value, FormatoDinero)
    txtDato(4).Text = Format("0" & Tabla("Costo $").Value, FormatoDinero)
    txtDato(5).Text = Format("0" & Tabla("Precio Bs").Value, FormatoDinero)
    txtDato(6).Text = Format("0" & Tabla("Precio $").Value, FormatoDinero)
    txtDato(7).Text = Tabla("comentario").Value & ""
    cmbGrupo.Text = CualGrupo(Tabla("grupo").Value & "")
End If
CERRAR_DB
For i = 0 To txtDato.UBound
    DatosTemp(i) = txtDato(i).Text
Next
DatosTemp(i) = cmbGrupo.Text
End Sub

Private Function CualGrupo(Index As String) As String
sqlStr2 = "SELECT * FROM grupo WHERE registro = " & Index
ABRIR_DB2
i = 0
CualGrupo = ""
If Tabla2.RecordCount > 0 Then
    CualGrupo = Tabla2("nombre").Value & ""
End If
CERRAR_DB2
End Function

Private Sub Inventario_Eliminar()
i = FListInventario.RowSel
If i = 0 Then Exit Sub
a = (FListInventario.TextMatrix(i, 1))
If (MsgBox("Ha seleccionado: " & a & "." & vbCrLf & "Al eliminar este artículo no podrá recuperar los datos relacionados." & vbCrLf & vbCrLf & "¿Está seguro de eliminar este artículo?", vbYesNo + vbQuestion, "Inventario")) = vbNo Then
    Exit Sub
End If
cual = FListInventario.RowData(i)
sqlStr = "SELECT * FROM Inventario WHERE referencia =" & cual
ABRIR_DB
If Tabla.RecordCount > 0 Then
    Tabla.MoveFirst
    Tabla.Delete
    Tabla.Update
End If
CERRAR_DB
If FListInventario.Rows - FListInventario.FixedRows = 1 Then
  FListInventario.Rows = FListInventario.Rows - 1
Else
FListInventario.RemoveItem (i)
End If
FListInventario.RowSel = 0
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
    Case 2
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
        Numerizar Index, Chr(KeyAscii), False: KeyAscii = 0
    Case 3, 4, 5, 6:
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
        Numerizar Index, Chr(KeyAscii), True: KeyAscii = 0
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

Private Sub CalcularCostoBs()
Dim D As Currency, B As Currency
D = "0" & txtDato(4).Text
B = D * ValorDolar
txtDato(3).Text = Format(B, FormatoDinero)
End Sub

Private Sub CalcularCostoDol()
Dim D As Currency, B As Currency
B = "0" & txtDato(3).Text
D = B / ValorDolar
txtDato(4).Text = Format(D, FormatoDinero)
End Sub

Private Sub CalcularPrecioBs()
Dim D As Currency, B As Currency
D = "0" & txtDato(6).Text
B = D * ValorDolar
txtDato(5).Text = Format(B, FormatoDinero)
End Sub

Private Sub CalcularPrecioDol()
Dim D As Currency, B As Currency
B = "0" & txtDato(5).Text
D = B / ValorDolar
txtDato(6).Text = Format(D, FormatoDinero)
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 3: CalcularCostoDol
        Case 4: CalcularCostoBs
        Case 5: CalcularPrecioDol
        Case 6: CalcularPrecioBs
        Case Else
    End Select

End Sub

Private Sub txtDato_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 3, 4, 5, 6: txtDato(Index).SelStart = Len(txtDato(Index).Text)
        Case Else
    End Select
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdFind_MouseUp 0, 0, 0, 0
End Sub

Private Sub CargarGrupos()
sqlStr = "SELECT * FROM grupo WHERE nombre <> ''"
ABRIR_DB
i = 0
While Tabla.EOF = False
    cmbGrupo.AddItem Tabla("nombre").Value & ""
    cmbGrupo.ItemData(i) = "0" & Tabla("registro").Value
    i = i + 1
    Tabla.MoveNext
Wend
CERRAR_DB
End Sub

Public Sub AnimarPanel()
    Timer1.Enabled = True
End Sub
