VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Generales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USUARIOS"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8490
   Begin VB.PictureBox Pic_Bancos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   30
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   98
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Bancos_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bancos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   75
         TabIndex        =   104
         Top             =   2490
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Bancos 
            Height          =   3150
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   5556
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Bancos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   75
         TabIndex        =   100
         Top             =   390
         Width           =   8300
         Begin VB.TextBox Txt_RFC_Banco 
            Height          =   315
            Left            =   1567
            TabIndex        =   200
            Top             =   990
            Width           =   2655
         End
         Begin VB.TextBox Txt_Consecutivo_Cheque 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7020
            Locked          =   -1  'True
            TabIndex        =   199
            Top             =   1335
            Width           =   1215
         End
         Begin VB.TextBox Txt_cheque_inicial 
            Height          =   315
            Left            =   5850
            TabIndex        =   7
            Top             =   1335
            Width           =   1125
         End
         Begin VB.TextBox Txt_Numero_Cuenta 
            Height          =   315
            Left            =   5850
            TabIndex        =   6
            Top             =   990
            Width           =   2385
         End
         Begin VB.TextBox Txt_Clave_Interbancaria 
            Height          =   315
            Left            =   1567
            TabIndex        =   2
            Top             =   1335
            Width           =   2655
         End
         Begin VB.ComboBox Cmb_Formato 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0000
            Left            =   1567
            List            =   "Frm_Cat_Generales.frx":0002
            TabIndex        =   3
            ToolTipText     =   "Formato que Usara para Impresiones."
            Top             =   1695
            Width           =   2655
         End
         Begin VB.TextBox Ttx_Sucursal 
            Height          =   315
            Left            =   5850
            MaxLength       =   100
            TabIndex        =   5
            Top             =   645
            Width           =   2385
         End
         Begin VB.TextBox Txt_Banco_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1567
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox Txt_Nombre_Banco 
            Height          =   315
            Left            =   1567
            MaxLength       =   100
            TabIndex        =   1
            Top             =   645
            Width           =   2655
         End
         Begin VB.ComboBox Cmb_Estatus_Banco 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0004
            Left            =   5850
            List            =   "Frm_Cat_Generales.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   300
            Width           =   2385
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RFC"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   201
            Top             =   1035
            Width           =   390
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cheque Inicial"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4260
            TabIndex        =   198
            Top             =   1380
            Width           =   1320
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero de Cuenta"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4260
            TabIndex        =   197
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clabe Interbancaria"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   196
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Lbl_Formato 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formato"
            Height          =   195
            Left            =   90
            TabIndex        =   165
            Top             =   1755
            Width           =   570
         End
         Begin VB.Label Lbl_Sucursal_Banco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4260
            TabIndex        =   105
            Top             =   690
            Width           =   765
         End
         Begin VB.Label Lbl_Banco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   103
            Top             =   690
            Width           =   540
         End
         Begin VB.Label Lbl_Banco_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   102
            Top             =   345
            Width           =   795
         End
         Begin VB.Label Llb_Estatus_Banco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4260
            TabIndex        =   101
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Label Lbl_Titulo_Bancos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "BANCOS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3285
         TabIndex        =   99
         Top             =   -15
         Width           =   1620
      End
   End
   Begin VB.CommandButton Btn_Cargar_Categorias 
      Caption         =   "Cargar Categoria"
      Height          =   330
      Left            =   5895
      TabIndex        =   169
      Top             =   7950
      Width           =   1755
   End
   Begin VB.CommandButton Btn_Alta_de_Categorias 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7665
      TabIndex        =   168
      Top             =   7950
      Width           =   465
   End
   Begin VB.CommandButton Btn_Agregar_Presentaciones 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7740
      TabIndex        =   167
      Top             =   7500
      Width           =   465
   End
   Begin VB.Data Dt_Excel 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5865
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   6855
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton Btn_Cargar_Presentaciones 
      Caption         =   "Cargar Presentaciones "
      Height          =   330
      Left            =   5940
      TabIndex        =   166
      Top             =   7500
      Width           =   1755
   End
   Begin MSComDlg.CommonDialog Cdg_Exel 
      Left            =   7965
      Top             =   6765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Btn_Nuevo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   60
      Picture         =   "Frm_Cat_Generales.frx":0024
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "A"
      Top             =   6075
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7035
      Picture         =   "Frm_Cat_Generales.frx":355B
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "A"
      Top             =   6090
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5280
      Picture         =   "Frm_Cat_Generales.frx":6C5A
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "B"
      Top             =   6090
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3510
      Picture         =   "Frm_Cat_Generales.frx":A214
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "C"
      Top             =   6090
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1800
      Picture         =   "Frm_Cat_Generales.frx":D7A0
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "M"
      Top             =   6090
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.PictureBox Pic_Marcas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   60
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   127
      Top             =   15
      Width           =   8400
      Begin VB.Frame Fra_Detalles_Marcas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Marcas"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   90
         TabIndex        =   133
         Top             =   2010
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Marcas 
            Height          =   3675
            Left            =   135
            TabIndex        =   19
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Marcas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   90
         TabIndex        =   128
         Top             =   465
         Width           =   8300
         Begin VB.ComboBox Cmb_Estatus_Marca 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10ED1
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":10EDB
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   1700
         End
         Begin VB.TextBox Txt_Marca_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox Txt_Comentarios_Marca 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1005
            Width           =   6870
         End
         Begin VB.TextBox Txt_Nombre_Marca 
            Height          =   315
            Left            =   1305
            TabIndex        =   17
            Top             =   660
            Width           =   6870
         End
         Begin VB.Label Lbl_Estatus_Marca 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   132
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Lbl_Comentarios_Marca 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   131
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   130
            Top             =   345
            Width           =   675
         End
         Begin VB.Label Lbl_Nombre_Marca 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   129
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Marcas"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3660
         TabIndex        =   134
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.PictureBox Pic_Categorias 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   0
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   119
      Top             =   -15
      Width           =   8400
      Begin VB.Frame Fra_Generales_Categorias 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   60
         TabIndex        =   121
         Top             =   405
         Width           =   8300
         Begin VB.TextBox Txt_Catgoria_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox Txt_Nombre_Categorias 
            Height          =   315
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   22
            Top             =   660
            Width           =   6870
         End
         Begin VB.TextBox Txt_Comentarios_Categorias 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   1005
            Width           =   6870
         End
         Begin VB.ComboBox Cmb_Estatus_Categoria 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10EF1
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":10EFB
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   300
            Width           =   1700
         End
         Begin VB.Label Lbl_Nombre_Categorias 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   125
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Lbl_Categoria_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoría ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   124
            Top             =   345
            Width           =   915
         End
         Begin VB.Label Lbl_Comentarios_Categorias 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   123
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Estatus_Categorias 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   122
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Frame Fra_Detalles_Categorias 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Presentaciones"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   60
         TabIndex        =   120
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Categorias 
            Height          =   3675
            Left            =   120
            TabIndex        =   24
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Categorias 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Categorias"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3420
         TabIndex        =   126
         Top             =   -60
         Width           =   1860
      End
   End
   Begin VB.PictureBox Pic_Cat_Productos_Tipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   -30
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   151
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Productos_Tipo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   105
         TabIndex        =   155
         Top             =   465
         Width           =   8300
         Begin VB.ComboBox Cmb_Estatus_Cat_Productos_Tipo 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10F11
            Left            =   6450
            List            =   "Frm_Cat_Generales.frx":10F1B
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   210
            Width           =   1700
         End
         Begin VB.TextBox Txt_Comentarios_Cat_Productos_Tipo 
            Height          =   435
            Left            =   1275
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   158
            Top             =   915
            Width           =   6870
         End
         Begin VB.TextBox Txt_Nombre_Cat_Productos_Tipo 
            Height          =   315
            Left            =   1275
            TabIndex        =   157
            Top             =   570
            Width           =   6870
         End
         Begin VB.TextBox Txt_Tipo_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   156
            Top             =   210
            Width           =   1995
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5580
            TabIndex        =   163
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   162
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label Lbl_Tipo_Id 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   161
            Top             =   255
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   160
            Top             =   630
            Width           =   660
         End
      End
      Begin VB.Frame Fra_Cat_Produstos_Tipo_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Presentaciones"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   105
         TabIndex        =   153
         Top             =   1995
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Productos_Tipo 
            Height          =   3675
            Left            =   120
            TabIndex        =   154
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipos Productos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   152
         Top             =   0
         Width           =   2820
      End
   End
   Begin VB.PictureBox Pic_Cat_Sustancia_Activa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   45
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   183
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Detalles_Cat_Sustancia_Activa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sustancia Activa"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   0
         TabIndex        =   193
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Sustancia_Activa 
            Height          =   3675
            Left            =   105
            TabIndex        =   194
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Cat_Sustancia_Activa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   15
         TabIndex        =   184
         Top             =   405
         Width           =   8300
         Begin VB.TextBox Txt_Comentarios_Cat_Sustancia_Activa 
            Height          =   435
            Left            =   1650
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   187
            Top             =   1005
            Width           =   6525
         End
         Begin VB.ComboBox Cmb_Cat_Sustancia_Activa 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10F31
            Left            =   5985
            List            =   "Frm_Cat_Generales.frx":10F3B
            Style           =   2  'Dropdown List
            TabIndex        =   188
            Top             =   300
            Width           =   2190
         End
         Begin VB.TextBox Txt_Nombre_Cat_Sustancia_Activa 
            Height          =   315
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   186
            Top             =   645
            Width           =   6525
         End
         Begin VB.TextBox Txt_ID_Cat_Sustancia_Activa 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   185
            Top             =   300
            Width           =   2190
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5160
            TabIndex        =   192
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   191
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lab_Sustancia_Activa_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sustancia Activa ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   190
            Top             =   345
            Width           =   1425
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   189
            Top             =   690
            Width           =   660
         End
      End
      Begin VB.Label Lbl_Sustancia_Activa 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sustancia  Activa"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2835
         TabIndex        =   195
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.PictureBox Pic_Cat_Impuestos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   45
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   170
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Generales_Cta_Impuestos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   30
         TabIndex        =   173
         Top             =   405
         Width           =   8300
         Begin VB.TextBox Txt_ID_Cat_Impuestos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   177
            Top             =   315
            Width           =   2535
         End
         Begin VB.TextBox Txt_Impuesto_Cat_Impuestos 
            Height          =   315
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   176
            Top             =   645
            Width           =   2535
         End
         Begin VB.TextBox Txt_Comentarios_Cat_Impuestos 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   175
            Top             =   1005
            Width           =   6870
         End
         Begin VB.ComboBox Cmb_Estatus_Cat_Impuestos 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10F51
            Left            =   5100
            List            =   "Frm_Cat_Generales.frx":10F5B
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   660
            Width           =   3075
         End
         Begin VB.Label Lbl_Impuesto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   181
            Top             =   705
            Width           =   780
         End
         Begin VB.Label Lbl_Impuesto_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   180
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Lbl_comentarios_Cat_Impuestos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   179
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Estatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4245
            TabIndex        =   178
            Top             =   705
            Width           =   660
         End
      End
      Begin VB.Frame Fra_Detalles_Cat_Impuestos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Impuestos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   15
         TabIndex        =   171
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Impuestos_Cat_Impuestos 
            Height          =   3675
            Left            =   105
            TabIndex        =   172
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Impuestos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Impuestos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3465
         TabIndex        =   182
         Top             =   0
         Width           =   1770
      End
   End
   Begin VB.PictureBox Pic_Presentaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   -15
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   106
      Top             =   15
      Width           =   8400
      Begin VB.Frame Fra_Detalles_Presentaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Presentaciones"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   75
         TabIndex        =   116
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Presentaciones 
            Height          =   3675
            Left            =   120
            TabIndex        =   117
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Presentaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   75
         TabIndex        =   107
         Top             =   405
         Width           =   8300
         Begin VB.ComboBox Cmb_Estaus_Presentaciones 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10F71
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":10F7B
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   300
            Width           =   1700
         End
         Begin VB.TextBox Txt_Comentarios_Presentaciones 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   110
            Top             =   1005
            Width           =   6870
         End
         Begin VB.TextBox Txt_Nombre_Presentaciones 
            Height          =   315
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   109
            Top             =   660
            Width           =   6870
         End
         Begin VB.TextBox Txt_Presentacion_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Lbl_Estatus_Precentaciones 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   115
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   114
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Presentcion_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presentación ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   113
            Top             =   345
            Width           =   1170
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   112
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Label Lbl_Presentaciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Presentaciones"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3090
         TabIndex        =   118
         Top             =   0
         Width           =   2670
      End
   End
   Begin VB.PictureBox Pic_Clasificacion_Proveedores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   45
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   82
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Detalles_Clasificacion_Proveedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clasificación"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   90
         TabIndex        =   86
         Top             =   1980
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Clasificaciones 
            Height          =   3675
            Left            =   90
            TabIndex        =   164
            Top             =   225
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Clasificacion_Proveedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   90
         TabIndex        =   83
         Top             =   420
         Width           =   8300
         Begin VB.TextBox Txt_Clasificacion_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   315
            Width           =   1700
         End
         Begin VB.ComboBox Cmb_Estatus_Clasificacion_Proveedor 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":10F91
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":10F9B
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   315
            Width           =   1700
         End
         Begin VB.TextBox Txt_Comentarios_Clasficacion_Proveedor 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   1035
            Width           =   6870
         End
         Begin VB.TextBox Txt_Nombre_Clasificacion_Proveedor 
            Height          =   315
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   37
            Top             =   675
            Width           =   6870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   89
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   88
            Top             =   1140
            Width           =   915
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Label Lbl_Clasificacion_Proveedores 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CLASIFICACIÓN DE PROVEEDORES"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   87
         Top             =   30
         Width           =   6660
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   615
         Picture         =   "Frm_Cat_Generales.frx":10FB1
         Top             =   150
         Width           =   360
      End
   End
   Begin VB.PictureBox Pic_Clasificacion_Clientes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   75
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   90
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Clasificacion_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   45
         TabIndex        =   92
         Top             =   405
         Width           =   8300
         Begin VB.TextBox Txt_Clasificacion_Cliente 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox Txt_Nombre_Clasificacion_Cliente 
            Height          =   315
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   32
            Top             =   675
            Width           =   6870
         End
         Begin VB.TextBox Txt_Comentarios_Clasificacion_Cliente 
            Height          =   435
            Left            =   1305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   1005
            Width           =   6870
         End
         Begin VB.ComboBox Cmb_Estatus_Clasificacion_Cliente 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":14505
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":1450F
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   300
            Width           =   1700
         End
         Begin VB.Label Lbl_Nombre_Clasificacion_Cliente 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   96
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Lbl_Clasificacion_ID_Clientes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   95
            Top             =   345
            Width           =   1125
         End
         Begin VB.Label Lbl_Comentarios_Clasificacion_Cliente 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   94
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Estatus_Clasificacion_Clientes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   93
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Frame Fra_Clasificacion_Clientes_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clasificación"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   60
         TabIndex        =   91
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Clasificacion_Clientes 
            Height          =   3675
            Left            =   120
            TabIndex        =   34
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Image Image4 
         Height          =   360
         Left            =   1185
         Picture         =   "Frm_Cat_Generales.frx":14525
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CLASIFICACIÓN DE CLIENTES"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1620
         TabIndex        =   97
         Top             =   0
         Width           =   5610
      End
   End
   Begin VB.PictureBox Pic_Apl_Cat_Usuarios 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   45
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   53
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Generales_Usuarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   90
         TabIndex        =   54
         Top             =   405
         Width           =   8300
         Begin VB.ComboBox Cmb_Estatus_Usuario 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":17A79
            Left            =   6460
            List            =   "Frm_Cat_Generales.frx":17A83
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   240
            Width           =   1700
         End
         Begin VB.ComboBox Cmb_Roles 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":17A99
            Left            =   1035
            List            =   "Frm_Cat_Generales.frx":17AA3
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1410
            Width           =   4530
         End
         Begin VB.TextBox Txt_Contraseña 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   3865
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   45
            Top             =   1020
            Width           =   1700
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Caducar_Usuario 
            Height          =   315
            Left            =   6460
            TabIndex        =   48
            Top             =   1410
            Width           =   1700
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   108331011
            CurrentDate     =   39367
         End
         Begin VB.TextBox Txt_Confirmar_Contraseña 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   6460
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   46
            Top             =   1020
            Width           =   1700
         End
         Begin VB.TextBox Txt_Login 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1035
            MaxLength       =   20
            TabIndex        =   44
            Top             =   1020
            Width           =   1700
         End
         Begin VB.TextBox Txt_Comentarios_Usuarios 
            Height          =   465
            Left            =   1035
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   1800
            Width           =   7125
         End
         Begin VB.TextBox Txt_Nombre_Usuario 
            Height          =   315
            Left            =   1035
            MaxLength       =   100
            TabIndex        =   9
            Top             =   630
            Width           =   7125
         End
         Begin VB.TextBox Txt_Usuario_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   1700
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   80
            Top             =   300
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   79
            Top             =   1470
            Width           =   540
         End
         Begin VB.Label Lbl_Confirmar_Contraseña 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirmar"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   78
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Lbl_Login 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Lbl_Comentarios 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   60
            Top             =   1935
            Width           =   915
         End
         Begin VB.Label Lbl_Contraseña 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2805
            TabIndex        =   59
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Lbl_Nombre_Usuario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   58
            Top             =   690
            Width           =   660
         End
         Begin VB.Label Lbl_Usuario_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   57
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Lbl_Tipo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rol"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   56
            Top             =   1470
            Width           =   285
         End
      End
      Begin VB.Frame Fra_Usuarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   90
         TabIndex        =   62
         Top             =   2775
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Usuarios 
            Height          =   2880
            Left            =   120
            TabIndex        =   63
            Top             =   210
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   2775
         Picture         =   "Frm_Cat_Generales.frx":17AC0
         Top             =   60
         Width           =   360
      End
      Begin VB.Label Lbl_USUARIOS 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIOS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   64
         Top             =   -15
         Width           =   2010
      End
   End
   Begin VB.PictureBox Pic_Apl_Cat_Roles 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   30
      ScaleHeight     =   6045
      ScaleWidth      =   8400
      TabIndex        =   65
      Top             =   0
      Width           =   8400
      Begin VB.CommandButton Btn_Acceso_Seguridad 
         Caption         =   "Control de Acceso"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5790
         TabIndex        =   52
         Tag             =   "C"
         Top             =   1395
         Width           =   2400
      End
      Begin VB.Frame Fra_Generales_Roles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   67
         Top             =   450
         Width           =   8220
         Begin VB.TextBox Txt_Comentarios_Rol 
            Height          =   375
            Left            =   1100
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   945
            Width           =   4530
         End
         Begin VB.TextBox Txt_Nombre_Rol 
            Height          =   285
            Left            =   1100
            MaxLength       =   100
            TabIndex        =   50
            Top             =   585
            Width           =   6990
         End
         Begin VB.TextBox Txt_Rol_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1100
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   225
            Width           =   2000
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   71
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   70
            Top             =   630
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rol ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   69
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Fra_Roles_Sistema 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Roles del Sistema"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4065
         Left            =   120
         TabIndex        =   68
         Top             =   1920
         Width           =   8220
         Begin MSFlexGridLib.MSFlexGrid Grid_Roles 
            Height          =   3705
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   6535
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Acceso_Sistema_Rol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accesos del Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4065
         Left            =   100
         TabIndex        =   74
         Top             =   1935
         Visible         =   0   'False
         Width           =   8180
         Begin VB.TextBox Txt_Habilitar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   77
            Top             =   450
            Visible         =   0   'False
            Width           =   500
         End
         Begin VB.CheckBox Chk_Habilitar_Menu_Submenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Habilitar"
            Height          =   200
            Left            =   225
            TabIndex        =   75
            Top             =   405
            Visible         =   0   'False
            Width           =   900
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Accesos_Seguridad 
            Height          =   3705
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   6535
            _Version        =   393216
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   1440
         Picture         =   "Frm_Cat_Generales.frx":1B014
         Top             =   60
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROLES DE SEGURIDAD"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2025
         TabIndex        =   66
         Top             =   45
         Width           =   4335
      End
   End
   Begin VB.PictureBox Pic_Cat_Almacenes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   15
      ScaleHeight     =   6045
      ScaleWidth      =   8460
      TabIndex        =   143
      Top             =   15
      Width           =   8460
      Begin VB.Frame Fra_Cat_Almacenes_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laboratorios"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   45
         TabIndex        =   149
         Top             =   1965
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Almacenes 
            Height          =   3675
            Left            =   135
            TabIndex        =   14
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Cat_Almacenes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   45
         TabIndex        =   144
         Top             =   405
         Width           =   8300
         Begin VB.ComboBox Cmb_Estatus_Cat_Almacenes 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":1E4E6
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":1E4F0
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   1700
         End
         Begin VB.TextBox Txt_Almacen_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox Txt_Comentarios_Cat_Almacenes 
            Height          =   435
            Left            =   1380
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1005
            Width           =   6795
         End
         Begin VB.TextBox Txt_Nombre_Cat_Almacenes 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   12
            Top             =   660
            Width           =   6795
         End
         Begin VB.Label Lbl_Estatus_Cat_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   148
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Lbl_Comentarios_Cat_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   147
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Almacen_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Almacen ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   146
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre_Cat_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   145
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Label Lbl_Almacenes 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacenes"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3435
         TabIndex        =   150
         Top             =   0
         Width           =   1890
      End
   End
   Begin VB.PictureBox Pic_Laboratorios 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   -15
      ScaleHeight     =   6045
      ScaleWidth      =   8460
      TabIndex        =   135
      Top             =   0
      Width           =   8460
      Begin VB.Frame Fra_Generales_Laboratorio 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   90
         TabIndex        =   137
         Top             =   405
         Width           =   8300
         Begin VB.TextBox Txt_Nombre_Laboratorios 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   27
            Top             =   660
            Width           =   6795
         End
         Begin VB.TextBox Txt_Cometarios_Laboratorios 
            Height          =   435
            Left            =   1380
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1005
            Width           =   6795
         End
         Begin VB.TextBox Txt_Laborartorio_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   300
            Width           =   1995
         End
         Begin VB.ComboBox Cmb_Estatus_Laboratorios 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":1E506
            Left            =   6475
            List            =   "Frm_Cat_Generales.frx":1E510
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   300
            Width           =   1700
         End
         Begin VB.Label Lbl_NombreLaboratorios 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   141
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Lbl_Laboratorio_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Laboratorio ID"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   140
            Top             =   345
            Width           =   1245
         End
         Begin VB.Label Lbl_Comentarios_Laboratorios 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   139
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Lbl_Estatus_Laboratorios 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5610
            TabIndex        =   138
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Frame Fra_Detalles_Laboratorios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laboratorios"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   90
         TabIndex        =   136
         Top             =   1950
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Laboratorios 
            Height          =   3675
            Left            =   135
            TabIndex        =   29
            Top             =   255
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Laboratoris 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratorios"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   142
         Top             =   0
         Width           =   2160
      End
   End
End
Attribute VB_Name = "Frm_Cat_Generales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Collapse_Grid
    'DESCRIPCIÓN: Hace las filas del grid con respecto al height igual a 0
    'PARÁMETROS :
    'CREO       : Oscar Alcantara
    'FECHA_CREO :
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Collapse_Grid()
    If Grid_Accesos_Seguridad.Rows > 0 Then
        Grid_Accesos_Seguridad.FixedRows = 1
        For Renglon_Procesar = 1 To Grid_Accesos_Seguridad.Rows - 1
            If Grid_Accesos_Seguridad.TextMatrix(Renglon_Procesar, 0) = "-" Then
                Grid_Accesos_Seguridad.Col = 1
                Call Grid_Accesos_Seguridad_Click
            End If
        Next Renglon_Procesar
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Rol
    'DESCRIPCIÓN: Da de alta un nuevo registro con los dtoa del rol que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Yazmin A Delgado Gómez
    'FECHA_CREO : 28-Mayo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Rol()
Dim Rs_Alta_Apl_Cat_Roles As rdoResultset    'Da de alta el  nuevo rol en la base de datos
Dim Rs_Alta_Apl_Cat_Accesos As rdoResultset  'Manejo de registro de Apl_Cat_Accesos, guarda a que menus son lo que va a tener acceso el rol en el sistema
Dim Menus As Integer                         'Contador que sirve para ver en que posición me encuentro en el grid
Dim Ctl As Control                           'Indica que tipo de control es el que se esta consultando de la pantalla principal

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Da de alta el rol en la base de datos
    Set Rs_Alta_Apl_Cat_Roles = Conectar_Ayudante.Recordset_Agregar("Apl_Cat_Roles")
    With Rs_Alta_Apl_Cat_Roles
        .AddNew
            .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Rol.text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Rol.text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Apl_Cat_Roles.Close
    'Da de alta los menus y submenus al cual va a tener acceso el rol
    Set Rs_Alta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Agregar("Apl_Cat_Accesos")
    'Llena el Grid con los datos actualizados
    For Menus = 1 To Grid_Accesos_Seguridad.Rows - 1
        With Rs_Alta_Apl_Cat_Accesos
            .AddNew
                .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.text)
                If Grid_Accesos_Seguridad.TextMatrix(Menus, 0) <> "" Then
                    .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 1))
                End If
                If Grid_Accesos_Seguridad.TextMatrix(Menus, 1) <> "" Then
                    .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 2))
                End If
                .rdoColumns("Nombre_Sistema") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 3))
                .rdoColumns("Tipo") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 4))
                .rdoColumns("Habilitar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 5))
                .rdoColumns("Alta") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 6))
                .rdoColumns("Cambio") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 7))
                .rdoColumns("Eliminar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 8))
                .rdoColumns("Consultar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 9))
            .Update
        End With
    Next Menus
    Rs_Alta_Apl_Cat_Accesos.Close
    Conexion_Base.CommitTrans
    'Envia mensaje de rol daddo de alta
    If Grid_Roles.Rows = 0 Then
        Grid_Roles.AddItem "Rol ID" & Chr(9) & "Nombre del Rol" & Chr(9) & "Comentarios"
    End If
    Grid_Roles.AddItem Trim(Txt_Rol_ID.text) & Chr(9) & Trim(UCase(Txt_Nombre_Rol.text)) _
    & Chr(9) & Trim(UCase(Txt_Comentarios_Rol.text))
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Roles.FixedRows = 1
    Grid_Roles.ColWidth(0) = 1600
    Grid_Roles.ColWidth(1) = 5000
    Grid_Roles.ColWidth(2) = 0
    Fra_Roles_Sistema.Visible = True
    Fra_Acceso_Sistema_Rol.Visible = False
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Acceso_Seguridad.Visible = True
    Btn_Acceso_Seguridad.Caption = "Control de Acceso"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Roles", Frm_Cat_Generales)
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    MsgBox "El rol ha sido dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Nombre_Rol
    'DESCRIPCIÓN: Consulta si ya se encuentra asigando el nombre del rol que se
    '             pretende dar de alta, si encuentra el nombre dado de alta
    '             entonces manda un valor de verdadero si no manda un valor de
    '             falso
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 29-Mayo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Function Consulta_Nombre_Rol() As Boolean
Dim Rs_Consulta_Apl_Cat_Roles As rdoResultset 'Consulta si existe ya el nombre del rol en la base de datos

'Consulta si ya esta asignado el nombre del rol que se pretende dar de alta
Mi_SQL = "SELECT Rol_ID, Nombre FROM Apl_Cat_Roles"
Mi_SQL = Mi_SQL & " WHERE Nombre='" & Trim(UCase(Txt_Nombre_Rol.text)) & "'"
If Btn_Modificar.Enabled = True Then
    Mi_SQL = Mi_SQL & " AND Rol_ID <> '" & Trim(Txt_Rol_ID.text) & "'"
End If
Set Rs_Consulta_Apl_Cat_Roles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Roles.EOF Then
    Consulta_Nombre_Rol = True
Else
    Consulta_Nombre_Rol = False
End If
Rs_Consulta_Apl_Cat_Roles.Close
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Roles
    'DESCRIPCIÓN: Consulta los roles que se tienen dados de alta
    'PARÁMETROS : Nombre: Indica el nombre del rol que se pretende buscar
    'CREO       : Yazmin A. Delgado Gómez
    'FECHA_CREO : 28-MAYO-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Consulta_Roles(Nombre As String)
Dim Rs_Consulta_Apl_Roles As rdoResultset 'Consulta los roles que se encuentran dados de alta

Grid_Roles.Rows = 0
'Consulta todos los roles que se encuentran dados de alta
Mi_SQL = "SELECT * FROM Apl_Cat_Roles"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Nombre"
Set Rs_Consulta_Apl_Roles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Roles.EOF Then
    With Rs_Consulta_Apl_Roles
        Grid_Roles.AddItem "Rol ID" & Chr(9) & "Nombre del Rol" & Chr(9) & "Comentarios"
        While Not .EOF
            Grid_Roles.AddItem .rdoColumns("Rol_ID") & Chr(9) & .rdoColumns("Nombre") _
            & Chr(9) & .rdoColumns("Comentarios")
            .MoveNext
        Wend
    End With
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Roles.FixedRows = 1
    Grid_Roles.ColWidth(0) = 1550
    Grid_Roles.ColWidth(1) = 6000
    Grid_Roles.ColWidth(2) = 0
End If
Rs_Consulta_Apl_Roles.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Rol
    'DESCRIPCIÓN: Modifica los datos del registro del rol que selecciono el
    '             usuario así como elimina y da de alta los menus y submenus
    '             que tiene dados de alta el sistema
    'PARÁMETROS :
    'CREO       : Yazmin A Delgado Gómez
    'FECHA_CREO : 28-Mayo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Rol()
Dim Rs_Modifica_Apl_Cat_Roles As rdoResultset 'Modifica el registro del rol que fue seleccionado por el usuario
Dim Rs_Alta_Apl_Cat_Accesos As rdoResultset   'Da de alta los accesos que va a contener el rol
Dim Menus As Integer                          'Contador que sirve para ver en que posición me encuentro en el grid
Dim Ctl As Control                            'Indica que tipo de control es el que se esta consultando de la pantalla principal

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
        'Consulta los datos que tiene asignado el rol que fue seleccionado por el usuario
        Mi_SQL = "SELECT * FROM Apl_Cat_Roles"
        Mi_SQL = Mi_SQL & " WHERE Rol_ID = '" & Trim(Txt_Rol_ID.text) & "'"
        Set Rs_Modifica_Apl_Cat_Roles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Modifica_Apl_Cat_Roles.EOF Then
            With Rs_Modifica_Apl_Cat_Roles
                .Edit
                    .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Rol.text))
                    .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Rol.text))
                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                    .rdoColumns("Fecha_Modifico") = Now
                .Update
            End With
        End If
        Rs_Modifica_Apl_Cat_Roles.Close
        'Si se elimino correctamente los menus y submenus que se tenian asignados entonces
        'da de alta nuevamente estos mismos
        If Conectar_Ayudante.Elimina_Catalogo("Apl_Cat_Accesos", "Rol_ID", Trim(Txt_Rol_ID)) = True Then
            'Da de alta los menus y submenus al cual va a tener acceso el rol
            Set Rs_Alta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Agregar("Apl_Cat_Accesos")
            'Llena el Grid con los datos actualizados
            For Menus = 1 To Grid_Accesos_Seguridad.Rows - 1
                With Rs_Alta_Apl_Cat_Accesos
                    .AddNew
                        .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.text)
                        If Grid_Accesos_Seguridad.TextMatrix(Menus, 1) <> "" Then
                            .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 1))
                        End If
                        If Grid_Accesos_Seguridad.TextMatrix(Menus, 2) <> "" Then
                            .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 2))
                        End If
                        .rdoColumns("Nombre_Sistema") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 3))
                        .rdoColumns("Tipo") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 4))
                        .rdoColumns("Habilitar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 5))
                        .rdoColumns("Alta") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 6))
                        .rdoColumns("Cambio") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 7))
                        .rdoColumns("Eliminar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 8))
                        .rdoColumns("Consultar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 9))
                    .Update
                End With
            Next Menus
            Rs_Alta_Apl_Cat_Accesos.Close
        End If
    Conexion_Base.CommitTrans
    Grid_Roles.TextMatrix(Grid_Roles.RowSel, 1) = Trim(Txt_Nombre_Rol.text)
    Grid_Roles.TextMatrix(Grid_Roles.RowSel, 2) = Trim(Txt_Comentarios_Rol.text)
    Fra_Roles_Sistema.Visible = True
    Fra_Acceso_Sistema_Rol.Visible = False
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Acceso_Seguridad.Visible = True
    Btn_Acceso_Seguridad.Caption = "Control de Acceso"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Roles", Frm_Cat_Generales)
    MsgBox "El rol " & Trim(UCase(Txt_Nombre_Rol.text)) & Chr(13) & Chr(13) & _
           "ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Acceso_Seguridad_Click()
    If Fra_Roles_Sistema.Visible = True Then
        Fra_Acceso_Sistema_Rol.Visible = True
        Fra_Roles_Sistema.Visible = False
        Btn_Acceso_Seguridad.Caption = "Roles"
    Else
        Fra_Acceso_Sistema_Rol.Visible = False
        Fra_Roles_Sistema.Visible = True
        Btn_Acceso_Seguridad.Caption = "Control de Acceso"
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Usuarios
    'DESCRIPCIÓN: Consulta todos los Usuarios que hay en la tabla Cat_Usuarios
    '             llenando el Grid
    'PARÁMETROS : Nombre: Indica el nombre del rol que se pretende buscar
    'CREO       : Jorge Razo
    'FECHA_CREO :
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Usuarios(Nombre As String)
Dim Rs_Consulta_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios
Set Conectar_Ayudante = New Ayudante
    
Grid_Usuarios.Rows = 0
'Consulta los datos generales del usuario
Mi_SQL = "SELECT Usuario_ID, Nombre, Login, Password"
Mi_SQL = Mi_SQL & " FROM Apl_Cat_Usuarios"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Usuario_ID"
Set Rs_Consulta_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Apl_Cat_Usuarios.EOF Then
    'Coloca un encabezado en el grid
    Grid_Usuarios.AddItem "Usuario ID" & Chr(9) & "Nombre" & Chr(9) & "Login" & Chr(9) & "Password"
    While Not Rs_Consulta_Apl_Cat_Usuarios.EOF
        Grid_Usuarios.AddItem Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Usuario_ID") _
        & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Nombre") _
        & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Login") _
        & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Password")
        Grid_Usuarios.FixedRows = 1
        Rs_Consulta_Apl_Cat_Usuarios.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid_usuarios
    Grid_Usuarios.ColWidth(0) = 1000
    Grid_Usuarios.ColWidth(1) = 5000
    Grid_Usuarios.ColWidth(2) = 1550
    Grid_Usuarios.ColWidth(3) = 0
End If
'Cierra el manejador del registro
Rs_Consulta_Apl_Cat_Usuarios.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Usuarios
    'DESCRIPCIÓN: Da de alta un nuevo registro en la base de datos con los datos
    '             que introdujo el usuario
    'PARÁMETROS :
    'CREO       :
    'FECHA_CREO :
    'MODIFICO          : Yazmin A Delgado Gómez
    'FECHA_MODIFICO    : 28-Mayo-2007
    'CAUSA_MODIFICACIÓN: Porque la seguridad del sistema no es a nivel usuario
    '                    como en un principio se encontraba
'*******************************************************************************
Private Sub Alta_Usuarios()
Dim Rs_Alta_Apl_Cat_Usuarios As rdoResultset 'Manejo del registro de Apl_Cat_Usuarios, da de alta los datos del usuario

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de Usuario
    Set Rs_Alta_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Agregar("Apl_Cat_Usuarios")
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Alta_Apl_Cat_Usuarios
        .AddNew
            .rdoColumns("Usuario_ID") = Trim(Txt_Usuario_ID.text)
            .rdoColumns("Rol_ID") = Format(Cmb_Roles.ItemData(Cmb_Roles.ListIndex), "00000")
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Usuario.text))
            .rdoColumns("Login") = Trim(Txt_Login.text)
            .rdoColumns("Password") = Trim(Txt_Contraseña.text)
            .rdoColumns("Estatus") = Trim(Cmb_Estatus_Usuario.text)
            .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Ultimo_Acceso") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Sesion_Abierta") = "NO"
            If Txt_Comentarios_Usuarios.text = "" Then
                .rdoColumns("Comentarios") = " "
            Else
            .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Usuarios.text)
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Apl_Cat_Usuarios.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Usuarios.Enabled = False
    Fra_Usuarios.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Usuario.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Usuarios.Rows = 0 Then
        Grid_Usuarios.AddItem "Usuario" & Chr(9) & "Nombre" & Chr(9) & "Login" & Chr(9) & "Password"
    End If
    'Llena el grid con los datos del nuevo usuario
    Grid_Usuarios.AddItem Trim(Txt_Usuario_ID.text) & Chr(9) & _
    UCase(Txt_Nombre_Usuario.text) & Chr(9) & Trim(Txt_Login.text) & Chr(9) & Trim(Txt_Contraseña.text)
    Grid_Usuarios.ColWidth(0) = 1000
    Grid_Usuarios.ColWidth(1) = 5000
    Grid_Usuarios.ColWidth(2) = 1550
    Grid_Usuarios.ColWidth(3) = 0
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Usuarios", Frm_Cat_Generales)
    MsgBox "Usuario dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Usuarios
    'DESCRIPCIÓN: Modifica al Usuario por medio del grid
    'PARÁMETROS :
    'CREO       : Jorge Razo
    'FECHA_CREO        :
    'MODIFICO          : Yazmin A Delgado Gómez
    'FECHA_MODIFICO    : 28-Mayo-2007
    'CAUSA_MODIFICACIÓN: Porque se quito la seguridad a nivel usuario
'*******************************************************************************
Private Sub Modifica_Usuarios()
Dim Rs_Modificacion_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro de la tabla Cat_Usuarios, modifica los valores del registro que tiene el usuario seleccionado

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID ='" & Trim(Txt_Usuario_ID.text) & "'"
    Set Rs_Modificacion_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Apl_Cat_Usuarios
        .Edit
            .rdoColumns("Rol_ID") = Format(Cmb_Roles.ItemData(Cmb_Roles.ListIndex), "00000")
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Usuario.text))
            .rdoColumns("Login") = Trim(Txt_Login.text)
            .rdoColumns("Password") = Trim(Txt_Contraseña.text)
            .rdoColumns("Estatus") = Trim(Cmb_Estatus_Usuario.text)
            .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Ultimo_Acceso") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Sesion_Abierta") = "NO"
            If Txt_Comentarios_Usuarios.text = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Usuarios.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Apl_Cat_Usuarios.Close
    Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 1) = Trim(UCase(Txt_Nombre_Usuario.text))
    Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 2) = Trim(Txt_Login.text)
    Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 3) = Trim(Txt_Contraseña.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Usuarios.Enabled = False
    Fra_Usuarios.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Usuarios", Frm_Cat_Generales)
    MsgBox "El usuario ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Agregar_Presentaciones_Click
'DESCRIPCION            : Da de alta las presentaciones contenidas en el grid de presentaciones
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''
Private Sub Btn_Agregar_Presentaciones_Click()
Dim Rs_Alta_Presentaciones As rdoResultset
Dim Cont_Fila As Integer
Dim Presentacion_ID As String


On Error GoTo handler
    Conexion_Base.BeginTrans
    
    Set Rs_Alta_Presentaciones = Conectar_Ayudante.Recordset_Agregar("Cat_Presentaciones")
        For Cont_Fila = 1 To Grid_Presentaciones.Rows - 1 Step 1
            Presentacion_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Presentaciones", "Presentacion_ID"), "00000")
            With Rs_Alta_Presentaciones
                .AddNew
                    .rdoColumns("Presentacion_ID") = Presentacion_ID
                    .rdoColumns("Nombre") = Grid_Presentaciones.TextMatrix(Cont_Fila, 0)
                    .rdoColumns("Estatus") = "ACTIVO"
                    .rdoColumns("Comentarios") = " "
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
        Next
    Rs_Alta_Presentaciones.Close
    Conexion_Base.CommitTrans
    MsgBox "Registros dados de Alta"
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    MsgBox Er.Description
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Alta_de_Categorias_Click
'DESCRIPCION            : Da de alta las Categorias contenidas en el grid de presentaciones
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''
Private Sub Btn_Alta_de_Categorias_Click()
Dim Rs_Alta_Categorias As rdoResultset
Dim Cont_Fila As Integer
Dim Categoria_ID As String


On Error GoTo handler
    Conexion_Base.BeginTrans
    
    Set Rs_Alta_Categorias = Conectar_Ayudante.Recordset_Agregar("Cat_Categorias")
        For Cont_Fila = 1 To Grid_Cat_Categorias.Rows - 1 Step 1
            Categoria_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Categorias", "Categoria_ID"), "00000")
            With Rs_Alta_Categorias
                .AddNew
                    .rdoColumns("Categoria_ID") = Categoria_ID
                    .rdoColumns("Nombre") = Grid_Cat_Categorias.TextMatrix(Cont_Fila, 0)
                    .rdoColumns("Estatus") = "ACTIVO"
                    .rdoColumns("Comentarios") = " "
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
        Next
    Rs_Alta_Categorias.Close
    Conexion_Base.CommitTrans
    MsgBox "Registros dados de Alta"
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    MsgBox Er.Description
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Cargar_Categorias_Click
'DESCRIPCION            : Carga las categorias del archivo de Eexcel seleccionado
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''
Private Sub Btn_Cargar_Categorias_Click()
Dim Clave As String
Dim Cantidad As String
Dim Descripcion As String
Dim Presentacion As String
Dim Marca As String
Dim Cont_Partidas As Integer
Dim Total_Piezas As Integer
Dim Total_Partidas_Sin_Cantidad As Integer
Dim Mi_SQL As String
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Edita_Cat_Productos As rdoResultset
Dim Editar_Clave As String
Dim Rs_Cat_Productos_Descripcion As rdoResultset
Dim Tem_Descripcion As String
Dim Tem_Producto_ID As String
Dim path_XLS As String
Dim Presentacion_ID As String
Dim Cont_Fila As Integer
Dim Agregar_Presentacion As String

    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Grid_Cat_Categorias.Rows = 0
    With Cdg_Exel
        .DialogTitle = " Seleccione archivo Excel para cargar"
        .Filter = "Archivos xls|*.xls"
        .ShowOpen
        If .fileName = "" Then Exit Sub
        path_XLS = .fileName
    End With
    
    On Local Error GoTo handler
        With Dt_Excel
            .DatabaseName = path_XLS
            'Asigna el Recordsource al control data
            .RecordSource = "Hoja1$"
            Grid_Presentaciones.Redraw = False
            .Refresh
            Grid_Presentaciones.Redraw = True
        End With
        Grid_Cat_Categorias.Rows = 0
        Grid_Cat_Categorias.Cols = 1
        Descripcion = ""
               
        'Se agrega encabezado
        Grid_Cat_Categorias.AddItem "Nombre"
        With Dt_Excel.Recordset
            While Not .EOF
                    If Not IsNull(Dt_Excel.Recordset(6).Value) Then Descripcion = Dt_Excel.Recordset(6).Value
                    Agregar_Presentacion = "SI"
                    For Cont_Fila = 1 To Grid_Cat_Categorias.Rows - 1 Step 1
                        If Trim(Descripcion) = Trim(Grid_Cat_Categorias.TextMatrix(Cont_Fila, 0)) Then
                            'no agregar
                            Agregar_Presentacion = "NO"
                            Exit For
                        End If
                    Next Cont_Fila
                    If Agregar_Presentacion = "SI" And Trim(Descripcion) <> "" Then
                        Grid_Cat_Categorias.AddItem Descripcion
                    End If
                    Descripcion = ""
                .MoveNext
            Wend
            
            'Configura el tamaño las columnas del Grid
            If Grid_Cat_Categorias.Rows > 1 Then
                Grid_Cat_Categorias.ColWidth(0) = 5000 'Descripcion
                'Pone el setfocus en la primera fila del Grid
                With Grid_Cat_Categorias
                    ''.Col = 0
                    ''.Row = 1
                    ''.ColSel = .Cols - 1
                    ''.RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
        End With
        Exit Sub
handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Cargar_Archivo_Click
'DESCRIPCION            : Carga las presentaciones del archivo de Eexcel seleccionado
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''
Private Sub Btn_Cargar_Presentaciones_Click()
Dim Clave As String
Dim Cantidad As String
Dim Descripcion As String
Dim Presentacion As String
Dim Marca As String
Dim Cont_Partidas As Integer
Dim Total_Piezas As Integer
Dim Total_Partidas_Sin_Cantidad As Integer
Dim Mi_SQL As String
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Edita_Cat_Productos As rdoResultset
Dim Editar_Clave As String
Dim Rs_Cat_Productos_Descripcion As rdoResultset
Dim Tem_Descripcion As String
Dim Tem_Producto_ID As String
Dim path_XLS As String
Dim Presentacion_ID As String
Dim Cont_Fila As Integer
Dim Agregar_Presentacion As String

    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Grid_Presentaciones.Rows = 0
    With Cdg_Exel
        .DialogTitle = " Seleccione archivo Excel para cargar"
        .Filter = "Archivos xls|*.xls"
        .ShowOpen
        If .fileName = "" Then Exit Sub
        path_XLS = .fileName
    End With
    
    On Local Error GoTo handler
        With Dt_Excel
            .DatabaseName = path_XLS
            'Asigna el Recordsource al control data
            .RecordSource = "Hoja1$"
            Grid_Presentaciones.Redraw = False
            .Refresh
            Grid_Presentaciones.Redraw = True
        End With
        Grid_Presentaciones.Rows = 0
        Grid_Presentaciones.Cols = 1
        Descripcion = ""
               
        'Se agrega encabezado
        Grid_Presentaciones.AddItem "Nombre"
        With Dt_Excel.Recordset
            While Not .EOF
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then Descripcion = Dt_Excel.Recordset(3).Value
                    Agregar_Presentacion = "SI"
                    For Cont_Fila = 1 To Grid_Presentaciones.Rows - 1 Step 1
                        If Trim(Descripcion) = Trim(Grid_Presentaciones.TextMatrix(Cont_Fila, 0)) Then
                            'no agregar
                            Agregar_Presentacion = "NO"
                            Exit For
                        End If
                    Next Cont_Fila
                    If Agregar_Presentacion = "SI" And Trim(Descripcion) <> "" Then
                        Grid_Presentaciones.AddItem Descripcion
                    End If
                    Descripcion = ""
                .MoveNext
            Wend
            
            'Configura el tamaño las columnas del Grid
            If Grid_Presentaciones.Rows > 1 Then
                Grid_Presentaciones.ColWidth(0) = 5000 'Descripcion
                'Pone el setfocus en la primera fila del Grid
                With Grid_Presentaciones
                    ''.Col = 0
                    ''.Row = 1
                    ''.ColSel = .Cols - 1
                    ''.RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
        End With
        Exit Sub
handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub


Private Sub Btn_Consultar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar

    Nombre = InputBox("Proporcione el nombre", UCase(MDIFrm_Apl_Principal.Caption))
    Select Case Catalogo
        Case "USUARIOS"
            Consulta_Usuarios Nombre
        Case "ROLES"
            Consulta_Roles Nombre
        Case "CLASIFICACION_PROVEEDORES"
            Consulta_Clasificacion_Proveedores Nombre
        Case "CLASIFICACION_CLIENTES"
            Consulta_Clasificacion_Clientes Nombre
        Case "BANCOS"
             Consulta_Bancos Nombre
        Case "PRESENTACIONES"
             Consulta_Cat_Presentaciones Nombre
        Case "CATEGORIAS"
             Consulta_Cat_Categorias Nombre
        Case "MARCAS"
             Consulta_Cat_Marcas Nombre
        Case "LABORATORIOS"
             Consulta_Cat_Laboratorios Nombre
        Case "ALMACENES"
             Consulta_Cat_Almacenes Nombre
        Case "PRODUCTOS_TIPO"
             Consulta_Cat_Productos_Tipo Nombre
        Case "IMPUESTOS"
             Consulta_Cat_Impuestos Nombre
        Case "SUSTANCIA_ACTIVA"
             Consulta_Cat_Sustancia_Activa Nombre
    End Select
End Sub

Private Sub Btn_Eliminar_Click()
Set Conectar_Ayudante = New Ayudante

On Error GoTo handler
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion) = vbYes Then
        Conexion_Base.BeginTrans
            Select Case Catalogo
                Case "ROLES":
                    If Trim(Txt_Rol_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Apl_Cat_Accesos", "Rol_ID", Trim(Txt_Rol_ID.text)) = True Then
                            If Conectar_Ayudante.Elimina_Catalogo("Apl_Cat_Roles", "Rol_ID", Trim(Txt_Rol_ID.text)) = True Then
                                If Grid_Roles.Rows = 2 Then
                                    Grid_Roles.Rows = 0
                                Else
                                    Grid_Roles.RemoveItem Grid_Roles.RowSel
                                End If 'Grid_Roles
                                Call Conectar_Ayudante.Limpiar_Textos(Me)
                                MsgBox "Rol eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                            End If 'Rol
                        End If 'Acceso
                    Else
                        MsgBox "Seleccione un rol para poder eliminar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                Case "USUARIOS": 'Corresponde al catálogo de usuarios
                    If Trim(Txt_Usuario_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Apl_Cat_Usuarios", "Usuario_ID", Txt_Usuario_ID.text) = True Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Usuarios.Rows = 2 Then
                                Grid_Usuarios.Rows = 0
                            Else
                                Grid_Usuarios.RemoveItem Grid_Usuarios.RowSel
                            End If 'Grid_Usuarios
                            MsgBox "Usuario Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If 'Elimina usuario
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If 'Txt_Usuario_ID.text
                    
                Case "CLASIFICACION_PROVEEDORES": 'Corresponde al catálogo Clasificacion de Proveedores
                    If Trim(Txt_Clasificacion_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Clasificacion_Proveedores", "Clasificacion_ID", Txt_Clasificacion_ID.text) = True Then
                            If Grid_Clasificaciones.Rows = 2 Then
                                Grid_Clasificaciones.Rows = 0
                            Else
                                Grid_Clasificaciones.RemoveItem Grid_Clasificaciones.RowSel
                            End If
                            MsgBox "Clasificación Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                       
                 Case "CLASIFICACION_CLIENTES": 'Corresponde al catálogo Clasificacion de Clientes
                    If Trim(Txt_Clasificacion_Cliente.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Clientes_Clasificacion", "Clasificacion_ID", Txt_Clasificacion_Cliente.text) = True Then
                            If Grid_Clasificacion_Clientes.Rows = 2 Then
                                Grid_Clasificacion_Clientes.Rows = 0
                            Else
                                Grid_Clasificacion_Clientes.RemoveItem Grid_Clasificacion_Clientes.RowSel
                            End If
                            MsgBox "Clasificación Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                      
                 Case "BANCOS": 'Corresponde al catálogo Bancos
                    If Trim(Txt_Banco_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Bancos", "Banco_ID", Format(Txt_Banco_ID.text, "00000")) = True Then
                            If Grid_Cat_Bancos.Rows = 2 Then
                                Grid_Cat_Bancos.Rows = 0
                            Else
                                Grid_Cat_Bancos.RemoveItem Grid_Cat_Bancos.RowSel
                            End If
                            MsgBox "Banco Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                  Case "PRESENTACIONES": 'Corresponde al catálogo Presentaciones
                    If Trim(Txt_Presentacion_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Presentaciones", "Presentacion_ID", Txt_Presentacion_ID.text) = True Then
                            If Grid_Presentaciones.Rows = 2 Then
                                Grid_Presentaciones.Rows = 0
                            Else
                                Grid_Presentaciones.RemoveItem Grid_Presentaciones.RowSel
                            End If
                            MsgBox "Presentación Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                  Case "CATEGORIAS": 'Corresponde al catálogo CATEGORIAS
                    If Trim(Txt_Catgoria_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Categorias", "Categoria_ID", Txt_Catgoria_ID.text) = True Then
                            If Grid_Cat_Categorias.Rows = 2 Then
                                Grid_Cat_Categorias.Rows = 0
                            Else
                                Grid_Cat_Categorias.RemoveItem Grid_Cat_Categorias.RowSel
                            End If
                            MsgBox "Categoria Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                  Case "MARCAS": 'Corresponde al catálogo marcas
                    If Trim(Txt_Marca_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Marcas", "Marca_ID", Txt_Marca_ID.text) = True Then
                            If Grid_Cat_Marcas.Rows = 2 Then
                                Grid_Cat_Marcas.Rows = 0
                            Else
                                Grid_Cat_Marcas.RemoveItem Grid_Cat_Marcas.RowSel
                            End If
                            MsgBox "Marca Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                  Case "LABORATORIOS": 'Corresponde al catálogo LABORATORIOS
                    If Trim(Txt_Laborartorio_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Laboratorios", "Laboratorio_ID", Txt_Laborartorio_ID.text) = True Then
                            If Grid_Cat_Laboratorios.Rows = 2 Then
                                Grid_Cat_Laboratorios.Rows = 0
                            Else
                                Grid_Cat_Laboratorios.RemoveItem Grid_Cat_Laboratorios.RowSel
                            End If
                            MsgBox "Laboratorio Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                  Case "ALMACENES": 'Corresponde al catálogo ALMACENES
                    If Trim(Txt_Almacen_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Almacenes", "Almacen_ID", Txt_Almacen_ID.text) = True Then
                            If Grid_Cat_Almacenes.Rows = 2 Then
                                Grid_Cat_Almacenes.Rows = 0
                            Else
                                Grid_Cat_Almacenes.RemoveItem Grid_Cat_Almacenes.RowSel
                            End If
                            MsgBox "Almacen Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                Case "PRODUCTOS_TIPO": 'Corresponde al catálogo Productos_Tipo
                    If Trim(Txt_Tipo_ID.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Productos_Tipo", "Tipo_ID", Txt_Tipo_ID.text) = True Then
                            If Grid_Cat_Productos_Tipo.Rows = 2 Then
                                Grid_Cat_Productos_Tipo.Rows = 0
                            Else
                                Grid_Cat_Productos_Tipo.RemoveItem Grid_Cat_Productos_Tipo.RowSel
                            End If
                            MsgBox "Tipo Producto Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                Case "IMPUESTOS": 'Corresponde al catálogo Impuestos
                    If Trim(Txt_ID_Cat_Impuestos.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Impuestos", "Impuesto_ID", Txt_ID_Cat_Impuestos.text) = True Then
                            If Grid_Impuestos_Cat_Impuestos.Rows = 2 Then
                                Grid_Impuestos_Cat_Impuestos.Rows = 0
                            Else
                                Grid_Impuestos_Cat_Impuestos.RemoveItem Grid_Impuestos_Cat_Impuestos.RowSel
                            End If
                            MsgBox "Impuesto Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                Case "SUSTANCIA_ACTIVA": 'Corresponde al catálogo de sustancia activa
                    If Trim(Txt_ID_Cat_Sustancia_Activa.text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Sustancia_Activa", "Sustancia_Activa_ID", Txt_ID_Cat_Sustancia_Activa.text) = True Then
                            If Grid_Cat_Sustancia_Activa.Rows = 2 Then
                                Grid_Cat_Sustancia_Activa.Rows = 0
                            Else
                                Grid_Cat_Sustancia_Activa.RemoveItem Grid_Cat_Sustancia_Activa.RowSel
                            End If
                            MsgBox "Sustancia activa Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    End If
                    
                    
            End Select
            Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia los textos de la forma
        Conexion_Base.CommitTrans
    End If
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
    Select Case Catalogo
        Case "ROLES":
            If Trim(Txt_Rol_ID.text) <> "" Then
                Btn_Acceso_Seguridad.Visible = False
                Fra_Acceso_Sistema_Rol.Visible = True
                Fra_Generales_Roles.Enabled = True
                Fra_Roles_Sistema.Visible = False
                Txt_Nombre_Rol.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Debe seleccionar un rol para poder modificar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
        Case "USUARIOS": 'Catalogo de usuarios
            'Revisa que exista un registro a modificar
            If Trim(Txt_Usuario_ID.text) <> "" Then
                Fra_Generales_Usuarios.Enabled = True
                Fra_Usuarios.Enabled = False
                Txt_Nombre_Usuario.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un usuario para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "CLASIFICACION_PROVEEDORES": 'Catalogo de Clasificacion de Proveedores
            'Revisa que exista un registro a modificar
            If Trim(Txt_Clasificacion_ID.text) <> "" Then
                Fra_Generales_Clasificacion_Proveedores.Enabled = True
                Fra_Detalles_Clasificacion_Proveedores.Enabled = False
                Txt_Nombre_Clasificacion_Proveedor.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una Clasificación para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "CLASIFICACION_CLIENTES": 'Catalogo de Clasificacion de Clientes
            If Trim(Txt_Clasificacion_Cliente.text) <> "" Then
                Fra_Clasificacion_Clientes.Enabled = True
                Fra_Clasificacion_Clientes_Detalles.Enabled = False
                Txt_Nombre_Clasificacion_Cliente.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una Clasificación para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "BANCOS": 'Catalogo de BANCOS
            If Trim(Txt_Banco_ID.text) <> "" Then
                Fra_Generales_Bancos.Enabled = True
                Fra_Bancos_Detalles.Enabled = False
                Txt_Nombre_Banco.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un banco para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "PRESENTACIONES": 'Catalogo de PRESENTACIONES
            If Trim(Txt_Presentacion_ID.text) <> "" Then
                Fra_Generales_Presentaciones.Enabled = True
                Fra_Detalles_Presentaciones.Enabled = False
                Txt_Nombre_Presentaciones.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una presentación para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "CATEGORIAS": 'Catalogo de CATEGORIAS
            If Trim(Txt_Catgoria_ID.text) <> "" Then
                Fra_Generales_Categorias.Enabled = True
                Fra_Detalles_Categorias.Enabled = False
                Txt_Nombre_Categorias.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una categoría para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "MARCAS": 'Catalogo de MARCAS
            If Trim(Txt_Marca_ID.text) <> "" Then
                Fra_Generales_Marcas.Enabled = True
                Fra_Detalles_Marcas.Enabled = False
                Txt_Nombre_Marca.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una marca para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
            
        Case "LABORATORIOS": 'Catalogo de LABORATORIOS
            If Trim(Txt_Laborartorio_ID.text) <> "" Then
                Fra_Generales_Laboratorio.Enabled = True
                Fra_Detalles_Laboratorios.Enabled = False
                Txt_Nombre_Laboratorios.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un laboratorio para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "ALMACENES": 'Catalogo de ALMACENES
            If Trim(Txt_Almacen_ID.text) <> "" Then
                Fra_Cat_Almacenes.Enabled = True
                Fra_Cat_Almacenes_Detalles.Enabled = False
                Txt_Nombre_Cat_Almacenes.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un laboratorio para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "PRODUCTOS_TIPO": 'Catalogo de PRODUCTOS_TIPO
            If Trim(Txt_Tipo_ID.text) <> "" Then
                Fra_Cat_Productos_Tipo.Enabled = True
                Fra_Cat_Produstos_Tipo_Detalles.Enabled = False
                Txt_Nombre_Cat_Productos_Tipo.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un Tipo producto para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "IMPUESTOS": 'Catalogo de IMPUESTOS
            If Trim(Txt_ID_Cat_Impuestos.text) <> "" Then
                Fra_Generales_Cta_Impuestos.Enabled = True
                Fra_Detalles_Cat_Impuestos.Enabled = False
                Txt_Impuesto_Cat_Impuestos.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione un impuesto para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
        Case "SUSTANCIA_ACTIVA": 'Catalogo de sustancia activa
            If Trim(Txt_ID_Cat_Sustancia_Activa.text) <> "" Then
                Fra_Cat_Sustancia_Activa.Enabled = True
                Fra_Detalles_Cat_Sustancia_Activa.Enabled = False
                Txt_Nombre_Cat_Sustancia_Activa.SetFocus
                SendKeys "{Home}+{End}"
            Else
                MsgBox "Seleccione una sustancia activa para poder modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Exit Sub
            End If
            
    End Select
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = False
    Btn_Nuevo.Enabled = False
    Btn_Consultar.Enabled = False
    Btn_Salir.Caption = "Cancelar"
Else
    Select Case Catalogo
        Case "ROLES":
            If Trim(Txt_Nombre_Rol.text) <> "" Then
                If Consulta_Nombre_Rol = False Then
                    Modifica_Rol 'Modifica el registro del rol
                Else
                    MsgBox "El nombre del rol ya esta dado de alta" & Chr(13) & Chr(13) & _
                           "favor de introducirlo nuevamente", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                    Txt_Nombre_Rol.SetFocus
                    SendKeys "{Home}+{End}"
                End If
            Else
                MsgBox "Proporcione el nombre del rol", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Txt_Nombre_Rol.SetFocus
                SendKeys "{Home}+{End}"
            End If
        Case "USUARIOS":
                'Si los datos requeridos no estan vacios entonces modifica los valores
                'del registro
                If Trim(Txt_Nombre_Usuario.text) <> "" And Trim(Txt_Contraseña.text) <> "" _
                    And Cmb_Roles.ListIndex > -1 And Trim(Txt_Login.text) <> "" And _
                    Trim(Txt_Confirmar_Contraseña.text) <> "" And Cmb_Estatus_Usuario.ListIndex > -1 Then
                    If Not Trim(Txt_Contraseña.text) = Trim(Txt_Confirmar_Contraseña.text) Then
                    'Valida que la contraseña ingresada sea correcta
                        MsgBox "La confirmacion de contraseña no es correcta. " + vbCrLf + _
                                "Favor de ingresar nuevamente la contraseña", vbInformation + vbOKOnly, "Contraseña"
                        Txt_Contraseña = ""
                        Txt_Confirmar_Contraseña = ""
                        Txt_Contraseña.SetFocus
                    Else
                         If Valida_Login_Password_Usuario("Login", Trim(Txt_Login.text), Trim(Txt_Usuario_ID.text)) = False Then
                                If Valida_Login_Password_Usuario("Password", Trim(Txt_Contraseña.text), Trim(Txt_Usuario_ID.text)) = False Then
                                    'Valida que la contraseña y login tenga por lo menos 6 caracteres
                                    If Len(Txt_Login.text) >= 6 Then
                                        If Len(Txt_Contraseña.text) >= 6 Then
                                            If Conectar_Ayudante.Es_Alfanumerico(Txt_Login.text) = True And Conectar_Ayudante.Es_Alfanumerico(Txt_Contraseña.text) = True Then
                                                Modifica_Usuarios 'Modifica el registro del usuario que fue seleccionado por el usuario
                                            Else
                                                MsgBox "Verifique su login y contraseña" & Chr(13) & Chr(13) & _
                                                       "ya que deben de contener números y letras", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                                Txt_Login.SetFocus
                                                SendKeys "{Home}+{End}"
                                            End If
                                        Else
                                            MsgBox "Su contraseña no es valida debe estar conformado" + vbCrLf + _
                                                   "por lo menos de 6 caracteres", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                            Txt_Contraseña.SetFocus
                                            SendKeys "{Home}+{End}"
                                        End If
                                    Else
                                        MsgBox "Su login no es valido debe estar conformado" + vbCrLf + _
                                               "por lo menos de 6 caracteres", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                        Txt_Login.SetFocus
                                        SendKeys "{Home}+{End}"
                                    End If
                                Else
                                    MsgBox "Ese password ya existe, verifíquelo por favor", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                    Txt_Contraseña.SetFocus
                                End If
                        Else
                            MsgBox "Ese login ya existe, verifíquelo por favor", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                            Txt_Login.SetFocus
                        End If
                    End If
                'Si no manda un mensaje al usuario
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
                
        Case "CLASIFICACION_PROVEEDORES":
            If Trim(Txt_Nombre_Clasificacion_Proveedor.text) <> "" And Cmb_Estatus_Clasificacion_Proveedor.text <> "" Then
                Modifica_Clasificacion_Proveedor 'Modifica la clasificacion del proveedor
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "CLASIFICACION_CLIENTES":
            If Trim(Txt_Nombre_Clasificacion_Cliente.text) <> "" And Cmb_Estatus_Clasificacion_Cliente.text <> "" Then
                Modifica_Clasificacion_Cliente 'Modifica la clasificacion del proveedor
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "BANCOS":
            If Trim(Txt_Nombre_Banco.text) <> "" And Cmb_Estatus_Banco.text <> "" And Ttx_Sucursal.text <> "" And Txt_RFC_Banco.text <> "" Then
                Modifica_Bancos 'Modifica el Banco
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "PRESENTACIONES":
            If Trim(Txt_Nombre_Presentaciones.text) <> "" And Cmb_Estaus_Presentaciones.text <> "" Then
                Modifica_Cat_Presentaciones 'Modifica la Presentacion
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "CATEGORIAS":
            If Trim(Txt_Nombre_Categorias.text) <> "" And Cmb_Estatus_Categoria.text <> "" Then
                Modifica_Cat_Categorias 'Modifica las Categorias
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        Case "MARCAS":
            If Trim(Txt_Marca_ID.text) <> "" And Cmb_Estatus_Marca.text <> "" Then
                Modifica_Cat_Marcas 'Modifica las MARCAS
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "LABORATORIOS":
            If Trim(Txt_Laborartorio_ID.text) <> "" And Cmb_Estatus_Laboratorios.text <> "" Then
                Modifica_Cat_Laboratorios 'Modifica LOS LABORATORIOS
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        Case "ALMACENES":
            If Trim(Txt_Almacen_ID.text) <> "" And Cmb_Estatus_Cat_Almacenes.text <> "" Then
                Modifica_Cat_Almacenes 'Modifica LOS almacenes
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        Case "PRODUCTOS_TIPO":
            If Trim(Txt_Tipo_ID.text) <> "" And Cmb_Estatus_Cat_Productos_Tipo.text <> "" Then
                Modifica_Cat_Productos_Tipo 'Modifica Los Tipo Producto
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "IMPUESTOS":
            If Trim(Txt_ID_Cat_Impuestos.text) <> "" And Cmb_Estatus_Cat_Impuestos.text <> "" Then
                Modifica_Cat_Impuestos 'Modifica Los Impuestos
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "SUSTANCIA_ACTIVA":
            If Trim(Txt_ID_Cat_Sustancia_Activa.text) <> "" And Cmb_Cat_Sustancia_Activa.text <> "" Then
                Modifica_Sustancia_Activa
            Else
                MsgBox "Faltan datos para Modificar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
    End Select
End If
End Sub

Private Sub Btn_Nuevo_Click()
Dim Catacter As String 'Indica el caractere que se desea comparar
    
If Btn_Nuevo.Caption = "Nuevo" Then
    Btn_Nuevo.Caption = "Dar de Alta"
    Btn_Modificar.Enabled = False
    Btn_Eliminar.Enabled = False
    Btn_Consultar.Enabled = False
    Btn_Salir.Caption = "Cancelar"
    Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales) 'Limpia las cajas de texto
    'Muestra el picture del catalogo seleccionado
    Select Case Catalogo
    Case "ROLES": 'Corresponde al catálgo de roles
        Txt_Rol_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Apl_Cat_Roles", "Rol_ID"), "00000")
        Btn_Acceso_Seguridad.Visible = False
        Fra_Acceso_Sistema_Rol.Visible = True
        Fra_Generales_Roles.Enabled = True
        Fra_Roles_Sistema.Visible = False
        Consulta_Configuracion 'Consulta los menus y submenus que se encuentren en el sistema
        Txt_Nombre_Rol.SetFocus
        
    Case "USUARIOS": 'Corresponde al catálogo de usuarios
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Usuario_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Apl_Cat_Usuarios", "Usuario_ID"), "00000")
            DTP_Fecha_Caducar_Usuario.Value = Now
            Fra_Generales_Usuarios.Enabled = True
            Fra_Usuarios.Enabled = False
            Cmb_Estatus_Usuario.ListIndex = 0
            Cmb_Estatus_Usuario.Enabled = False
            Txt_Nombre_Usuario.SetFocus
            
    Case "CLASIFICACION_PROVEEDORES": 'Corresponde al catálogo de Clasificacion de proveedores
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Clasificacion_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Clasificacion_Proveedores", "Clasificacion_ID"), "00000")
            Fra_Generales_Clasificacion_Proveedores.Enabled = True
            Fra_Detalles_Clasificacion_Proveedores.Enabled = False
            Cmb_Estatus_Clasificacion_Proveedor.ListIndex = 0
            Cmb_Estatus_Clasificacion_Proveedor.Enabled = False
            Txt_Nombre_Clasificacion_Proveedor.SetFocus
            
    Case "CLASIFICACION_CLIENTES": 'Corresponde al catálogo de Clasificacion de clientes
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Clasificacion_Cliente.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Clientes_Clasificacion", "Clasificacion_ID"), "00000")
            Fra_Clasificacion_Clientes.Enabled = True
            Fra_Clasificacion_Clientes_Detalles.Enabled = False
            Cmb_Estatus_Clasificacion_Cliente.ListIndex = 0
            Cmb_Estatus_Clasificacion_Cliente.Enabled = False
            Txt_Nombre_Clasificacion_Cliente.SetFocus
            
    Case "BANCOS": 'Corresponde al catálogo de BANCOS
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Banco_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Bancos", "Banco_ID"), "00")
            Fra_Generales_Bancos.Enabled = True
            Fra_Bancos_Detalles.Enabled = False
            Cmb_Estatus_Banco.ListIndex = 0
            Cmb_Estatus_Banco.Enabled = False
            Txt_Nombre_Banco.SetFocus
            
            
    Case "PRESENTACIONES": 'Corresponde al catálogo de PRESENTACIONES
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Presentacion_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Presentaciones", "Presentacion_ID"), "00000")
            Fra_Generales_Presentaciones.Enabled = True
            Fra_Detalles_Presentaciones.Enabled = False
            Cmb_Estaus_Presentaciones.ListIndex = 0
            Cmb_Estaus_Presentaciones.Enabled = False
            Txt_Nombre_Presentaciones.SetFocus
            
    Case "CATEGORIAS": 'Corresponde al catálogo de CATEGORIAS
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Catgoria_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Categorias", "Categoria_ID"), "00000")
            Fra_Generales_Categorias.Enabled = True
            Fra_Detalles_Categorias.Enabled = False
            Cmb_Estatus_Categoria.ListIndex = 0
            Cmb_Estatus_Categoria.Enabled = False
            Txt_Nombre_Categorias.SetFocus
            
    Case "MARCAS": 'Corresponde al catálogo de MARCAS
            'Llama al último registro de la tabla y asigna el siguiente
            Txt_Marca_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Marcas", "Marca_ID"), "00000")
            Fra_Generales_Marcas.Enabled = True
            Fra_Detalles_Marcas.Enabled = False
            Cmb_Estatus_Marca.ListIndex = 0
            Cmb_Estatus_Marca.Enabled = False
            Txt_Nombre_Marca.SetFocus
        
    Case "LABORATORIOS": 'Corresponde al catálogo de LABORATORIOS
       'Llama al último registro de la tabla y asigna el siguiente
       Txt_Laborartorio_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Laboratorios", "Laboratorio_ID"), "00000")
       Fra_Generales_Laboratorio.Enabled = True
       Fra_Detalles_Laboratorios.Enabled = False
       Cmb_Estatus_Laboratorios.ListIndex = 0
       Cmb_Estatus_Laboratorios.Enabled = False
       Txt_Nombre_Laboratorios.SetFocus
       
    Case "ALMACENES": 'Corresponde al catálogo de ALMACENES
       'Llama al último registro de la tabla y asigna el siguiente
       Txt_Almacen_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Almacenes", "Almacen_ID"), "00000")
       Fra_Cat_Almacenes.Enabled = True
       Fra_Cat_Almacenes_Detalles.Enabled = False
       Cmb_Estatus_Cat_Almacenes.ListIndex = 0
       Cmb_Estatus_Cat_Almacenes.Enabled = False
       Txt_Nombre_Cat_Almacenes.SetFocus
       
    Case "PRODUCTOS_TIPO": 'Corresponde al catálogo de Productos_Tipo
       'Llama al último registro de la tabla y asigna el siguiente
       Txt_Tipo_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Tipo", "Tipo_ID"), "00000")
       Fra_Cat_Productos_Tipo.Enabled = True
       Fra_Cat_Produstos_Tipo_Detalles.Enabled = False
       Cmb_Estatus_Cat_Productos_Tipo.ListIndex = 0
       Cmb_Estatus_Cat_Productos_Tipo.Enabled = False
       Txt_Nombre_Cat_Productos_Tipo.SetFocus
            
    Case "IMPUESTOS": 'Corresponde al catálogo de Cat_Impuestos
       'Llama al último registro de la tabla y asigna el siguiente
       Txt_ID_Cat_Impuestos.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Impuestos", "Impuesto_ID"), "00000")
       Fra_Generales_Cta_Impuestos.Enabled = True
       Fra_Detalles_Cat_Impuestos.Enabled = False
       Cmb_Estatus_Cat_Impuestos.ListIndex = 0
       Cmb_Estatus_Cat_Impuestos.Enabled = False
       Txt_Impuesto_Cat_Impuestos.SetFocus
       
    Case "SUSTANCIA_ACTIVA": 'Corresponde al catálogo de Cat_Sustancia_Activa
       'Llama al último registro de la tabla y asigna el siguiente
       Txt_ID_Cat_Sustancia_Activa.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Sustancia_Activa", "Sustancia_Activa_ID"), "00000")
       Fra_Cat_Sustancia_Activa.Enabled = True
       Fra_Detalles_Cat_Sustancia_Activa.Enabled = False
       Cmb_Cat_Sustancia_Activa.ListIndex = 0
       Cmb_Cat_Sustancia_Activa.Enabled = False
       Txt_Nombre_Cat_Sustancia_Activa.SetFocus
            
    End Select
Else
    Select Case Catalogo
        Case "ROLES": 'Corresponde al catálogo de roles
            If Trim(Txt_Nombre_Rol.text) <> "" Then
                If Consulta_Nombre_Rol = False Then
                    Alta_Rol
                Else
                    MsgBox "Proporcione otro nombre para el rol" & Chr(13) & Chr(13) & _
                           "porque ya se encuentra asignado a otro rol", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                    Txt_Nombre_Rol.SetFocus
                End If
            Else
                MsgBox "Proporcione el nombre del rol", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                Txt_Nombre_Rol.SetFocus
            End If
        Case "USUARIOS": 'Corresponde el catálago de usuarios
            'Alta de usuarios sólo si estan llenos los campos obligatorios
            If Trim(Txt_Nombre_Usuario.text) <> "" And Trim(Txt_Contraseña.text) <> "" _
            And Cmb_Roles.ListIndex > -1 And Trim(Txt_Login.text) <> "" _
            And Trim(Txt_Confirmar_Contraseña) <> "" Then
                'Valida que las contraseña ingresada sea correcta
                If Not Txt_Contraseña = Txt_Confirmar_Contraseña Then
                    MsgBox "La confirmacion de contraseña no es correcta. " + vbCrLf + _
                        "Favor de ingresar nuevamente la contraseña", vbInformation + vbOKOnly, "Contraseña"
                    Txt_Contraseña = ""
                    Txt_Confirmar_Contraseña = ""
                    Txt_Contraseña.SetFocus
                Else
                     If Valida_Login_Password_Usuario("Login", Trim(Txt_Login.text), Trim(Txt_Usuario_ID.text)) = False Then
                            If Valida_Login_Password_Usuario("Password", Trim(Txt_Contraseña.text), Trim(Txt_Usuario_ID.text)) = False Then
                                'Valida que el login y password tengan por lo manos 6 caracteres
                                'para poder der dar de alta
                                If Len(Txt_Login.text) >= 6 Then
                                    If Len(Txt_Contraseña.text) >= 6 Then
                                        If Conectar_Ayudante.Es_Alfanumerico(Txt_Login.text) = True And Conectar_Ayudante.Es_Alfanumerico(Txt_Contraseña.text) = True Then
                                            Alta_Usuarios 'Da de alta los datos del usuario
                                        Else
                                            MsgBox "Verifique su login y contraseña" & Chr(13) & Chr(13) & _
                                                   "ya que deben de contener números y letras", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                            Txt_Login.SetFocus
                                            SendKeys "{Home}+{End}"
                                        End If
                                    Else
                                        MsgBox "Su contraseña no es valida debe estar conformado" + vbCrLf + _
                                               "por lo menos de 6 caracteres", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                        Txt_Contraseña.SetFocus
                                        SendKeys "{Home}+{End}"
                                    End If
                                Else
                                    MsgBox "Su login no es valido debe estar conformado" + vbCrLf + _
                                           "por lo menos de 6 caracteres", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                    Txt_Login.SetFocus
                                    SendKeys "{Home}+{End}"
                                End If
                            Else
                                MsgBox "Ese password ya existe, verifíquelo por favor", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                                Txt_Contraseña.SetFocus
                            End If
                        Else
                            MsgBox "Ese login ya existe, verifíquelo por favor", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                            Txt_Login.SetFocus
                        End If
                End If
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
        Case "CLASIFICACION_PROVEEDORES": 'Corresponde el catálago de Clasificacion de proveedores
            'Alta de usuarios sólo si estan llenos los campos obligatorios
            If Trim(Txt_Nombre_Clasificacion_Proveedor.text) <> "" And Trim(Cmb_Estatus_Clasificacion_Proveedor.text) <> "" Then
                 Alta_Clasificacion_Proveedor 'Da de alta la clasificacion del proveedor
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
         Case "CLASIFICACION_CLIENTES": 'Corresponde el catálago de Clasificacion de Clientes
            If Trim(Txt_Nombre_Clasificacion_Cliente.text) <> "" And Trim(Cmb_Estatus_Clasificacion_Cliente.text) <> "" Then
                 Alta_Clasificacion_Cliente 'Da de alta la clasificacion del cliente
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "BANCOS": 'Corresponde el catálago de Bancos
            If Trim(Txt_Nombre_Banco.text) <> "" And Trim(Cmb_Estatus_Banco.text) <> "" And Trim(Ttx_Sucursal.text) <> "" And Trim(Txt_RFC_Banco) <> "" Then
                 Alta_Bancos 'Da de alta el Banco
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "PRESENTACIONES": 'Corresponde el catálago de Presentaciones
            If Trim(Txt_Nombre_Presentaciones.text) <> "" And Trim(Cmb_Estaus_Presentaciones.text) <> "" Then
                 Alta_Cat_Presentaciones 'Da de alta la presentación
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "CATEGORIAS": 'Corresponde el catálago de Categorias
            If Trim(Txt_Nombre_Categorias.text) <> "" And Trim(Cmb_Estatus_Categoria.text) <> "" Then
                 Alta_Cat_Categorias 'Da de alta las categorias
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "MARCAS": 'Corresponde el catálago de MARCAS
            If Trim(Txt_Nombre_Marca.text) <> "" And Trim(Cmb_Estatus_Marca.text) <> "" Then
                 Alta_Cat_Marcas 'Da de alta las Marcas
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "LABORATORIOS": 'Corresponde el catálago de LABORATORIOS
            If Trim(Txt_Nombre_Laboratorios.text) <> "" And Trim(Cmb_Estatus_Laboratorios.text) <> "" Then
                 Alta_Cat_Laboratorios 'Da de ALTA LOS LABORATORIOS
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "ALMACENES": 'Corresponde el catálago de ALMACENES
            If Trim(Txt_Nombre_Cat_Almacenes.text) <> "" And Trim(Cmb_Estatus_Cat_Almacenes.text) <> "" Then
                 Alta_Cat_Almacenes 'Da de ALTA LOS ALMACENES
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "PRODUCTOS_TIPO": 'Corresponde el catálago de PRODUCTOS_TIPO
            If Trim(Txt_Nombre_Cat_Productos_Tipo.text) <> "" And Trim(Cmb_Estatus_Cat_Productos_Tipo.text) <> "" Then
                 Alta_Cat_Productos_Tipo 'Da de ALTA LOS PRODUCTOS TIPO
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "IMPUESTOS": 'Corresponde el catálago de Cat_Impuestos
            If Trim(Txt_Impuesto_Cat_Impuestos.text) <> "" And Trim(Cmb_Estatus_Cat_Impuestos.text) <> "" Then
                 Alta_Cat_Impuestos 'Da de ALTA LOS IMPUESTOS
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
          Case "SUSTANCIA_ACTIVA": 'Corresponde el catálago de Cat_Sustancia_Activa
            If Trim(Txt_Nombre_Cat_Sustancia_Activa.text) <> "" And Trim(Cmb_Cat_Sustancia_Activa.text) <> "" Then
                 Alta_Cat_Sustancia_Activa 'Da de ALTA la sustancia activa
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
            
    End Select
End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Configuracion
    'DESCRIPCIÓN: Consulta por medio de la Clave de usuario los Menus del sistema
    'PARÁMETROS :
    'CREO       : Jorge Razo
    'FECHA_CREO :
    'MODIFICO          : Yazmin Delgado Gómez
    'FECHA_MODIFICO    : 28-Mayo-2007
    'CAUSA_MODIFICACIÓN: Porque se modifico la forma de accesar al sistema
'*******************************************************************************
Private Sub Consulta_Configuracion()
Dim Ctl As Control               'Indica que control es el que se esta consultando en el sistema
Dim Contador_Columnas As Integer 'Indica que columna del grid se esta consultando

    Grid_Accesos_Seguridad.Rows = 0
    'Pone en el encabezado del grid los nombres de la columnas
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "Menu" & Chr(9) & "Submenu" & _
    Chr(9) & "Nombre" & Chr(9) & "Tipo" & Chr(9) & "Habilitar" & _
    Chr(9) & "Alta" & Chr(9) & "Cambio" & Chr(9) & "Eliminar" & Chr(9) & "Consultar"
    'Agrega los menus y submenus que se encuentran asignados en el sistema
    For Each Ctl In MDIFrm_Apl_Principal.Controls
        If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Or UCase(Mid(Ctl.Name, 1, 7)) = UCase("SUBMENU") Then
            If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                Grid_Accesos_Seguridad.AddItem "-" & Chr(9) & UCase(Ctl.Caption) _
                & Chr(9) & "" & Chr(9) & Ctl.Name & Chr(9) & "Encabezado" _
                & Chr(9) & "S" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                'Agrega el color gris a la fila que tiene el encabezado
                For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                    Grid_Accesos_Seguridad.Col = Contador_Columnas
                    Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                Next Contador_Columnas
            Else
                Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "" & Chr(9) & UCase(Ctl.Caption) & _
                Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & Chr(9) & "S" & _
                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "S"
            End If
        End If
    Next Ctl
    'Configura el tamaño de las columnas del grid_accesos_seguridad
    If Grid_Accesos_Seguridad.Rows > 1 Then
        Grid_Accesos_Seguridad.FixedRows = 1
        Grid_Accesos_Seguridad.ColWidth(0) = 200 '-
        Grid_Accesos_Seguridad.ColWidth(1) = 1250 'Menu
        Grid_Accesos_Seguridad.ColWidth(2) = 2500 'SubMenu
        Grid_Accesos_Seguridad.ColWidth(3) = 0    'Nombre Menu/Submenu
        Grid_Accesos_Seguridad.ColWidth(4) = 0    'Tipo
        Grid_Accesos_Seguridad.ColWidth(5) = 900 'Habilitar
        Grid_Accesos_Seguridad.ColAlignment(5) = 3
        Grid_Accesos_Seguridad.ColWidth(6) = 600  'Alta
        Grid_Accesos_Seguridad.ColAlignment(6) = 3
        Grid_Accesos_Seguridad.ColWidth(7) = 650  'Cambio
        Grid_Accesos_Seguridad.ColAlignment(7) = 3
        Grid_Accesos_Seguridad.ColWidth(8) = 650  'Eliminar
        Grid_Accesos_Seguridad.ColAlignment(8) = 3
        Grid_Accesos_Seguridad.ColWidth(9) = 750  'Consultar
        Grid_Accesos_Seguridad.ColAlignment(9) = 3
        Collapsing = True
        Call Collapse_Grid
        Collapsing = False
    End If
End Sub
'Cierra la forma
Private Sub Btn_Salir_Click()
If Btn_Salir.Caption = "Salir" Then
    Unload Me
Else
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Consultar.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Select Case Catalogo
        Case "ROLES":
            Btn_Acceso_Seguridad.Visible = True
            Fra_Acceso_Sistema_Rol.Visible = False
            Fra_Roles_Sistema.Visible = True
            Fra_Generales_Roles.Enabled = False
            Grid_Accesos_Seguridad.Rows = 0
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Roles", Frm_Cat_Generales)
        Case "USUARIOS":
            Fra_Generales_Usuarios.Enabled = False
            Fra_Usuarios.Enabled = True
            Cmb_Estatus_Usuario.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Usuarios", Frm_Cat_Generales)
            
        Case "CLASIFICACION_PROVEEDORES":
            Fra_Generales_Clasificacion_Proveedores.Enabled = False
            Fra_Detalles_Clasificacion_Proveedores.Enabled = True
            Cmb_Estatus_Clasificacion_Proveedor.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Proveedor", Frm_Cat_Generales)
            
        Case "CLASIFICACION_CLIENTES":
            Fra_Clasificacion_Clientes.Enabled = False
            Fra_Clasificacion_Clientes_Detalles.Enabled = True
            Cmb_Estatus_Clasificacion_Cliente.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Clientes", Frm_Cat_Generales)
            
        Case "BANCOS":
            Fra_Generales_Bancos.Enabled = False
            Fra_Bancos_Detalles.Enabled = True
            Cmb_Estatus_Banco.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Bancos", Frm_Cat_Generales)
            
        Case "PRESENTACION":
            Fra_Generales_Presentaciones.Enabled = False
            Fra_Detalles_Presentaciones.Enabled = True
            Cmb_Estaus_Presentaciones.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Presentaciones", Frm_Cat_Generales)
            
        Case "MARCAS":
            Fra_Generales_Marcas.Enabled = False
            Fra_Detalles_Marcas.Enabled = True
            Cmb_Estatus_Marca.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Marcas", Frm_Cat_Generales)
            
        Case "LABORATORIOS":
            Fra_Generales_Laboratorio.Enabled = False
            Fra_Detalles_Laboratorios.Enabled = True
            Cmb_Estatus_Laboratorios.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Laboratorios", Frm_Cat_Generales)
            
        Case "ALMACENES":
            Fra_Cat_Almacenes.Enabled = False
            Fra_Cat_Almacenes_Detalles.Enabled = True
            Cmb_Estatus_Laboratorios.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Almacenes", Frm_Cat_Generales)
            
        Case "PRODUCTOS_TIPO":
            Fra_Cat_Productos_Tipo.Enabled = False
            Fra_Cat_Produstos_Tipo_Detalles.Enabled = True
            Cmb_Estatus_Cat_Productos_Tipo.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos_Tipo", Frm_Cat_Generales)
            
        Case "IMPUESTOS":
            Fra_Generales_Cta_Impuestos.Enabled = False
            Fra_Detalles_Cat_Impuestos.Enabled = True
            Cmb_Estatus_Cat_Impuestos.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Impuestos", Frm_Cat_Generales)
            
        Case "SUSTANCIA_ACTIVA":
            Fra_Cat_Sustancia_Activa.Enabled = False
            Fra_Detalles_Cat_Sustancia_Activa.Enabled = True
            Cmb_Cat_Sustancia_Activa.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Sustancia_Activa", Frm_Cat_Generales)
            
    End Select
End If
End Sub

Private Sub Chk_Habilitar_Menu_Submenu_Click()
Dim Fila As Integer          'Indica que columna del grid se esta consultando
Dim Contador_Fila As Integer 'Indica que fila del grid es el que se esta consultando
Dim Inicio_Fila As Integer   'Indica en donde empieza el menu del submenu que se esta consultando
Dim Bandera As Boolean       'Indica si se encuentra habilitado algun submenu del menu que se pretende ocultar

    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0) = "" Then
        If Chk_Habilitar_Menu_Submenu.Value = 0 Then
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "N"
            For Fila = Grid_Accesos_Seguridad.RowSel To 1 Step -1
                If Grid_Accesos_Seguridad.TextMatrix(Fila, 0) = "-" Then
                    Exit For
                End If
                Inicio_Fila = Fila
            Next Fila
            For Renglon = Inicio_Fila To Grid_Accesos_Seguridad.Rows - 1
                If Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "-" Or Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "+" Then
                    Exit For
                End If
            Next Renglon
            Bandera = False 'Indica que no se tiene habilitado ningun submenu
            For Fila = Inicio_Fila To Renglon - 1
                If Grid_Accesos_Seguridad.TextMatrix(Fila, 5) = "S" Then
                    Bandera = True
                End If
            Next Fila
            If Bandera = False Then
                Grid_Accesos_Seguridad.TextMatrix(Inicio_Fila - 1, 5) = "N"
                Chk_Habilitar_Menu_Submenu.Visible = False
            End If
        Else
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "N"
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "S"
        End If
    Else
        If Chk_Habilitar_Menu_Submenu.Value = 0 Then
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "N"
        Else
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S"
        End If
        For Fila = Grid_Accesos_Seguridad.RowSel + 1 To Grid_Accesos_Seguridad.Rows - 1
            If Grid_Accesos_Seguridad.TextMatrix(Fila, 0) = "" Then
                If Chk_Habilitar_Menu_Submenu.Value = 0 Then
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 5) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 6) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 7) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 8) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 9) = "N"
                Else
                    For Contador_Fila = Grid_Accesos_Seguridad.RowSel To Grid_Accesos_Seguridad.RowSel - 1
                        If Grid_Accesos_Seguridad.TextMatrix(Contador_Fila, 0) = "-" Then
                            Exit For
                        Else
                            If Grid_Accesos_Seguridad.TextMatrix(Contador_Fila, 5) = "S" Then
                                Exit Sub
                            End If
                        End If
                    Next Contador_Fila
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 5) = "S"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 6) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 7) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 8) = "N"
                    Grid_Accesos_Seguridad.TextMatrix(Fila, 9) = "S"
                End If
            Else
                Exit Sub
            End If
        Next Fila
    End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Me.Height = 7150
    Me.Width = 8500
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
    Select Case Catalogo
        Case "USUARIOS":
            Consulta_Usuarios "" 'Consulta todos los usuario que estan dados de alta
            Call Conectar_Ayudante.Llena_Combo_Item("Rol_ID, Nombre", "Apl_Cat_Roles", Cmb_Roles, 0, "")
            DTP_Fecha_Caducar_Usuario.Value = Now
        Case "ROLES":
            Consulta_Roles "" 'Consulta todos los roles que estan dados de alta
        Case "CLASIFICACION_PROVEEDORES":
            Call Consulta_Clasificacion_Proveedores("")
        Case "CLASIFICACION_CLIENTES":
            Call Consulta_Clasificacion_Clientes("")
        Case "BANCOS":
            Call Consulta_Bancos("")
            Call Conectar_Ayudante.Llena_Combo_Item("Formato_ID,Nombre", "Cfg_Formatos", Cmb_Formato, 1, "Nombre")
        Case "PRESENTACIONES":
            Call Consulta_Cat_Presentaciones("")
        Case "CATEGORIAS":
            Call Consulta_Cat_Categorias("")
        Case "MARCAS":
            Call Consulta_Cat_Marcas("")
        Case "LABORATORIOS":
            Call Consulta_Cat_Laboratorios("")
        Case "ALMACENES":
            Call Consulta_Cat_Almacenes("")
        Case "PRODUCTOS_TIPO":
            Call Consulta_Cat_Productos_Tipo("")
            
        Case "IMPUESTOS":
            Call Consulta_Cat_Impuestos("")
            
        Case "SUSTANCIA_ACTIVA":
            Call Consulta_Cat_Sustancia_Activa("")
    End Select
End Sub

Private Sub Grid_Accesos_Seguridad_Click()
Dim Renglon As Integer     'Indica que renglon se esta consulltando

If Grid_Accesos_Seguridad.Rows > 1 Then
    Chk_Habilitar_Menu_Submenu.Visible = False
    Txt_Habilitar.Visible = False
    If Grid_Accesos_Seguridad.Col <= 1 And Grid_Accesos_Seguridad.Row > 0 Then
        'And Grid_Accesos_Seguridad.TextMatrix(Renglon, 1) = "-"
        If Collapsing = False Then
            Renglon = Grid_Accesos_Seguridad.MouseRow
        Else
            Renglon = Renglon_Procesar
        End If
        If Renglon < 1 Then Exit Sub
        
        While Renglon > 0 And Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = ""
            Renglon = Renglon - 1
        Wend
        If Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "-" Then
            Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "+"
        Else
            Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "-"
        End If
        Renglon = Renglon + 1
        If Grid_Accesos_Seguridad.RowHeight(Renglon) = 0 Then
            Do While Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = ""
                Grid_Accesos_Seguridad.RowHeight(Renglon) = -1
                Renglon = Renglon + 1
                If Renglon >= Grid_Accesos_Seguridad.Rows Then Exit Do
            Loop
        Else
            Do While Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = ""
                Grid_Accesos_Seguridad.RowHeight(Renglon) = 0
                Renglon = Renglon + 1
                If Renglon >= Grid_Accesos_Seguridad.Rows Then Exit Do
            Loop
        End If
        Grid_Accesos_Seguridad.Col = 0
    End If
    If Btn_Nuevo.Caption = "Dar de Alta" Or Btn_Modificar.Caption = "Actualizar" Then
        If Trim(UCase(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 4))) <> "ENCABEZADO" Then
            Chk_Habilitar_Menu_Submenu.Visible = False
            For Fila = Grid_Accesos_Seguridad.RowSel To 1 Step -1
                If Grid_Accesos_Seguridad.TextMatrix(Fila, 0) = "-" Then
                    If Grid_Accesos_Seguridad.TextMatrix(Fila, 5) = "S" Then
                        Exit For
                    Else
                        Chk_Habilitar_Menu_Submenu.BackColor = &HFFFFFF
                        Chk_Habilitar_Menu_Submenu.Visible = False
                        Exit Sub
                    End If
                End If
            Next Fila
            Chk_Habilitar_Menu_Submenu.Visible = False
            Txt_Habilitar.Visible = False
            Select Case Grid_Accesos_Seguridad.ColSel
                Case 5:
                    Chk_Habilitar_Menu_Submenu.BackColor = &HFFFFFF
                    Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Accesos_Seguridad, Chk_Habilitar_Menu_Submenu)
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Chk_Habilitar_Menu_Submenu.Value = 1
                    Else
                        Chk_Habilitar_Menu_Submenu.Value = 0
                    End If
                Case 6:
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Accesos_Seguridad, Txt_Habilitar)
                        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "S" Then
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "N"
                        Else
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "S"
                        End If
                        Txt_Habilitar = Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6)
                    End If
                Case 7:
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Accesos_Seguridad, Txt_Habilitar)
                        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "S" Then
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "N"
                        Else
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "S"
                        End If
                        Txt_Habilitar = Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7)
                    End If
                Case 8:
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Accesos_Seguridad, Txt_Habilitar)
                        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "S" Then
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "N"
                        Else
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "S"
                        End If
                        Txt_Habilitar = Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8)
                    End If
                Case 9:
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Accesos_Seguridad, Txt_Habilitar)
                        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "S" Then
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "N"
                        Else
                            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "S"
                        End If
                        Txt_Habilitar = Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9)
                    End If
            End Select
        Else
            If Grid_Accesos_Seguridad.Rows >= 1 Then
                Chk_Habilitar_Menu_Submenu.Visible = False
                Select Case Grid_Accesos_Seguridad.ColSel
                    Case 5:
                        Chk_Habilitar_Menu_Submenu.BackColor = vbButtonFace
                        Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Accesos_Seguridad, Chk_Habilitar_Menu_Submenu)
                        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                            Chk_Habilitar_Menu_Submenu.Value = 1
                        Else
                            Chk_Habilitar_Menu_Submenu.Value = 0
                        End If
                End Select
            End If
        End If
    End If
End If
End Sub

Private Sub Grid_Accesos_Seguridad_EnterCell()
Dim Fila As Integer           'Obtiene la fila que se esta consultando
Dim Columna As Integer        'Obtiene la columna que se esta consultando
Dim Fila_Actual As Integer    'Obtiene la fila actual en donde esta posicionado el setfocus
Dim Columna_Actual As Integer 'Obtiene la columna actual en donde esta posicionado el setfocus

Chk_Habilitar_Menu_Submenu.Visible = False
Txt_Habilitar.Visible = False
If Btn_Nuevo.Caption = "Dar de Alta" Or Btn_Modificar.Caption = "Actualizar" Then
    If Grid_Accesos_Seguridad.Col = 5 Then
        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 4) <> "Encabezado" Then
            Chk_Habilitar_Menu_Submenu.BackColor = &HFFFFFF
            Chk_Habilitar_Menu_Submenu.Visible = False
            For Fila = Grid_Accesos_Seguridad.RowSel To 1 Step -1
                If Grid_Accesos_Seguridad.TextMatrix(Fila, 0) = "-" Then
                    If Grid_Accesos_Seguridad.TextMatrix(Fila, 5) = "S" Then
                        Exit For
                    Else
                        Chk_Habilitar_Menu_Submenu.BackColor = &HFFFFFF
                        Chk_Habilitar_Menu_Submenu.Visible = False
                        Exit Sub
                    End If
                End If
            Next Fila
            Select Case Grid_Accesos_Seguridad.ColSel
            Case 5:
                Txt_Habilitar.Visible = False
                Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Accesos_Seguridad, Chk_Habilitar_Menu_Submenu)
                If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                    Chk_Habilitar_Menu_Submenu.Value = 1
                Else
                    Chk_Habilitar_Menu_Submenu.Value = 0
                End If
                Chk_Habilitar_Menu_Submenu.SetFocus
            End Select
        Else
            If Fra_Acceso_Sistema_Rol.Visible = True Then
                Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Accesos_Seguridad, Chk_Habilitar_Menu_Submenu)
                If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                    Chk_Habilitar_Menu_Submenu.Value = 1
                Else
                    Chk_Habilitar_Menu_Submenu.Value = 0
                End If
                Chk_Habilitar_Menu_Submenu.Visible = True
                Chk_Habilitar_Menu_Submenu.BackColor = vbButtonFace
            End If
        End If
    Else
        If Fra_Acceso_Sistema_Rol.Visible = True Then
            If Grid_Accesos_Seguridad.ColSel > 5 And Grid_Accesos_Seguridad.ColSel <= 9 Then
                Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Accesos_Seguridad, Txt_Habilitar)
            End If
        End If
    End If
End If
End Sub

Private Sub Grid_Accesos_Seguridad_Scroll()
    Txt_Habilitar.Visible = False
End Sub

Private Sub Grid_Cat_Almacenes_Click()
Dim Rs_Consulta_Cat_Almacenes As rdoResultset

    Set Conectar_Ayudante = New Ayudante
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Cat_Almacenes.Rows > 1 Then
        Txt_Almacen_ID.text = Trim(Grid_Cat_Almacenes.TextMatrix(Grid_Cat_Almacenes.RowSel, 0))
        'consulta todos los valores del registro que tiene seleccionado el usuario
        Mi_SQL = "SELECT * FROM Cat_Almacenes"
        Mi_SQL = Mi_SQL & " WHERE Almacen_ID ='" & Trim(Txt_Almacen_ID.text) & "'"
        Set Rs_Consulta_Cat_Almacen = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Almacen.EOF Then
            With Rs_Consulta_Cat_Almacen
                Txt_Comentarios_Cat_Almacenes.text = UCase(.rdoColumns("Comentarios"))
                Txt_Almacen_ID.text = .rdoColumns("Almacen_ID")
                Txt_Nombre_Cat_Almacenes.text = .rdoColumns("Nombre")
                Cmb_Estatus_Cat_Almacenes.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Cat_Almacenes)
            End With
        End If
        Rs_Consulta_Cat_Almacen.Close
    End If
End Sub

Private Sub Grid_Cat_Almacenes_EnterCell()
    Call Grid_Cat_Almacenes_Click
End Sub


Private Sub Grid_Cat_Bancos_Click()
Dim Rs_Consulta_Cat_Bancos As rdoResultset

    Set Conectar_Ayudante = New Ayudante
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Cat_Bancos.Rows > 1 Then
        Txt_Banco_ID.text = Trim(Grid_Cat_Bancos.TextMatrix(Grid_Cat_Bancos.RowSel, 0))
        'consulta todos los valores del registro que tiene seleccionado el usuario
        Mi_SQL = "SELECT * FROM Cat_Bancos"
        Mi_SQL = Mi_SQL & " WHERE Banco_ID ='" & Format(Txt_Banco_ID.text, "00000") & "'"
        Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Bancos.EOF Then
            With Rs_Consulta_Cat_Bancos
                If Not IsNull(.rdoColumns("Clabe_Interbancaria")) Then Txt_Clave_Interbancaria.text = .rdoColumns("Clabe_Interbancaria")
                If Not IsNull(.rdoColumns("Numero_Cuenta")) Then Txt_Numero_Cuenta.text = .rdoColumns("Numero_Cuenta")
                If Not IsNull(.rdoColumns("Numero_Inial_Cheque")) Then Txt_cheque_inicial.text = .rdoColumns("Numero_Inial_Cheque")
                If Not IsNull(.rdoColumns("Banco_ID")) Then Txt_Banco_ID.text = Format(.rdoColumns("Banco_ID"), "00")
                If Not IsNull(.rdoColumns("Nombre")) Then Txt_Nombre_Banco.text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Sucursal")) Then Ttx_Sucursal.text = .rdoColumns("Sucursal")
                Cmb_Estatus_Banco.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Banco)
                If Not IsNull(.rdoColumns("Formato")) Then Cmb_Formato.text = .rdoColumns("Formato")
                If Not IsNull(.rdoColumns("RFC")) Then Txt_RFC_Banco.text = .rdoColumns("RFC")
                If Not IsNull(.rdoColumns("Consecutivo_Cheque")) Then Txt_Consecutivo_Cheque.text = .rdoColumns("Consecutivo_Cheque")
            End With
        End If
        Rs_Consulta_Cat_Bancos.Close
    End If
End Sub

Private Sub Grid_Cat_Bancos_EnterCell()
    Call Grid_Cat_Bancos_Click
End Sub


Private Sub Grid_Cat_Categorias_Click()
Dim Rs_Consulta_Cat_Categorias As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Cat_Categorias.Rows > 1 Then
    Txt_Catgoria_ID.text = Trim(Grid_Cat_Categorias.TextMatrix(Grid_Cat_Categorias.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Categorias"
    Mi_SQL = Mi_SQL & " WHERE Categoria_ID ='" & Trim(Txt_Catgoria_ID.text) & "'"
    Set Rs_Consulta_Cat_Categorias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Categorias.EOF Then
        With Rs_Consulta_Cat_Categorias
            Txt_Comentarios_Categorias.text = UCase(Rs_Consulta_Cat_Categorias.rdoColumns("Comentarios"))
            Txt_Catgoria_ID.text = Rs_Consulta_Cat_Categorias.rdoColumns("Categoria_ID")
            Txt_Nombre_Categorias.text = Rs_Consulta_Cat_Categorias.rdoColumns("Nombre")
            Cmb_Estatus_Categoria.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Rs_Consulta_Cat_Categorias.rdoColumns("Estatus"), Cmb_Estatus_Categoria)
        End With
    End If
    Rs_Consulta_Cat_Categorias.Close
End If
End Sub

Private Sub Grid_Cat_Categorias_EnterCell()
    Call Grid_Cat_Categorias_Click
End Sub

Private Sub Grid_Cat_Laboratorios_Click()
Dim Rs_Consulta_Cat_Laboratorios As rdoResultset

    Set Conectar_Ayudante = New Ayudante
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Cat_Laboratorios.Rows > 1 Then
        Txt_Laborartorio_ID.text = Trim(Grid_Cat_Laboratorios.TextMatrix(Grid_Cat_Laboratorios.RowSel, 0))
        'consulta todos los valores del registro que tiene seleccionado el usuario
        Mi_SQL = "SELECT * FROM Cat_Laboratorios"
        Mi_SQL = Mi_SQL & " WHERE Laboratorio_ID ='" & Trim(Txt_Laborartorio_ID.text) & "'"
        Set Rs_Consulta_Cat_Laboratorios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Laboratorios.EOF Then
            With Rs_Consulta_Cat_Laboratorios
                Txt_Cometarios_Laboratorios.text = UCase(Rs_Consulta_Cat_Laboratorios.rdoColumns("Comentarios"))
                Txt_Laborartorio_ID.text = Rs_Consulta_Cat_Laboratorios.rdoColumns("Laboratorio_ID")
                Txt_Nombre_Laboratorios.text = Rs_Consulta_Cat_Laboratorios.rdoColumns("Nombre")
                Cmb_Estatus_Laboratorios.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Rs_Consulta_Cat_Laboratorios.rdoColumns("Estatus"), Cmb_Estatus_Laboratorios)
            End With
        End If
        Rs_Consulta_Cat_Laboratorios.Close
    End If
End Sub

Private Sub Grid_Cat_Laboratorios_EnterCell()
    Call Grid_Cat_Laboratorios_Click
End Sub


Private Sub Grid_Cat_Marcas_Click()
Dim Rs_Consulta_Cat_Marcas As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Cat_Marcas.Rows > 1 Then
    Txt_Marca_ID.text = Trim(Grid_Cat_Marcas.TextMatrix(Grid_Cat_Marcas.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Marcas"
    Mi_SQL = Mi_SQL & " WHERE Marca_ID ='" & Trim(Txt_Marca_ID.text) & "'"
    Set Rs_Consulta_Cat_Marcas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Marcas.EOF Then
        With Rs_Consulta_Cat_Marcas
            Txt_Comentarios_Marca.text = UCase(Rs_Consulta_Cat_Marcas.rdoColumns("Comentarios"))
            Txt_Marca_ID.text = Rs_Consulta_Cat_Marcas.rdoColumns("Marca_ID")
            Txt_Nombre_Marca.text = Rs_Consulta_Cat_Marcas.rdoColumns("Nombre")
            Cmb_Estatus_Marca.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Rs_Consulta_Cat_Marcas.rdoColumns("Estatus"), Cmb_Estatus_Marca)
        End With
    End If
    Rs_Consulta_Cat_Marcas.Close
End If
End Sub

Private Sub Grid_Cat_Marcas_EnterCell()
    Call Grid_Cat_Marcas_Click
End Sub

Private Sub Grid_Cat_Productos_Tipo_Click()
Dim Rs_Consulta_Cat_Productos_Tipo As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Cat_Productos_Tipo.Rows > 1 Then
    Txt_Tipo_ID.text = Trim(Grid_Cat_Productos_Tipo.TextMatrix(Grid_Cat_Productos_Tipo.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Productos_Tipo"
    Mi_SQL = Mi_SQL & " WHERE Tipo_ID ='" & Trim(Txt_Tipo_ID.text) & "'"
    Set Rs_Consulta_Cat_Productos_Tipo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Productos_Tipo.EOF Then
        With Rs_Consulta_Cat_Productos_Tipo
            Txt_Comentarios_Cat_Productos_Tipo.text = UCase(Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Comentarios"))
            Txt_Tipo_ID.text = Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Tipo_ID")
            Txt_Nombre_Cat_Productos_Tipo.text = Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Nombre")
            Cmb_Estatus_Cat_Productos_Tipo.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Estatus"), Cmb_Estatus_Cat_Productos_Tipo)
        End With
    End If
    Rs_Consulta_Cat_Productos_Tipo.Close
End If
End Sub

Private Sub Grid_Cat_Productos_Tipo_EnterCell()
    Call Grid_Cat_Productos_Tipo_Click
End Sub


Private Sub Grid_Cat_Sustancia_Activa_Click()
Dim Rs_Consulta_Cat_Sustancia_Activa As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Cat_Sustancia_Activa.Rows > 1 Then
    Txt_ID_Cat_Sustancia_Activa.text = Trim(Grid_Cat_Sustancia_Activa.TextMatrix(Grid_Cat_Sustancia_Activa.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Sustancia_Activa"
    Mi_SQL = Mi_SQL & " WHERE Sustancia_Activa_ID ='" & Trim(Txt_ID_Cat_Sustancia_Activa.text) & "'"
    Set Rs_Consulta_Cat_Sustancia_Activa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Sustancia_Activa.EOF Then
        With Rs_Consulta_Cat_Sustancia_Activa
            Txt_Comentarios_Cat_Sustancia_Activa.text = UCase(.rdoColumns("Comentarios"))
            Txt_ID_Cat_Sustancia_Activa.text = .rdoColumns("Sustancia_Activa_ID")
            Txt_Nombre_Cat_Sustancia_Activa.text = .rdoColumns("Nombre")
            Cmb_Cat_Sustancia_Activa.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Cat_Sustancia_Activa)
        End With
    End If
    Rs_Consulta_Cat_Sustancia_Activa.Close
End If
End Sub

Private Sub Grid_Cat_Sustancia_Activa_EnterCell()
    Call Grid_Cat_Sustancia_Activa_Click
End Sub


Private Sub Grid_Clasificacion_Clientes_Click()
Dim Rs_Consulta_Cat_Calsificacion_Clientes As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Clasificacion_Clientes.Rows > 1 Then
    Txt_Clasificacion_Cliente.text = Trim(Grid_Clasificacion_Clientes.TextMatrix(Grid_Clasificacion_Clientes.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Clientes_Clasificacion"
    Mi_SQL = Mi_SQL & " WHERE Clasificacion_ID ='" & Trim(Txt_Clasificacion_Cliente.text) & "'"
    Set Rs_Consulta_Cat_Calsificacion_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Calsificacion_Clientes.EOF Then
        With Rs_Consulta_Cat_Calsificacion_Clientes
            Txt_Comentarios_Clasificacion_Cliente.text = UCase(.rdoColumns("Comentarios"))
            Txt_Clasificacion_Cliente.text = .rdoColumns("Clasificacion_ID")
            Txt_Nombre_Clasificacion_Cliente.text = .rdoColumns("Nombre")
            Cmb_Estatus_Clasificacion_Cliente.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Clasificacion_Cliente)
        End With
    End If
    Rs_Consulta_Cat_Calsificacion_Clientes.Close
End If
End Sub

Private Sub Grid_Clasificacion_Clientes_EnterCell()
    Call Grid_Clasificacion_Clientes_Click
End Sub


Private Sub Grid_Clasificaciones_Click()
Dim Rs_Consulta_Cat_Calsificacion_Proveedores As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Clasificaciones.Rows > 1 Then
    Txt_Clasificacion_ID.text = Trim(Grid_Clasificaciones.TextMatrix(Grid_Clasificaciones.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Clasificacion_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Clasificacion_ID ='" & Trim(Txt_Clasificacion_ID.text) & "'"
    Set Rs_Consulta_Cat_Calsificacion_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Calsificacion_Proveedores.EOF Then
        With Rs_Consulta_Cat_Calsificacion_Proveedores
            Txt_Comentarios_Clasficacion_Proveedor.text = UCase(.rdoColumns(2))
            Txt_Clasificacion_ID.text = .rdoColumns("Clasificacion_ID")
            Txt_Nombre_Clasificacion_Proveedor.text = .rdoColumns("Nombre")
            Cmb_Estatus_Clasificacion_Proveedor.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Clasificacion_Proveedor)
        End With
    End If
    Rs_Consulta_Cat_Calsificacion_Proveedores.Close
End If
End Sub

Private Sub Grid_Clasificaciones_EnterCell()
    Call Grid_Clasificaciones_Click
End Sub


Private Sub Grid_Impuestos_Cat_Impuestos_Click()
Dim Rs_Consulta_Cat_Impuestos As rdoResultset

    Set Conectar_Ayudante = New Ayudante
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Impuestos_Cat_Impuestos.Rows > 1 Then
        Txt_ID_Cat_Impuestos.text = Trim(Grid_Impuestos_Cat_Impuestos.TextMatrix(Grid_Impuestos_Cat_Impuestos.RowSel, 0))
        'consulta todos los valores del registro que tiene seleccionado el usuario
        Mi_SQL = "SELECT * FROM Cat_Impuestos"
        Mi_SQL = Mi_SQL & " WHERE Impuesto_ID ='" & Trim(Txt_ID_Cat_Impuestos.text) & "'"
        Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Impuestos.EOF Then
            With Rs_Consulta_Cat_Impuestos
                Txt_Comentarios_Cat_Impuestos.text = UCase(.rdoColumns("Comentarios"))
                Txt_ID_Cat_Impuestos.text = .rdoColumns("Impuesto_ID")
                Txt_Impuesto_Cat_Impuestos.text = .rdoColumns("Impuesto")
                Cmb_Estatus_Cat_Impuestos.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Cat_Impuestos)
            End With
        End If
        Rs_Consulta_Cat_Impuestos.Close
    End If
End Sub

Private Sub Grid_Impuestos_Cat_Impuestos_EnterCell()
    Call Grid_Impuestos_Cat_Impuestos_Click
End Sub


Private Sub Grid_Presentaciones_Click()
Dim Rs_Consulta_Cat_Presentaciones As rdoResultset

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
If Grid_Presentaciones.Rows > 1 Then
    Txt_Presentacion_ID.text = Trim(Grid_Presentaciones.TextMatrix(Grid_Presentaciones.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Cat_Presentaciones"
    Mi_SQL = Mi_SQL & " WHERE Presentacion_ID ='" & Trim(Txt_Presentacion_ID.text) & "'"
    Set Rs_Consulta_Cat_Presentaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Presentaciones.EOF Then
        With Rs_Consulta_Cat_Bancos
            Txt_Comentarios_Presentaciones.text = UCase(Rs_Consulta_Cat_Presentaciones.rdoColumns("Comentarios"))
            Txt_Presentacion_ID.text = Rs_Consulta_Cat_Presentaciones.rdoColumns("Presentacion_ID")
            Txt_Nombre_Presentaciones.text = Rs_Consulta_Cat_Presentaciones.rdoColumns("Nombre")
            Cmb_Estaus_Presentaciones.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Rs_Consulta_Cat_Presentaciones.rdoColumns("Estatus"), Cmb_Estaus_Presentaciones)
        End With
    End If
    Rs_Consulta_Cat_Presentaciones.Close
End If
End Sub

Private Sub Grid_Presentaciones_EnterCell()
    Call Grid_Presentaciones_Click
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Grid_Roles_Click
    'DESCRIPCIÓN: Se consulta los menus y submenus que tiene asignado el rol
    '             que el usuario selecciono así como agrega los datos del rol
    '             en los controles correspondientes
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 28-Abril-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Grid_Roles_Click()
Dim Contador_Columnas As Integer                'Indica que columna del grid se esta consultando
Dim Ctl As Control                              'Indica que control es el que se esta consultando en el sistema
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta los menus y submnus que tiene asignados el usuario

'Asigna los valores correspondientes a los controles de la forma
Txt_Rol_ID.text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 0))
Txt_Nombre_Rol.text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 1))
Txt_Comentarios_Rol.text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 2))
Grid_Accesos_Seguridad.Rows = 0
'Agrega el encabezado el grid_seguridad
Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "Menu" & Chr(9) & "Submenu" & _
Chr(9) & "Nombre" & Chr(9) & "Tipo" & Chr(9) & "Habilitar" & _
Chr(9) & "Alta" & Chr(9) & "Cambio" & Chr(9) & "Eliminar" & Chr(9) & "Consultar"
'Consulta todos los controles que tiene la pantalla MDIFrm_Apl_Principal
For Each Ctl In MDIFrm_Apl_Principal.Controls
    If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Or UCase(Mid(Ctl.Name, 1, 7)) = UCase("SUBMENU") Then
        'Consulta si el usuario que fue seleccionado tiene habilitado el menu o submenu
        'que se esta consultando
        Mi_SQL = "SELECT * FROM Apl_Cat_Accesos"
        Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.text) & "'"
        Mi_SQL = Mi_SQL & " AND Nombre_Sistema = '" & Ctl.Name & "'"
        Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Si no se encuentra el menu que se esta consultando entonces este menu o submenu lo
        'agrega al grid_seguridad y con estatus deshabilitaho
        If Rs_Consulta_Apl_Cat_Accesos.EOF Then
            If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                Grid_Accesos_Seguridad.AddItem "-" & Chr(9) & _
                UCase(Ctl.Caption) & Chr(9) & "" & _
                Chr(9) & Ctl.Name & Chr(9) & "Encabezado" & _
                Chr(9) & "N" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                'Agrega el color gris a la fila que tiene el encabezado
                For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                    Grid_Accesos_Seguridad.Col = Contador_Columnas
                    Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                Next Contador_Columnas
            Else
                Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "" & Chr(9) & _
                UCase(Ctl.Caption) & Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & _
                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
            End If
        'Si lo encuentra entonces agrega el menu o submenu al grid_seguridad con el estatus que tiene
        'asignado
        Else
            If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                Grid_Accesos_Seguridad.AddItem "-" & _
                Chr(9) & UCase(Ctl.Caption) & Chr(9) & "" & _
                Chr(9) & Ctl.Name & Chr(9) & "Encabezado" & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Habilitar") & _
                Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                'Agrega el color gris a la fila que tiene el encabezado
                Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                    Grid_Accesos_Seguridad.Col = Contador_Columnas
                    Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                Next Contador_Columnas
            Else
                Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "" & _
                Chr(9) & UCase(Ctl.Caption) & _
                Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Habilitar") & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Alta") & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Cambio") & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Eliminar") & _
                Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Consultar")
            End If
        End If
        Rs_Consulta_Apl_Cat_Accesos.Close
    End If 'Menu/Submenu
Next Ctl
'Configura el tamaño de las columnas del grid_accesos_seguridad
If Grid_Accesos_Seguridad.Rows > 1 Then
    Grid_Accesos_Seguridad.FixedRows = 1
        Grid_Accesos_Seguridad.ColWidth(0) = 200 '-
        Grid_Accesos_Seguridad.ColWidth(1) = 1250 'Menu
        Grid_Accesos_Seguridad.ColWidth(2) = 2500 'SubMenu
        Grid_Accesos_Seguridad.ColWidth(3) = 0    'Nombre Menu/Submenu
        Grid_Accesos_Seguridad.ColWidth(4) = 0    'Tipo
        Grid_Accesos_Seguridad.ColWidth(5) = 900 'Habilitar
        Grid_Accesos_Seguridad.ColAlignment(5) = 3
        Grid_Accesos_Seguridad.ColWidth(6) = 600  'Alta
        Grid_Accesos_Seguridad.ColAlignment(6) = 3
        Grid_Accesos_Seguridad.ColWidth(7) = 650  'Cambio
        Grid_Accesos_Seguridad.ColAlignment(7) = 3
        Grid_Accesos_Seguridad.ColWidth(8) = 650  'Eliminar
        Grid_Accesos_Seguridad.ColAlignment(8) = 3
        Grid_Accesos_Seguridad.ColWidth(9) = 750  'Consultar
        Grid_Accesos_Seguridad.ColAlignment(9) = 3
    Collapsing = True
    Call Collapse_Grid
    Collapsing = False
End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Grid_Usuarios_Click
    'DESCRIPCIÓN: Consulta todos los datos del usuario que fue seleccionado
    'PARÁMETROS :
    'CREO       : Jorge Razo
    'FECHA_CREO:
    'MODIFICO          : Yazmin Delgado Gómez
    'FECHA_MODIFICO    : 13-Oct-2007
    'CAUSA_MODIFICACIÓN: Porque se añadieron los campos de fecha a caducar y
    '                    estatus del usuario
'*******************************************************************************
Private Sub Grid_Usuarios_Click()
Dim Rs_Consulta_Alp_Cat_Usuarios As rdoResultset  'Consulta los datos del registro que fue selecciondo por el usuario
Dim Comentarios As String

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
'Si el grid_usuarios tiene más de un solo registro entonces consulta los datos del registro
'que fue seleccionado por el usuario
If Grid_Usuarios.Rows > 1 Then
    Txt_Usuario_ID.text = Trim(Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID ='" & Trim(Txt_Usuario_ID.text) & "'"
    Set Rs_Consulta_Alp_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si encuentra los valores entonces agrega los valores a los controles correspondientes de
    'la forma
    If Not Rs_Consulta_Alp_Cat_Usuarios.EOF Then
        With Rs_Consulta_Alp_Cat_Usuarios
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Rol_ID"), Cmb_Roles)
            Txt_Nombre_Usuario.text = .rdoColumns("Nombre")
            If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Comentarios_Usuarios.text = .rdoColumns("Comentarios")
            Txt_Login.text = .rdoColumns("Login")
            Txt_Contraseña.text = .rdoColumns("Password")
            Txt_Confirmar_Contraseña.text = .rdoColumns("Password")
            DTP_Fecha_Caducar_Usuario.Value = Format(.rdoColumns("Fecha_Caduca"), "dd MMM yyyy")
            Cmb_Estatus_Usuario.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Usuario)
        End With
    End If
    Rs_Consulta_Alp_Cat_Usuarios.Close
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Txt_Comentarios_Rol_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Comentarios_Usuarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Confirmar_Contraseña_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Contraseña_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Habilitar_Click()
    If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5)) = "S" Then
        If Trim(Txt_Habilitar.text) = "S" Then
            Txt_Habilitar.text = "N"
        Else
            Txt_Habilitar.text = "S"
        End If
        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = Txt_Habilitar.text
    End If
End Sub

Private Sub Txt_Habilitar_KeyDown(KeyCode As Integer, Shift As Integer)
If Grid_Accesos_Seguridad.Rows > 1 Then
    If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) = "" Then
        If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
            If KeyCode > 37 Then Grid_Accesos_Seguridad.SetFocus
                If KeyCode = 37 Then
                    If Txt_Habilitar.SelStart = 0 Then
                        Grid_Accesos_Seguridad.SetFocus
                        If Grid_Accesos_Seguridad.Col > 5 Then
                            Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel - 1
                        End If
                    End If
                End If
                If Grid_Accesos_Seguridad.Row > 2 Then
                    If KeyCode = 38 Then Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel - 1
                    If KeyCode = 40 Then
                        If Grid_Accesos_Seguridad.Row < Grid_Accesos_Seguridad.Rows - 1 Or Grid_Accesos_Seguridad.Row = 1 Then
                            Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    If Grid_Accesos_Seguridad.Col >= 6 And Grid_Accesos_Seguridad.Col < 9 Then
                        If KeyCode = 39 Then Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel + 1
                    End If
                Else
                    If KeyCode = 40 Then
                        If Grid_Accesos_Seguridad.Row < Grid_Accesos_Seguridad.Rows - 1 Or Grid_Accesos_Seguridad.Row > 2 Then
                            Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    If Grid_Accesos_Seguridad.Col >= 6 And Grid_Accesos_Seguridad.Col < 9 Then
                        If KeyCode = 39 Then Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel + 1
                    End If
                End If
                If Txt_Habilitar.Visible = True Then
                    Txt_Habilitar.SetFocus
                    SendKeys "{Home}+{End}"
                End If
            End If
        Else
            Txt_Habilitar.Visible = False
        End If
    Else
        Txt_Habilitar.Visible = False
    End If
End Sub

Private Sub Txt_Habilitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5)) = "S" Then
            If Trim(Txt_Habilitar.text) = "S" Then
                Txt_Habilitar.text = "N"
            Else
                Txt_Habilitar.text = "S"
            End If
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = Txt_Habilitar.text
        End If
    End If
End Sub

Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Nombre_Rol_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Nombre_Usuario_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Valida_Login_Password_Usuario
    'DESCRIPCIÓN: Valida que no se repita algun Login ya existente cuando se agrega
    '             o se modifica algun usuario
    'PARÁMETROS: 1. Campo: Indica el campo que se va a comparar
    '            2. Valor: Indica el valor con el que se va a comparar el campo
    '            3. Cat_Usuario_ID: Si tiene algun valor, es porque se va a comprarar cuando se haga alguna modificacion
    'CREO: Susana Ledesma Ramírez
    'FECHA_CREO: 26/Abril/2006
    'MODIFICO:
    'FECHA_MODIFICO: 25/Octubre/2007
    'CAUSA_MODIFICACIÓN: Porque se necesitaba validar tambien el password
'*******************************************************************************

Public Function Valida_Login_Password_Usuario(Campo As String, Valor As String, Optional Cat_Usuario_ID As String) As Boolean
Dim Rs_Consulta_Cat_Usuarios As rdoResultset    'Maneja el registro de la Tabla de Cat_Usuarios

Set Conectar_Ayudante = New Ayudante
'Establece la consulta en Cat_Usuarios para saber el Login o el Password ya existen
Mi_SQL = "SELECT Login FROM Apl_Cat_Usuarios WHERE " & Campo & " = '" & Valor & "'"
'Si es alguna modificaciones, entonces se busca en todos los usuarios, excepto en el actual
If Cat_Usuario_ID <> "" Then Mi_SQL = Mi_SQL & " AND Usuario_ID<>'" & Cat_Usuario_ID & "'"
Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Si encuantra algun dato; entonces ese login o password ya existe
If Not Rs_Consulta_Cat_Usuarios.EOF Then
    Valida_Login_Password_Usuario = True 'Si el Login o password ya existe
End If
Rs_Consulta_Cat_Usuarios.Close
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Clasificacion_Proveedor
'DESCRIPCIÓN                : Da de alta la clasificacion de Proveedores
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 12 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Clasificacion_Proveedor()
Dim Rs_Alta_Cat_Clasificacion_Proveedores As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de Clasificacion Proveedores
    Set Rs_Alta_Cat_Clasificacion_Proveedores = Conectar_Ayudante.Recordset_Agregar("Cat_Clasificacion_Proveedores")
    'Llena la tabla de Cat_Clasificacion_proveedores con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Clasificacion_Proveedores
        .AddNew
            .rdoColumns("Clasificacion_ID") = Trim(Txt_Clasificacion_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Clasificacion_Proveedor.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Clasificacion_Proveedor.text))
            If Trim(Txt_Comentarios_Clasficacion_Proveedor.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Clasficacion_Proveedor.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Clasificacion_Proveedores.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Clasificacion_Proveedores.Enabled = False
    Fra_Detalles_Clasificacion_Proveedores.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Clasificacion_Proveedor.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Clasificaciones.Rows = 0 Then
        Grid_Clasificaciones.AddItem "Clasificación ID" & Chr(9) & "Nombre"
        Grid_Clasificaciones.AddItem Trim(Txt_Clasificacion_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Clasificacion_Proveedor.text)
        Grid_Clasificaciones.ColWidth(0) = 1000
        Grid_Clasificaciones.ColAlignment(0) = 3
        Grid_Clasificaciones.ColWidth(1) = 6629
        Grid_Clasificaciones.ColAlignment(1) = 2
        Grid_Clasificaciones.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Clasificacion_Proveedores("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Proveedores", Frm_Cat_Generales)
    MsgBox "Clasificación dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Clasificacion_Proveedor
'DESCRIPCIÓN                : Modifica la clasificacion de Proveedores
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 12 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Clasificacion_Proveedor()
Dim Rs_Modificacion_Cat_Clasificacion_Proveedores As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Clasificacion_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Clasificacion_ID ='" & Trim(Txt_Clasificacion_ID.text) & "'"
    Set Rs_Modificacion_Cat_Clasificacion_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Clasificacion_Proveedores
    With Rs_Modificacion_Cat_Clasificacion_Proveedores
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Clasificacion_Proveedor.text))
            .rdoColumns("Estatus") = Trim(Cmb_Estatus_Clasificacion_Proveedor.text)
            If Trim(Txt_Comentarios_Clasficacion_Proveedor.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Clasficacion_Proveedor.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Clasificacion_Proveedores.Close
    Grid_Clasificaciones.TextMatrix(Grid_Clasificaciones.RowSel, 0) = Trim(UCase(Txt_Clasificacion_ID.text))
    Grid_Clasificaciones.TextMatrix(Grid_Clasificaciones.RowSel, 1) = Trim(Txt_Nombre_Clasificacion_Proveedor.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Clasificacion_Proveedores.Enabled = False
    Fra_Detalles_Clasificacion_Proveedores.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Proveedor", Frm_Cat_Generales)
    MsgBox "La Clasificacion ha sido modificada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Clasificacion_Proveedores
'DESCRIPCIÓN                : Consulta la clasificacion de Proveedores
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo cpomo criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 12 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Clasificacion_Proveedores(Nombre As String)
Dim Rs_Consulta_Cat_Clasificacion_Proveedores As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Clasificaciones.Rows = 0
'Consulta los datos generales del la clasificacion de proveedor
Mi_SQL = "SELECT ClasificaCion_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Clasificacion_Proveedores"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY ClasificaCion_ID"
Set Rs_Consulta_Cat_Clasificacion_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Clasificacion_Proveedores.EOF Then
    'Coloca un encabezado en el grid
    Grid_Clasificaciones.AddItem "Clasificacion ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Clasificacion_Proveedores.EOF
        Grid_Clasificaciones.AddItem Rs_Consulta_Cat_Clasificacion_Proveedores.rdoColumns("Clasificacion_ID") _
       & Chr(9) & Rs_Consulta_Cat_Clasificacion_Proveedores.rdoColumns("Nombre")
        Rs_Consulta_Cat_Clasificacion_Proveedores.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Clasificaciones.ColWidth(0) = 1000
    Grid_Clasificaciones.ColAlignment(0) = 3
    Grid_Clasificaciones.ColWidth(1) = 6629
    Grid_Clasificaciones.ColAlignment(1) = 2
    Grid_Clasificaciones.FixedRows = 1
    Grid_Clasificaciones.FixedCols = 1
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Clasificacion_Proveedores.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Clasificacion_Cliente
'DESCRIPCIÓN                : Da de alta la clasificacion de los clientes
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 13 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Clasificacion_Cliente()
Dim Rs_Alta_Cat_Clientes_Clasificacion As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de Clasificacion Clientes
    Set Rs_Alta_Cat_Clientes_Clasificacion = Conectar_Ayudante.Recordset_Agregar("Cat_Clientes_Clasificacion")
    'Llena la tabla de Cat_Clientes_Clasificacion con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Clientes_Clasificacion
        .AddNew
            .rdoColumns("Clasificacion_ID") = Trim(Txt_Clasificacion_Cliente.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Clasificacion_Cliente.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Clasificacion_Cliente.text))
            If Trim(Txt_Comentarios_Clasficacion_Proveedor.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Clasificacion_Cliente.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Clientes_Clasificacion.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Clasificacion_Clientes.Enabled = False
    Fra_Clasificacion_Clientes_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Clasificacion_Cliente.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Clasificacion_Clientes.Rows = 0 Then
        Grid_Clasificacion_Clientes.AddItem "Clasificación ID" & Chr(9) & "Nombre"
        Grid_Clasificacion_Clientes.AddItem Trim(Txt_Clasificacion_Cliente.text) & Chr(9) & _
        UCase(Txt_Nombre_Clasificacion_Cliente.text)
        Grid_Clasificacion_Clientes.ColWidth(0) = 1000
        Grid_Clasificacion_Clientes.ColAlignment(0) = 3
        Grid_Clasificacion_Clientes.ColWidth(1) = 6629
        Grid_Clasificacion_Clientes.ColAlignment(1) = 2
        Grid_Clasificacion_Clientes.FixedRows = 1
     End If

    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Clasificacion_Clientes("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Clientes", Frm_Cat_Generales)
    MsgBox "Clasificación dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Clasificacion_Cliente
'DESCRIPCIÓN                : Modifica la clasificacion de Proveedores
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 12 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Clasificacion_Cliente()
Dim Rs_Modificacion_Cat_Clasificacion_Clientes As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Clientes_Clasificacion"
    Mi_SQL = Mi_SQL & " WHERE Clasificacion_ID ='" & Trim(Txt_Clasificacion_Cliente.text) & "'"
    Set Rs_Modificacion_Cat_Clasificacion_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Clientes_Clasificacion
    With Rs_Modificacion_Cat_Clasificacion_Clientes
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Clasificacion_Cliente.text))
            .rdoColumns("Estatus") = Trim(Cmb_Estatus_Clasificacion_Cliente.text)
            If Trim(Txt_Comentarios_Clasificacion_Cliente.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Clasificacion_Cliente.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Clasificacion_Clientes.Close
    Grid_Clasificacion_Clientes.TextMatrix(Grid_Clasificacion_Clientes.RowSel, 0) = Trim(UCase(Txt_Clasificacion_Cliente.text))
    Grid_Clasificacion_Clientes.TextMatrix(Grid_Clasificacion_Clientes.RowSel, 1) = Trim(Txt_Nombre_Clasificacion_Cliente.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Clasificacion_Clientes.Enabled = False
    Fra_Clasificacion_Clientes_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Clientes", Frm_Cat_Generales)
    MsgBox "La Clasificacion ha sido modificada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Clasificacion_Clientes
'DESCRIPCIÓN                : Consulta la clasificacion de los clientes
'PARÁMETROS                 : Nombre; sirve para hacer la consulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 13 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Clasificacion_Clientes(Nombre As String)
Dim Rs_Consulta_Cat_Clasificacion_Clientes As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Clasificacion_Clientes.Rows = 0
'Consulta los datos generales de la clasificacion cliente
Mi_SQL = "SELECT ClasificaCion_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Clientes_Clasificacion"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY ClasificaCion_ID"
Set Rs_Consulta_Cat_Clasificacion_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Clasificacion_Clientes.EOF Then
    'Coloca un encabezado en el grid
    Grid_Clasificacion_Clientes.AddItem "Clasificacion ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Clasificacion_Clientes.EOF
        Grid_Clasificacion_Clientes.AddItem Rs_Consulta_Cat_Clasificacion_Clientes.rdoColumns("Clasificacion_ID") _
        & Chr(9) & Rs_Consulta_Cat_Clasificacion_Clientes.rdoColumns("Nombre")
        Grid_Clasificacion_Clientes.FixedRows = 1
        Rs_Consulta_Cat_Clasificacion_Clientes.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Clasificacion_Clientes.ColWidth(0) = 1200
    Grid_Clasificacion_Clientes.ColAlignment(0) = 3
    Grid_Clasificacion_Clientes.ColWidth(1) = 6429
    Grid_Clasificacion_Clientes.ColAlignment(1) = 2
    Grid_Clasificacion_Clientes.Col = 0
    Grid_Clasificacion_Clientes.Row = 1
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Clasificacion_Clientes.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Bancos
'DESCRIPCIÓN                : Da de alta los Bancos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Bancos()
Dim Rs_Alta_Cat_Bancos As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Bancos
    Set Rs_Alta_Cat_Bancos = Conectar_Ayudante.Recordset_Agregar("Cat_Bancos")
    'Llena la tabla de Cat_Bancos
    With Rs_Alta_Cat_Bancos
        .AddNew
            .rdoColumns("Banco_ID") = Format(Txt_Banco_ID.text, "00000")
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Banco.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Banco.text))
            .rdoColumns("Sucursal") = Trim(UCase(Ttx_Sucursal.text))
            .rdoColumns("Formato") = Trim(UCase(Cmb_Formato.text))
           .rdoColumns("Clabe_Interbancaria") = Txt_Clave_Interbancaria.text
           .rdoColumns("Numero_Cuenta") = Txt_Numero_Cuenta.text
           .rdoColumns("Numero_Inial_Cheque") = Val(Txt_cheque_inicial.text)
           .rdoColumns("Consecutivo_Cheque") = 0
           .rdoColumns("RFC") = Txt_RFC_Banco.text
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Bancos.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Bancos.Enabled = False
    Fra_Bancos_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Banco.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Bancos.Rows = 0 Then
        Grid_Cat_Bancos.AddItem "Banco ID" & Chr(9) & "Nombre"
        Grid_Cat_Bancos.AddItem Trim(Txt_Banco_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Banco.text)
        Grid_Cat_Bancos.ColWidth(0) = 1000
        Grid_Cat_Bancos.ColAlignment(0) = 3
        Grid_Cat_Bancos.ColWidth(1) = 6629
        Grid_Cat_Bancos.ColAlignment(1) = 2
        Grid_Cat_Bancos.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Bancos("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Bancos", Frm_Cat_Generales)
    MsgBox "Banco dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Bancos
'DESCRIPCIÓN                : Consulta los bancos
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Bancos(Nombre As String)
Dim Rs_Consulta_Cat_Bancos As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Bancos.Rows = 0
'Consulta los datos generales del la clasificacion de proveedor
Mi_SQL = "SELECT Banco_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Bancos"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Banco_ID"
Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Bancos.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Bancos.AddItem "Clasificacion ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Bancos.EOF
        Grid_Cat_Bancos.AddItem Format(Rs_Consulta_Cat_Bancos.rdoColumns("Banco_ID"), "00") _
        & Chr(9) & Rs_Consulta_Cat_Bancos.rdoColumns("Nombre")
        Grid_Cat_Bancos.FixedRows = 1
        Rs_Consulta_Cat_Bancos.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Bancos.ColWidth(0) = 1000
    Grid_Cat_Bancos.ColAlignment(0) = 3
    Grid_Cat_Bancos.ColWidth(1) = 6629
    Grid_Cat_Bancos.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Bancos.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Bancos
'DESCRIPCIÓN                : Modifica los Bancos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Bancos()
Dim Rs_Modificacion_Cat_Bancos As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Banco actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Banco_ID ='" & Format(Txt_Banco_ID.text, "00000") & "'"
    Set Rs_Modificacion_Cat_Bancos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Bancos
    With Rs_Modificacion_Cat_Bancos
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Banco.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Banco.text))
            .rdoColumns("Sucursal") = Trim(UCase(Ttx_Sucursal.text))
            .rdoColumns("Clabe_Interbancaria") = Txt_Clave_Interbancaria.text
            .rdoColumns("Numero_Cuenta") = Txt_Numero_Cuenta.text
            .rdoColumns("Numero_Inial_Cheque") = Val(Txt_cheque_inicial.text)
            .rdoColumns("Consecutivo_Cheque") = 0
            .rdoColumns("Formato") = Trim(UCase(Cmb_Formato.text))
            .rdoColumns("RFC") = Txt_RFC_Banco.text
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Bancos.Close
    Grid_Cat_Bancos.TextMatrix(Grid_Cat_Bancos.RowSel, 0) = Trim(UCase(Txt_Banco_ID.text))
    Grid_Cat_Bancos.TextMatrix(Grid_Cat_Bancos.RowSel, 1) = Trim(Txt_Nombre_Banco.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Bancos.Enabled = False
    Fra_Bancos_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Bancos", Frm_Cat_Generales)
    MsgBox "El Banco ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Presentaciones
'DESCRIPCIÓN                : Da de alta las Presentaciones
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Presentaciones()
Dim Rs_Alta_Cat_Presentaciones As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Presentacion
    Set Rs_Alta_Cat_Presentaciones = Conectar_Ayudante.Recordset_Agregar("Cat_Presentaciones")
    'Llena la tabla de Cat_Presentaciones
    With Rs_Alta_Cat_Presentaciones
        .AddNew
            .rdoColumns("Presentacion_ID") = Trim(Txt_Presentacion_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Presentaciones.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estaus_Presentaciones.text))
            If Trim(Txt_Comentarios_Presentaciones.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Presentaciones.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Presentaciones.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Presentaciones.Enabled = False
    Fra_Detalles_Presentaciones.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estaus_Presentaciones.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Presentaciones.Rows = 0 Then
        Grid_Presentaciones.AddItem "Presentación ID" & Chr(9) & "Nombre"
        Grid_Presentaciones.AddItem Trim(Txt_Presentacion_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Presentaciones.text)
        Grid_Presentaciones.ColWidth(0) = 1000
        Grid_Presentaciones.ColAlignment(0) = 3
        Grid_Presentaciones.ColWidth(1) = 6629
        Grid_Presentaciones.ColAlignment(1) = 2
        Grid_Presentaciones.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Presentaciones("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Presentaciones", Frm_Cat_Generales)
    MsgBox "Presentación dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta Cat_Presentaciones
'DESCRIPCIÓN                : Consulta las Presentaciones
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Presentaciones(Nombre As String)
Dim Rs_Consulta_Cat_Presentaciones As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Presentaciones.Rows = 0
'Consulta los datos generales de Cat_Presentaciones
Mi_SQL = "SELECT Presentacion_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Presentaciones"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Presentacion_ID"
Set Rs_Consulta_Cat_Presentaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Presentaciones.EOF Then
    'Coloca un encabezado en el grid
    Grid_Presentaciones.AddItem "Presentacion ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Presentaciones.EOF
        Grid_Presentaciones.AddItem Rs_Consulta_Cat_Presentaciones.rdoColumns("Presentacion_ID") _
        & Chr(9) & Rs_Consulta_Cat_Presentaciones.rdoColumns("Nombre")
        Grid_Presentaciones.FixedRows = 1
        Rs_Consulta_Cat_Presentaciones.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Presentaciones.ColWidth(0) = 1000
    Grid_Presentaciones.ColAlignment(0) = 3
    Grid_Presentaciones.ColWidth(1) = 6629
    Grid_Presentaciones.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Presentaciones.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Presentaciones
'DESCRIPCIÓN                : Modifica las Presentaciones
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Presentaciones()
Dim Rs_Modificacion_Cat_Presentaciones As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta la Presentaciones
    Mi_SQL = "SELECT * FROM Cat_Presentaciones"
    Mi_SQL = Mi_SQL & " WHERE Presentacion_ID ='" & Trim(Txt_Presentacion_ID.text) & "'"
    Set Rs_Modificacion_Cat_Presentaciones = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Bancos
    With Rs_Modificacion_Cat_Presentaciones
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Presentaciones.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estaus_Presentaciones.text))
            If Trim(Txt_Comentarios_Presentaciones.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Presentaciones.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Presentaciones.Close
    Grid_Presentaciones.TextMatrix(Grid_Presentaciones.RowSel, 0) = Trim(UCase(Txt_Presentacion_ID.text))
    Grid_Presentaciones.TextMatrix(Grid_Presentaciones.RowSel, 1) = Trim(Txt_Nombre_Presentaciones.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Presentaciones.Enabled = False
    Fra_Detalles_Presentaciones.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Presentaciones", Frm_Cat_Generales)
    MsgBox "La Presentación ha sido modificada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Categorias
'DESCRIPCIÓN                : Da de alta las Categorias
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Categorias()
Dim Rs_Alta_Cat_Categorias As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Categorias
    Set Rs_Alta_Cat_Categorias = Conectar_Ayudante.Recordset_Agregar("Cat_Categorias")
    'Llena la tabla de Cat_Categorias
    With Rs_Alta_Cat_Categorias
        .AddNew
            .rdoColumns("Categoria_ID") = Trim(Txt_Catgoria_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Categorias.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Categoria.text))
            If Trim(Txt_Comentarios_Categorias.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Categorias.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Categorias.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Categorias.Enabled = False
    Fra_Detalles_Categorias.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Categoria.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Categorias.Rows = 0 Then
        Grid_Cat_Categorias.AddItem "Categoría ID" & Chr(9) & "Nombre"
        Grid_Cat_Categorias.AddItem Trim(Txt_Catgoria_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Categorias.text)
        Grid_Cat_Categorias.ColWidth(0) = 1000
        Grid_Cat_Categorias.ColAlignment(0) = 3
        Grid_Cat_Categorias.ColWidth(1) = 6629
        Grid_Cat_Categorias.ColAlignment(1) = 2
        Grid_Cat_Categorias.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Categorias("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Categorias", Frm_Cat_Generales)
    MsgBox "Categoría dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta Consulta_Cat_Categorias
'DESCRIPCIÓN                : Consulta las Categorias
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Categorias(Nombre As String)
Dim Rs_Consulta_Cat_Categorias As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Categorias.Rows = 0
'Consulta los datos generales de Cat_Categoria
Mi_SQL = "SELECT Categoria_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Categorias"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Categoria_ID"
Set Rs_Consulta_Cat_Categorias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Categorias.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Categorias.AddItem "Categoria ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Categorias.EOF
        Grid_Cat_Categorias.AddItem Rs_Consulta_Cat_Categorias.rdoColumns("Categoria_ID") _
        & Chr(9) & Rs_Consulta_Cat_Categorias.rdoColumns("Nombre")
        Grid_Cat_Categorias.FixedRows = 1
        Rs_Consulta_Cat_Categorias.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Categorias.ColWidth(0) = 1000
    Grid_Cat_Categorias.ColAlignment(0) = 3
    Grid_Cat_Categorias.ColWidth(1) = 6629
    Grid_Cat_Categorias.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Categorias.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Categorias
'DESCRIPCIÓN                : Modifica las Categorias
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Categorias()
Dim Rs_Modificacion_Cat_Categorias As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta la Categoria
    Mi_SQL = "SELECT * FROM Cat_Categorias"
    Mi_SQL = Mi_SQL & " WHERE Categoria_ID ='" & Trim(Txt_Catgoria_ID.text) & "'"
    Set Rs_Modificacion_Cat_Categorias = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Categorias
    With Rs_Modificacion_Cat_Categorias
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Categorias.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Categoria.text))
            If Trim(Txt_Comentarios_Categorias.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Categorias.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Categorias.Close
    Grid_Cat_Categorias.TextMatrix(Grid_Cat_Categorias.RowSel, 0) = Trim(UCase(Txt_Catgoria_ID.text))
    Grid_Cat_Categorias.TextMatrix(Grid_Cat_Categorias.RowSel, 1) = Trim(Txt_Nombre_Categorias.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Categorias.Enabled = False
    Fra_Detalles_Categorias.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Categorias", Frm_Cat_Generales)
    MsgBox "La Categoria ha sido modificada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Marcas
'DESCRIPCIÓN                : Da de alta las Marcas
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Marcas()
Dim Rs_Alta_Cat_Marcas As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Marcas
    Set Rs_Alta_Cat_Marcas = Conectar_Ayudante.Recordset_Agregar("Cat_Marcas")
    'Llena la tabla de Cat_Marcas
    With Rs_Alta_Cat_Marcas
        .AddNew
            .rdoColumns("Marca_ID") = Trim(Txt_Marca_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Marca.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Marca.text))
            If Trim(Txt_Comentarios_Marca.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Marca.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Marcas.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Marcas.Enabled = False
    Fra_Detalles_Marcas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Marca.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Marcas.Rows = 0 Then
        Grid_Cat_Marcas.AddItem "Marca ID" & Chr(9) & "Nombre"
        Grid_Cat_Marcas.AddItem Trim(Txt_Marca_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Marca.text)
        Grid_Cat_Marcas.ColWidth(0) = 1000
        Grid_Cat_Marcas.ColAlignment(0) = 3
        Grid_Cat_Marcas.ColWidth(1) = 6629
        Grid_Cat_Marcas.ColAlignment(1) = 2
        Grid_Cat_Marcas.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Marcas("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Marcas", Frm_Cat_Generales)
    MsgBox "Marca dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Cat_Marcas
'DESCRIPCIÓN                : Consulta las Marcas
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Marcas(Nombre As String)
Dim Rs_Consulta_Cat_Marcas As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Marcas.Rows = 0
'Consulta los datos generales de Cat_Marcas
Mi_SQL = "SELECT Marca_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Marcas"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Nombre"
Set Rs_Consulta_Cat_Marcas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Marcas.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Marcas.AddItem "Categoria ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Marcas.EOF
        Grid_Cat_Marcas.AddItem Rs_Consulta_Cat_Marcas.rdoColumns("Marca_ID") _
        & Chr(9) & Rs_Consulta_Cat_Marcas.rdoColumns("Nombre")
        Grid_Cat_Marcas.FixedRows = 1
        Rs_Consulta_Cat_Marcas.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Marcas.ColWidth(0) = 1000
    Grid_Cat_Marcas.ColAlignment(0) = 3
    Grid_Cat_Marcas.ColWidth(1) = 6629
    Grid_Cat_Marcas.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Marcas.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Marcas
'DESCRIPCIÓN                : Modifica las Marcas
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 19 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Marcas()
Dim Rs_Modificacion_Cat_Marcas As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta las Marcas
    Mi_SQL = "SELECT * FROM Cat_Marcas"
    Mi_SQL = Mi_SQL & " WHERE Marca_ID ='" & Trim(Txt_Marca_ID.text) & "'"
    Set Rs_Modificacion_Cat_Marcas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Marcas
    With Rs_Modificacion_Cat_Marcas
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Marca.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Marca.text))
            If Trim(Txt_Comentarios_Marca.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Marca.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Marcas.Close
    Grid_Cat_Marcas.TextMatrix(Grid_Cat_Marcas.RowSel, 0) = Trim(UCase(Txt_Marca_ID.text))
    Grid_Cat_Marcas.TextMatrix(Grid_Cat_Marcas.RowSel, 1) = Trim(Txt_Nombre_Marca.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Marcas.Enabled = False
    Fra_Detalles_Marcas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Marcas", Frm_Cat_Generales)
    MsgBox "La Marca ha sido modificada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Laboratorios
'DESCRIPCIÓN                : Da de alta los laboratorios
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 23 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Laboratorios()
Dim Rs_Alta_Cat_Laboratorios As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Categorias
    Set Rs_Alta_Cat_Laboratorios = Conectar_Ayudante.Recordset_Agregar("Cat_Laboratorios")
    'Llena la tabla de Cat_Categorias
    With Rs_Alta_Cat_Laboratorios
        .AddNew
            .rdoColumns("Laboratorio_ID") = Trim(Txt_Laborartorio_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Laboratorios.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Laboratorios.text))
            If Trim(Txt_Cometarios_Laboratorios.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Cometarios_Laboratorios.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Laboratorios.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Laboratorio.Enabled = False
    Fra_Detalles_Laboratorios.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Laboratorios.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Laboratorios.Rows = 0 Then
        Grid_Cat_Laboratorios.AddItem "Laboratorio ID" & Chr(9) & "Nombre"
        Grid_Cat_Laboratorios.AddItem Trim(Txt_Laborartorio_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Laboratorios.text)
        Grid_Cat_Laboratorios.ColWidth(0) = 1000
        Grid_Cat_Laboratorios.ColAlignment(0) = 3
        Grid_Cat_Laboratorios.ColWidth(1) = 6629
        Grid_Cat_Laboratorios.ColAlignment(1) = 2
        Grid_Cat_Laboratorios.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Laboratorios("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Laboratorios", Frm_Cat_Generales)
    MsgBox "Laboratorio dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta Consulta_Cat_Laboratorios
'DESCRIPCIÓN                : Consulta los Laboratorios
'PARÁMETROS                 : Nombre; sirve para hacer la consulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 23 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Laboratorios(Nombre As String)
Dim Rs_Consulta_Cat_Laboratorios As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Laboratorios.Rows = 0
'Consulta los datos generales de Cat_Laboratorios
Mi_SQL = "SELECT Laboratorio_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Laboratorios"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Nombre"
Set Rs_Consulta_Cat_Laboratorios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Laboratorios.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Laboratorios.AddItem "Laboratorio ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Laboratorios.EOF
        Grid_Cat_Laboratorios.AddItem Rs_Consulta_Cat_Laboratorios.rdoColumns("Laboratorio_ID") _
        & Chr(9) & Rs_Consulta_Cat_Laboratorios.rdoColumns("Nombre")
        Grid_Cat_Laboratorios.FixedRows = 1
        Rs_Consulta_Cat_Laboratorios.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Laboratorios.ColWidth(0) = 1000
    Grid_Cat_Laboratorios.ColAlignment(0) = 3
    Grid_Cat_Laboratorios.ColWidth(1) = 6629
    Grid_Cat_Laboratorios.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Laboratorios.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Laboratorios
'DESCRIPCIÓN                : Modifica los laboratorios
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 23 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Laboratorios()
Dim Rs_Modificacion_Cat_Laboratorios As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta los laboratorios
    Mi_SQL = "SELECT * FROM Cat_Laboratorios"
    Mi_SQL = Mi_SQL & " WHERE Laboratorio_ID ='" & Trim(Txt_Laborartorio_ID.text) & "'"
    Set Rs_Modificacion_Cat_Laboratorios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Laboratorios
    With Rs_Modificacion_Cat_Laboratorios
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Laboratorios.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Laboratorios.text))
            If Trim(Txt_Cometarios_Laboratorios.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Cometarios_Laboratorios.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Laboratorios.Close
    Grid_Cat_Laboratorios.TextMatrix(Grid_Cat_Laboratorios.RowSel, 0) = Trim(UCase(Txt_Laborartorio_ID.text))
    Grid_Cat_Laboratorios.TextMatrix(Grid_Cat_Laboratorios.RowSel, 1) = Trim(Txt_Nombre_Laboratorios.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Laboratorio.Enabled = False
    Fra_Detalles_Laboratorios.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Laboratorios", Frm_Cat_Generales)
    MsgBox "El Laboratorio ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Almacenes
'DESCRIPCIÓN                : Da de alta los almacenes
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 26 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Almacenes()
Dim Rs_Alta_Cat_Almacenes As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de almacenes
    Set Rs_Alta_Cat_Almacenes = Conectar_Ayudante.Recordset_Agregar("Cat_Almacenes")
    'Llena la tabla de Cat_Almacenes
    With Rs_Alta_Cat_Almacenes
        .AddNew
            .rdoColumns("Almacen_ID") = Trim(Txt_Almacen_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Almacenes.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Almacenes.text))
            If Trim(Txt_Comentarios_Cat_Almacenes.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Almacenes.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Almacenes.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Almacenes.Enabled = False
    Fra_Cat_Almacenes_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Cat_Almacenes.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Almacenes.Rows = 0 Then
        Grid_Cat_Almacenes.AddItem "Almacen ID" & Chr(9) & "Nombre"
        Grid_Cat_Almacenes.AddItem Trim(Txt_Almacen_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Cat_Almacenes.text)
        Grid_Cat_Almacenes.ColWidth(0) = 1000
        Grid_Cat_Almacenes.ColAlignment(0) = 3
        Grid_Cat_Almacenes.ColWidth(1) = 6629
        Grid_Cat_Almacenes.ColAlignment(1) = 2
        Grid_Cat_Almacenes.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Almacenes("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Almacenes", Frm_Cat_Generales)
    MsgBox "Almacen dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Cat_Almacenes
'DESCRIPCIÓN                : Consulta los almacenes
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 26 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Almacenes(Nombre As String)
Dim Rs_Consulta_Cat_Almacenes As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Almacenes.Rows = 0
'Consulta los datos generales de Cat_Almacenes
Mi_SQL = "SELECT Almacen_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Almacenes "
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Almacen_ID"
Set Rs_Consulta_Cat_Almacenes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Almacenes.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Almacenes.AddItem "Almcen ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Almacenes.EOF
        Grid_Cat_Almacenes.AddItem Rs_Consulta_Cat_Almacenes.rdoColumns("aLMACEN_ID") _
        & Chr(9) & Rs_Consulta_Cat_Almacenes.rdoColumns("Nombre")
        Grid_Cat_Almacenes.FixedRows = 1
        Rs_Consulta_Cat_Almacenes.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Almacenes.ColWidth(0) = 1000
    Grid_Cat_Almacenes.ColAlignment(0) = 3
    Grid_Cat_Almacenes.ColWidth(1) = 6629
    Grid_Cat_Almacenes.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Almacenes.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Almacenes
'DESCRIPCIÓN                : Modifica los Almacenes
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 26 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Almacenes()
Dim Rs_Modificacion_Cat_Almacenes As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta los Almacenes
    Mi_SQL = "SELECT * FROM Cat_Almacenes"
    Mi_SQL = Mi_SQL & " WHERE Almacen_ID ='" & Trim(Txt_Almacen_ID.text) & "'"
    Set Rs_Modificacion_Cat_Almacenes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Almacenes
    With Rs_Modificacion_Cat_Almacenes
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Almacenes.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Almacenes.text))
            If Trim(Txt_Comentarios_Cat_Almacenes.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Almacenes.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Almacenes.Close
    Grid_Cat_Almacenes.TextMatrix(Grid_Cat_Almacenes.RowSel, 0) = Trim(UCase(Txt_Almacen_ID.text))
    Grid_Cat_Almacenes.TextMatrix(Grid_Cat_Almacenes.RowSel, 1) = Trim(Txt_Nombre_Cat_Almacenes.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Cat_Almacenes.Enabled = False
    Fra_Cat_Almacenes_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Almacenes", Frm_Cat_Generales)
    MsgBox "El Almacen ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Productos_Tipo
'DESCRIPCIÓN                : Da de alta los Tipo de Productos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Productos_Tipo()
Dim Rs_Alta_Cat_Productos_Tipo As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Tipos de Producto
    Set Rs_Alta_Cat_Productos_Tipo = Conectar_Ayudante.Recordset_Agregar("Cat_Productos_Tipo")
    'Llena la tabla de Cat_Productos_Tipo
    With Rs_Alta_Cat_Productos_Tipo
        .AddNew
            .rdoColumns("Tipo_ID") = Trim(Txt_Tipo_ID.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Productos_Tipo.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Productos_Tipo.text))
            If Trim(Txt_Comentarios_Cat_Productos_Tipo.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Productos_Tipo.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Productos_Tipo.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Productos_Tipo.Enabled = False
    Fra_Cat_Produstos_Tipo_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Cat_Productos_Tipo.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Productos_Tipo.Rows = 0 Then
        Grid_Cat_Productos_Tipo.AddItem "Tipo ID" & Chr(9) & "Nombre"
        Grid_Cat_Productos_Tipo.AddItem Trim(Txt_Tipo_ID.text) & Chr(9) & _
        UCase(Txt_Nombre_Cat_Productos_Tipo.text)
        Grid_Cat_Productos_Tipo.ColWidth(0) = 1000
        Grid_Cat_Productos_Tipo.ColAlignment(0) = 3
        Grid_Cat_Productos_Tipo.ColWidth(1) = 6629
        Grid_Cat_Productos_Tipo.ColAlignment(1) = 2
        Grid_Cat_Productos_Tipo.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Productos_Tipo("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos_Tipo", Frm_Cat_Generales)
    MsgBox "Tipo Producto dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Cat_Productos_Tipo
'DESCRIPCIÓN                : Consulta los Tipos de Producto
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Productos_Tipo(Nombre As String)
Dim Rs_Consulta_Cat_Productos_Tipo As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Cat_Productos_Tipo.Rows = 0
'Consulta los datos generales del la clasificacion de proveedor
Mi_SQL = "SELECT Tipo_ID, Nombre"
Mi_SQL = Mi_SQL & " FROM Cat_Productos_Tipo"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Tipo_ID"
Set Rs_Consulta_Cat_Productos_Tipo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Productos_Tipo.EOF Then
    'Coloca un encabezado en el grid
    Grid_Cat_Productos_Tipo.AddItem "Tipo ID" & Chr(9) & "Nombre"
    While Not Rs_Consulta_Cat_Productos_Tipo.EOF
        Grid_Cat_Productos_Tipo.AddItem Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Tipo_ID") _
        & Chr(9) & Rs_Consulta_Cat_Productos_Tipo.rdoColumns("Nombre")
        Grid_Cat_Productos_Tipo.FixedRows = 1
        Rs_Consulta_Cat_Productos_Tipo.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Cat_Productos_Tipo.ColWidth(0) = 1000
    Grid_Cat_Productos_Tipo.ColAlignment(0) = 3
    Grid_Cat_Productos_Tipo.ColWidth(1) = 6629
    Grid_Cat_Productos_Tipo.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Productos_Tipo.Close
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Productos_Tipo
'DESCRIPCIÓN                : Modifica los TipoProductos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Productos_Tipo()
Dim Rs_Modificacion_Cat_Productos_Tipo As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Tipo actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Productos_Tipo"
    Mi_SQL = Mi_SQL & " WHERE Tipo_ID ='" & Trim(Txt_Tipo_ID.text) & "'"
    Set Rs_Modificacion_Cat_Productos_Tipo = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Productos_Tipo
    With Rs_Modificacion_Cat_Productos_Tipo
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Productos_Tipo.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Productos_Tipo.text))
            If Trim(Txt_Comentarios_Cat_Productos_Tipo.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Productos_Tipo.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Productos_Tipo.Close
    Grid_Cat_Productos_Tipo.TextMatrix(Grid_Cat_Productos_Tipo.RowSel, 0) = Trim(UCase(Txt_Tipo_ID.text))
    Grid_Cat_Productos_Tipo.TextMatrix(Grid_Cat_Productos_Tipo.RowSel, 1) = Trim(Txt_Nombre_Cat_Productos_Tipo.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Cat_Productos_Tipo.Enabled = False
    Fra_Cat_Produstos_Tipo_Detalles.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos_Tipo", Frm_Cat_Generales)
    MsgBox "El Tipo Producto ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Impuestos
'DESCRIPCIÓN                : Da de alta los Impuestos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 29 OCTUBRE 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Impuestos()
Dim Rs_Alta_Cat_Impuestos As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  Impùestos
    Set Rs_Alta_Cat_Impuestos = Conectar_Ayudante.Recordset_Agregar("Cat_Impuestos")
    'Llena la tabla de Cat_Impuestos
    With Rs_Alta_Cat_Impuestos
        .AddNew
            .rdoColumns("Impuesto_ID") = Trim(Txt_ID_Cat_Impuestos.text)
            .rdoColumns("Impuesto") = Trim(UCase(Txt_Impuesto_Cat_Impuestos.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Impuestos.text))
            If Trim(Txt_Comentarios_Cat_Impuestos.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Impuestos.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Impuestos.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Generales_Cta_Impuestos.Enabled = False
    Fra_Detalles_Cat_Impuestos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Estatus_Cat_Impuestos.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Impuestos_Cat_Impuestos.Rows = 0 Then
        Grid_Impuestos_Cat_Impuestos.AddItem "Impuesto ID" & Chr(9) & "Impuesto"
        Grid_Impuestos_Cat_Impuestos.AddItem Trim(Txt_ID_Cat_Impuestos.text) & Chr(9) & _
        UCase(Txt_Impuesto_Cat_Impuestos.text)
        Grid_Impuestos_Cat_Impuestos.ColWidth(0) = 1000
        Grid_Impuestos_Cat_Impuestos.ColAlignment(0) = 3
        Grid_Impuestos_Cat_Impuestos.ColWidth(1) = 6629
        Grid_Impuestos_Cat_Impuestos.ColAlignment(1) = 2
        Grid_Impuestos_Cat_Impuestos.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Impuestos("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Impuestos", Frm_Cat_Generales)
    MsgBox "Impuesto dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Cat_Impuestos
'DESCRIPCIÓN                : Consulta los Impuestos
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 29 Octubre 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Impuestos(Nombre As String)
Dim Rs_Consulta_Cat_Impuestos As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
Grid_Impuestos_Cat_Impuestos.Rows = 0
'Consulta los datos generales del Impuesto
Mi_SQL = "SELECT Impuesto_ID, Impuesto"
Mi_SQL = Mi_SQL & " FROM Cat_Impuestos"
Mi_SQL = Mi_SQL & " WHERE Impuesto LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Impuesto"
Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Cat_Impuestos.EOF Then
    'Coloca un encabezado en el grid
    Grid_Impuestos_Cat_Impuestos.AddItem "Impuesto ID" & Chr(9) & "Impuesto"
    While Not Rs_Consulta_Cat_Impuestos.EOF
        Grid_Impuestos_Cat_Impuestos.AddItem Rs_Consulta_Cat_Impuestos.rdoColumns("Impuesto_ID") _
        & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Impuesto")
        Grid_Impuestos_Cat_Impuestos.FixedRows = 1
        Rs_Consulta_Cat_Impuestos.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid
    Grid_Impuestos_Cat_Impuestos.ColWidth(0) = 1000
    Grid_Impuestos_Cat_Impuestos.ColAlignment(0) = 3
    Grid_Impuestos_Cat_Impuestos.ColWidth(1) = 6629
    Grid_Impuestos_Cat_Impuestos.ColAlignment(1) = 2
  
End If
'Cierra el manejador del registro
Rs_Consulta_Cat_Impuestos.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Cat_Impuestos
'DESCRIPCIÓN                : Modifica los Impuestos
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 29 Octubre 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Cat_Impuestos()
Dim Rs_Modificacion_Cat_Impuestos As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Impuesto actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Impuestos"
    Mi_SQL = Mi_SQL & " WHERE Impuesto_ID ='" & Trim(Txt_ID_Cat_Impuestos.text) & "'"
    Set Rs_Modificacion_Cat_Impuestos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Impuestos
    With Rs_Modificacion_Cat_Impuestos
        .Edit
            .rdoColumns("Impuesto_ID") = Trim(Txt_ID_Cat_Impuestos.text)
            .rdoColumns("Impuesto") = Trim(UCase(Txt_Impuesto_Cat_Impuestos.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Estatus_Cat_Impuestos.text))
            If Trim(Txt_Comentarios_Cat_Impuestos.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Impuestos.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Impuestos.Close
    Grid_Impuestos_Cat_Impuestos.TextMatrix(Grid_Impuestos_Cat_Impuestos.RowSel, 0) = Trim(UCase(Txt_ID_Cat_Impuestos.text))
    Grid_Impuestos_Cat_Impuestos.TextMatrix(Grid_Impuestos_Cat_Impuestos.RowSel, 1) = Trim(Txt_Impuesto_Cat_Impuestos.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_Cta_Impuestos.Enabled = False
    Fra_Detalles_Cat_Impuestos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Impuestos", Frm_Cat_Generales)
    MsgBox "El Impuesto ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Alta_Cat_Sustancia_Activa
'DESCRIPCIÓN                : Da de alta la sustancia activa
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 29 OCTUBRE 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Alta_Cat_Sustancia_Activa()
Dim Rs_Alta_Cat_Sustancia_Activa As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de  sustancia activa
    Set Rs_Alta_Cat_Sustancia_Activa = Conectar_Ayudante.Recordset_Agregar("Cat_Sustancia_Activa")
    'Llena la tabla de Cat_Sustancia_Activa
    With Rs_Alta_Cat_Sustancia_Activa
        .AddNew
            .rdoColumns("Sustancia_Activa_ID") = Trim(Txt_ID_Cat_Sustancia_Activa.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Sustancia_Activa.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Cat_Sustancia_Activa.text))
            If Trim(Txt_Comentarios_Cat_Sustancia_Activa.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Sustancia_Activa.text))
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Sustancia_Activa.Close
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Sustancia_Activa.Enabled = False
    Fra_Detalles_Cat_Sustancia_Activa.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Cmb_Cat_Sustancia_Activa.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Sustancia_Activa.Rows = 0 Then
        Grid_Cat_Sustancia_Activa.AddItem "Sustancia Activa ID" & Chr(9) & "Nombre"
        Grid_Cat_Sustancia_Activa.AddItem Trim(Txt_ID_Cat_Sustancia_Activa.text) & Chr(9) & _
        UCase(Txt_Nombre_Cat_Sustancia_Activa.text)
        Grid_Cat_Sustancia_Activa.ColWidth(0) = 1000
        Grid_Cat_Sustancia_Activa.ColAlignment(0) = 3
        Grid_Cat_Sustancia_Activa.ColWidth(1) = 6629
        Grid_Cat_Sustancia_Activa.ColAlignment(1) = 2
        Grid_Cat_Sustancia_Activa.FixedRows = 1
     End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Sustancia_Activa("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Sustancia_Activa", Frm_Cat_Generales)
    MsgBox "Sustancia Activa dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Cat_Sustancia_Activa
'DESCRIPCIÓN                : Consulta los Sustancias activa
'PARÁMETROS                 : Nombre; sirve para hacer laconsulta usandolo como criterio de busqueda
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Cat_Sustancia_Activa(Nombre As String)
Dim Rs_Consulta_Cat_Sustancia_Activa As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
    Grid_Cat_Sustancia_Activa.Rows = 0
    'Consulta los datos generales
    Mi_SQL = "SELECT Sustancia_Activa_ID, Nombre"
    Mi_SQL = Mi_SQL & " FROM Cat_Sustancia_Activa"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Sustancia_Activa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Sustancia_Activa.EOF Then
        'Coloca un encabezado en el grid
        Grid_Cat_Sustancia_Activa.AddItem "Sustancia Activa ID" & Chr(9) & "Nombre"
        While Not Rs_Consulta_Cat_Sustancia_Activa.EOF
            Grid_Cat_Sustancia_Activa.AddItem Rs_Consulta_Cat_Sustancia_Activa.rdoColumns("Sustancia_Activa_ID") _
            & Chr(9) & Rs_Consulta_Cat_Sustancia_Activa.rdoColumns("Nombre")
            Grid_Cat_Sustancia_Activa.FixedRows = 1
            Rs_Consulta_Cat_Sustancia_Activa.MoveNext
        Wend
        'Configura el tamaño de las columnas del grid
        Grid_Cat_Sustancia_Activa.ColWidth(0) = 1000
        Grid_Cat_Sustancia_Activa.ColAlignment(0) = 3
        Grid_Cat_Sustancia_Activa.ColWidth(1) = 6629
        Grid_Cat_Sustancia_Activa.ColAlignment(1) = 2
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Sustancia_Activa.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Modifica_Sustancia_Activa
'DESCRIPCIÓN                : Modifica la sustancia Activa
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30 Agosto 2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Modifica_Sustancia_Activa()
Dim Rs_Modificacion_Cat_Sustancia_Activa As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo handler
    Conexion_Base.BeginTrans
    'Consulta el Banco actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Sustancia_Activa"
    Mi_SQL = Mi_SQL & " WHERE Sustancia_Activa_ID ='" & Trim(Txt_ID_Cat_Sustancia_Activa.text) & "'"
    Set Rs_Modificacion_Cat_Sustancia_Activa = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Bancos
    With Rs_Modificacion_Cat_Sustancia_Activa
        .Edit
            .rdoColumns("Sustancia_Activa_ID") = Trim(Txt_ID_Cat_Sustancia_Activa.text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Cat_Sustancia_Activa.text))
            .rdoColumns("Estatus") = Trim(UCase(Cmb_Cat_Sustancia_Activa.text))
            If Trim(Txt_Comentarios_Cat_Sustancia_Activa.text) = "" Then
                .rdoColumns("Comentarios") = " "
            Else
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Cat_Sustancia_Activa.text))
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Sustancia_Activa.Close
    Grid_Cat_Sustancia_Activa.TextMatrix(Grid_Cat_Sustancia_Activa.RowSel, 0) = Trim(UCase(Txt_ID_Cat_Sustancia_Activa.text))
    Grid_Cat_Sustancia_Activa.TextMatrix(Grid_Cat_Sustancia_Activa.RowSel, 1) = Trim(Txt_Nombre_Cat_Sustancia_Activa.text)
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Cat_Sustancia_Activa.Enabled = False
    Fra_Detalles_Cat_Sustancia_Activa.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Sustancia_Activa", Frm_Cat_Generales)
    MsgBox "La Sustancia Activa ha sido modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
