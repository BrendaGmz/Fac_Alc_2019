VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Clientes_Facturas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas  Clientes"
   ClientHeight    =   8010
   ClientLeft      =   2310
   ClientTop       =   405
   ClientWidth     =   14355
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8467.229
   ScaleMode       =   0  'User
   ScaleWidth      =   14386.29
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra_Busqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      TabIndex        =   75
      Top             =   5856
      Visible         =   0   'False
      Width           =   8355
      Begin VB.Frame Fra_Busqueda_Con_Controles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busqueda "
         Height          =   960
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   8175
         Begin VB.OptionButton Opt_Nota_Cargo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nota Cargo"
            Height          =   195
            Left            =   1440
            TabIndex        =   90
            Top             =   315
            Width           =   1575
         End
         Begin VB.ComboBox Cmb_Consulta_Tipo_Factura 
            Height          =   315
            ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":0000
            Left            =   3000
            List            =   "Frm_Adm_Facturacion_Clientes.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   300
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.CommandButton Btn_Cerrar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cerrar"
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
            Left            =   7080
            Picture         =   "Frm_Adm_Facturacion_Clientes.frx":0022
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "A"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton Opt_Remision 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remision"
            Height          =   195
            Left            =   165
            TabIndex        =   30
            Top             =   630
            Width           =   1050
         End
         Begin VB.OptionButton Opt_Factura 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Factura"
            Height          =   195
            Left            =   165
            TabIndex        =   28
            Top             =   315
            Width           =   1095
         End
         Begin VB.CommandButton Btn_Sincronizar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aceptar"
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
            Left            =   5880
            Picture         =   "Frm_Adm_Facturacion_Clientes.frx":36B9
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "A"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   990
         End
      End
   End
   Begin VB.Frame Fra_Detalle_Factura 
      BackColor       =   &H8000000E&
      Caption         =   "Detalle del Documento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   44
      Top             =   3120
      Width           =   14295
      Begin VB.ComboBox Cmb_Descripcion_Sat 
         Height          =   315
         Left            =   6840
         TabIndex        =   97
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox Cmb_Unidad 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":5F88
         Left            =   915
         List            =   "Frm_Adm_Facturacion_Clientes.frx":5F98
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Txt_Presio_Sin_IVA 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   13440
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   855
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Fra_Totales 
         BackColor       =   &H8000000E&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1179
         Left            =   11130
         TabIndex        =   65
         Top             =   2685
         Width           =   3030
         Begin VB.TextBox Txt_IVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   480
            Width           =   2160
         End
         Begin VB.TextBox Txt_Subtotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   165
            Width           =   2160
         End
         Begin VB.TextBox Txt_Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   795
            Width           =   2160
         End
         Begin VB.Label Lbl_Subtotal 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   75
            TabIndex        =   71
            Top             =   210
            Width           =   585
         End
         Begin VB.Label Lbl_IVA 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "I.V.A."
            Height          =   195
            Left            =   75
            TabIndex        =   70
            Top             =   525
            Width           =   390
         End
         Begin VB.Label Lbl_Total 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Total"
            Height          =   195
            Left            =   75
            TabIndex        =   69
            Top             =   840
            Width           =   360
         End
      End
      Begin VB.TextBox Txt_Aplica_IVA 
         Height          =   315
         Left            =   2610
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton Btn_Agregar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frm_Adm_Facturacion_Clientes.frx":5FBC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Txt_Cantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   105
         TabIndex        =   20
         Top             =   480
         Width           =   810
      End
      Begin VB.TextBox Txt_Precio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11040
         TabIndex        =   23
         Top             =   480
         Width           =   1050
      End
      Begin VB.TextBox Txt_Importe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   480
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Detalle_Factura 
         Height          =   1815
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.ComboBox Cmb_Descripcion 
         Height          =   315
         Left            =   2640
         TabIndex        =   22
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox Txt_No_Salida 
         Height          =   285
         Left            =   13320
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   443
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox Text_Impuesto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   165
         TabIndex        =   63
         Top             =   480
         Width           =   705
      End
      Begin VB.Frame Fra_Comentarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comentarios"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1365
         TabIndex        =   61
         Top             =   2700
         Width           =   6255
         Begin VB.TextBox Txt_Comentarios 
            Height          =   330
            Left            =   30
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   180
            Width           =   6195
         End
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
         Height          =   420
         Left            =   90
         Picture         =   "Frm_Adm_Facturacion_Clientes.frx":9272
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2790
         Width           =   1260
      End
      Begin VB.Label lbl_estatus_cancel 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   3840
         Visible         =   0   'False
         Width           =   10815
      End
      Begin VB.Label Lbl_Descripcion_Sat 
         BackColor       =   &H80000005&
         Caption         =   "Descripción SAT"
         Height          =   255
         Left            =   8280
         TabIndex        =   98
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Unidad"
         Height          =   195
         Left            =   1440
         TabIndex        =   80
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Lbl_Cantidad 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   330
         TabIndex        =   56
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Lbl_Descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Descripción"
         Height          =   195
         Left            =   4200
         TabIndex        =   55
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Lbl_Precio 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Precio"
         Height          =   195
         Left            =   11280
         TabIndex        =   54
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Lbl_Importe 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Importe"
         Height          =   195
         Left            =   12360
         TabIndex        =   53
         Top             =   240
         Width           =   525
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Relacionados 
      Height          =   780
      Left            =   10920
      TabIndex        =   111
      Top             =   1080
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1376
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      Enabled         =   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.CommandButton Btn_Enviar_Email 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enviar Email"
      Enabled         =   0   'False
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
      Left            =   7800
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":C524
      Style           =   1  'Graphical
      TabIndex        =   88
      Tag             =   "A"
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Buscar 
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
      Left            =   10320
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":CAAE
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "C"
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Cancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
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
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":1003A
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "A"
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Imprimir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
      Enabled         =   0   'False
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
      Left            =   2640
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":136D1
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "A"
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   1350
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
      Left            =   120
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":16B97
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "A"
      Top             =   7140
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
      Left            =   12840
      Picture         =   "Frm_Adm_Facturacion_Clientes.frx":1A0CE
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "A"
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Datos_Cliente 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos del Cliente"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   30
      TabIndex        =   38
      Top             =   30
      Width           =   6555
      Begin VB.TextBox Txt_Cuenta_Pago 
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   109
         Top             =   2520
         Width           =   1560
      End
      Begin VB.ComboBox Cmb_Forma_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1D7CD
         Left            =   1110
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1D7CF
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2520
         Width           =   3015
      End
      Begin VB.ComboBox Cmb_Uso_CFDI 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1D7D1
         Left            =   1110
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1D7D3
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   2160
         Width           =   5370
      End
      Begin VB.PictureBox Pic_Logotipo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1545
         Left            =   0
         Picture         =   "Frm_Adm_Facturacion_Clientes.frx":1D7D5
         ScaleHeight     =   1485
         ScaleWidth      =   1470
         TabIndex        =   89
         Top             =   0
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox Txt_Email 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1835
         Width           =   5370
      End
      Begin VB.TextBox Txt_Estado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1513
         Width           =   1095
      End
      Begin VB.TextBox Txt_Pais 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2535
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1513
         Width           =   975
      End
      Begin VB.TextBox Txt_Codigo_Postal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2535
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox Txt_Colonia_Cliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1191
         Width           =   2415
      End
      Begin VB.TextBox Txt_No_Interior 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5780
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   869
         Width           =   700
      End
      Begin VB.TextBox Txt_No_Exterior 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4932
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   869
         Width           =   820
      End
      Begin VB.TextBox Txt_Direccion_Cliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   869
         Width           =   3810
      End
      Begin VB.TextBox Txt_Telefono_Cliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5175
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1513
         Width           =   1305
      End
      Begin VB.TextBox Txt_Dias_Credito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4425
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1513
         Width           =   480
      End
      Begin VB.ComboBox Cmb_Nombre_Cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1110
         TabIndex        =   0
         Top             =   195
         Width           =   5370
      End
      Begin VB.TextBox Txt_RFC_Cliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox Txt_Ciudad_Cliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1191
         Width           =   2055
      End
      Begin VB.TextBox Txt_Cliente_ID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   547
         Width           =   1095
      End
      Begin VB.Label Lbl_Cta_Pago 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cta Pago"
         Height          =   195
         Left            =   4200
         TabIndex        =   110
         Top             =   2595
         Width           =   660
      End
      Begin VB.Label Lbl_Forma_Pago 
         BackColor       =   &H80000005&
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   2595
         Width           =   900
      End
      Begin VB.Label Lbl_Uso_CFDI 
         BackColor       =   &H80000005&
         Caption         =   "Uso de CFDI"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Email"
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Pais"
         Height          =   195
         Left            =   2212
         TabIndex        =   83
         Top             =   1558
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Estado"
         Height          =   195
         Left            =   105
         TabIndex        =   82
         Top             =   1558
         Width           =   495
      End
      Begin VB.Label Lbl_CP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "CP"
         Height          =   195
         Left            =   2220
         TabIndex        =   60
         Top             =   585
         Width           =   210
      End
      Begin VB.Label Lbl_Telefono_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Tel"
         Height          =   195
         Left            =   4920
         TabIndex        =   58
         Top             =   1558
         Width           =   225
      End
      Begin VB.Label Lbl_Dias_Credito 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Días Créd"
         Height          =   195
         Left            =   3675
         TabIndex        =   57
         Top             =   1558
         Width           =   720
      End
      Begin VB.Label Lbl_RFC_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "R.F.C."
         Height          =   195
         Left            =   3675
         TabIndex        =   50
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Lbl_Ciudad_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   3675
         TabIndex        =   49
         Top             =   1236
         Width           =   495
      End
      Begin VB.Label Lbl_Colonia_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Colonia"
         Height          =   195
         Left            =   105
         TabIndex        =   48
         Top             =   1236
         Width           =   525
      End
      Begin VB.Label Lbl_Direccion_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Dirección"
         Height          =   195
         Left            =   105
         TabIndex        =   47
         Top             =   914
         Width           =   675
      End
      Begin VB.Label Lbl_Nombre_Cliente 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Nombre"
         Height          =   195
         Left            =   105
         TabIndex        =   46
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Lbl_Cliente_ID 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente ID"
         Height          =   195
         Left            =   105
         TabIndex        =   45
         Top             =   585
         Width           =   690
      End
   End
   Begin VB.Frame Fra_Datos_Factura 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos del Documento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   6600
      TabIndex        =   37
      Top             =   30
      Width           =   7740
      Begin VB.ComboBox Cmb_Metodo_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5AB
         Left            =   1440
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5AD
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox Txt_Plazo_Pago 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   114
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox Cmb_Tipo_Adenda 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5AF
         Left            =   5400
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5B9
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   1950
         Width           =   2175
      End
      Begin VB.TextBox Txt_Orden_Compra 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   103
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox Chc_Adenda 
         BackColor       =   &H80000014&
         Caption         =   "Adenda"
         Height          =   315
         Left            =   4320
         TabIndex        =   102
         Top             =   2000
         Width           =   855
      End
      Begin VB.ComboBox Cmb_Serie 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5CB
         Left            =   5450
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5CD
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Cmb_FacRef 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5CF
         Left            =   6285
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5D1
         TabIndex        =   99
         Text            =   "Cmb_FacRef"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Txt_Condicion_Pago 
         Height          =   285
         Left            =   1440
         TabIndex        =   96
         Top             =   2400
         Width           =   2775
      End
      Begin VB.ComboBox Cmb_Relacionados 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5D3
         Left            =   5450
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5D5
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   405
         Width           =   2175
      End
      Begin VB.CheckBox Chc_Relacionados 
         BackColor       =   &H80000014&
         Caption         =   "Comprobantes relacionados"
         Height          =   315
         Left            =   4320
         TabIndex        =   91
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox Txt_Factura_ID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   520
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Txt_Serie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   860
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Tipo_Factura 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5D7
         Left            =   1215
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5E1
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   520
         Width           =   2940
      End
      Begin VB.ComboBox Cmb_Salidas 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":1F5F9
         Left            =   600
         List            =   "Frm_Adm_Facturacion_Clientes.frx":1F5FB
         TabIndex        =   16
         Top             =   1513
         Width           =   1335
      End
      Begin VB.CommandButton Btn_Agregar_Salida 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agregar Salida"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frm_Adm_Facturacion_Clientes.frx":1F5FD
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1493
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_Tipo_Documento 
         Height          =   315
         ItemData        =   "Frm_Adm_Facturacion_Clientes.frx":228B3
         Left            =   1215
         List            =   "Frm_Adm_Facturacion_Clientes.frx":228C0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   2940
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_Factura 
         Height          =   300
         Left            =   600
         TabIndex        =   18
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   3801091
         CurrentDate     =   38712
      End
      Begin VB.TextBox Txt_No_Factura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   860
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_Pago 
         Height          =   300
         Left            =   2880
         TabIndex        =   19
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   3801091
         CurrentDate     =   38712
      End
      Begin VB.TextBox Txt_UUID_Relacion 
         Height          =   285
         Left            =   6840
         TabIndex        =   100
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Lbl_Metodo_Pago 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Método Pago"
         Height          =   195
         Left            =   75
         TabIndex        =   117
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Plazo Pago"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4320
         TabIndex        =   115
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Orden Compra"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4320
         TabIndex        =   104
         Top             =   2430
         Width           =   1260
      End
      Begin VB.Label Lbl_Condicion_Pago 
         BackColor       =   &H80000005&
         Caption         =   "Condición Pago"
         Height          =   255
         Left            =   75
         TabIndex        =   95
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label Lbl_UUID_Relacion 
         BackColor       =   &H80000005&
         Caption         =   "Relacionado"
         Height          =   255
         Left            =   4320
         TabIndex        =   94
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Lbl_Tipo_Relacion 
         BackColor       =   &H80000005&
         Caption         =   "Tipo Relación"
         Height          =   255
         Left            =   4320
         TabIndex        =   92
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Tipo Factura"
         Height          =   195
         Left            =   80
         TabIndex        =   79
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salida"
         Height          =   195
         Left            =   75
         TabIndex        =   78
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento"
         Height          =   195
         Left            =   80
         TabIndex        =   77
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Lbl_Facturacion 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   15
         Top             =   1185
         Width           =   4095
      End
      Begin VB.Label Lbl_Fecha_Pago 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Fecha  pago"
         Height          =   195
         Left            =   1920
         TabIndex        =   59
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Lbl_Fecha_Factura 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Fecha"
         Height          =   195
         Left            =   75
         TabIndex        =   52
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Lbl_No_Factura 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "No. Documento"
         Height          =   195
         Left            =   80
         TabIndex        =   51
         Top             =   900
         Width           =   1125
      End
   End
   Begin VB.Frame Fra_Datos_Remision 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos Remisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   72
      Top             =   6120
      Width           =   7530
      Begin VB.TextBox Txt_Nombre_Almacenista 
         Height          =   330
         Left            =   5040
         TabIndex        =   35
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox Txt_Recibe 
         Height          =   330
         Left            =   1305
         TabIndex        =   34
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Lbl_Nombre_Almacenista 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre del almacenista"
         Height          =   195
         Left            =   3240
         TabIndex        =   74
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Lbl_Recibe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre Recibe"
         Height          =   195
         Left            =   90
         TabIndex        =   73
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Frm_Adm_Clientes_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UM As String
 
Private Sub Formatea_Columnas_Grid()
    If Grid_Detalle_Factura.Rows > 1 Then
        Grid_Detalle_Factura.FixedRows = 1
        Grid_Detalle_Factura.ColWidth(0) = 800      'Cantidad
        Grid_Detalle_Factura.ColWidth(1) = 3000    'Unidad
        Grid_Detalle_Factura.ColWidth(2) = 5500     'Descripcion
        Grid_Detalle_Factura.ColAlignment(1) = 1
        Grid_Detalle_Factura.ColAlignment(2) = 1
        Grid_Detalle_Factura.ColAlignment(3) = 1
        Grid_Detalle_Factura.ColWidth(3) = 3000   'Clave producto SAT
        Grid_Detalle_Factura.ColWidth(4) = 1000     'Precio
        Grid_Detalle_Factura.ColWidth(5) = 1100     'Importe
        Grid_Detalle_Factura.ColWidth(6) = 0        'Producto_ID
        Grid_Detalle_Factura.ColWidth(7) = 0        'No_Salida
        Grid_Detalle_Factura.ColWidth(8) = 0        'Impuesto
        Grid_Detalle_Factura.ColWidth(9) = 0        '
        Grid_Detalle_Factura.ColWidth(10) = 0        'IVA
        Grid_Detalle_Factura.ColWidth(11) = 0        'aplica_IVA
        Grid_Detalle_Factura.ColWidth(12) = 0 '600    'Incluir
        If Cmb_Tipo_Adenda.text = "NADRO" Then
            Grid_Detalle_Factura.ColWidth(13) = 800
        End If
        'Grid_Detalle_Factura.ColAlignment(12) = 3
        
         
    Else
        Grid_Detalle_Factura.Rows = 0
    End If
    
End Sub

'Botón para agregar un registro en el grid de detalles de productos
Private Sub Btn_Agregar_Click()
Dim Cont_Detalles As Integer        'Usada para contar detalle del grid
Dim Suma As Double                  'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double              'Suma I.V.A.
Dim Posicion_OC As Integer
    
    
    If Trim(Cmb_Tipo_Documento.text) = "FACTURA" Or Trim(Cmb_Tipo_Documento.text) = "REMISION" Or Trim(Cmb_Tipo_Documento.text) = "NOTA CARGO" Then
        If Cmb_Tipo_Adenda.text = "NADRO" Then
            Posicion_OC = Val(InputBox("Teclee la posición del material en la orden de compra", "Posición en Orden de Compra"))
            If Val(Posicion_OC) = 0 Then
                MsgBox "Introduce la posición del producto en la orden de compra", vbCritical
                Exit Sub
            End If
        End If
        'Valida que los campos tengan valores
        If Val(Txt_Cantidad.text) >= 0 And (Cmb_Descripcion.text <> "" And Cmb_Descripcion_Sat.text <> "" And Val(Txt_Precio.text) >= 0 And Val(Txt_Importe.text) >= 0) Then
            If Cmb_Unidad.ListIndex = -1 Then
                MsgBox "Indique la unidad de medida", vbExclamation
                Cmb_Unidad.SetFocus
                Exit Sub
            End If
            If Grid_Detalle_Factura.Rows = 0 Then
                'Coloca el número de columnas
                Grid_Detalle_Factura.Cols = 14
                'Pone el encabezado en las columnas
'                Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) _
'                    & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) & "" & Chr(9) & "IVA" & Chr(9) _
'                    & "Aplica_IVA" & Chr(9) & "Unidad" & Chr(9) & "Incluir" & Chr(9) & "Codigo_Producto_Servicio"
                Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Unidad" & Chr(9) & "Descripción" & Chr(9) & "Descripción SAT" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) _
                    & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) & "" & Chr(9) & "IVA" & Chr(9) _
                    & "Aplica_IVA" & Chr(9) & "Incluir" & Chr(9) & "Posición OC"
            End If
            'Agrega el dato en el grid
            If Cmb_Descripcion.ListIndex > -1 And Trim(Cmb_Tipo_Documento.text) = "FACTURA" Then
                If Trim(Txt_Aplica_IVA.text) = "SI" Then
                    Grid_Detalle_Factura.AddItem Txt_Cantidad.text & Chr(9) & Cmb_Unidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) & Cmb_Descripcion_Sat.text & Chr(9) _
                        & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                        & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                        & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & Cmb_Salidas.text & Chr(9) _
                        & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) _
                        & Val(PG_Retencion_IVA) * ((Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")))) & Chr(9) _
                        & "SI" & Chr(9) & "SI" & Chr(9) & Posicion_OC
                Else
                    Grid_Detalle_Factura.AddItem Txt_Cantidad.text & Chr(9) & Cmb_Unidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) & Cmb_Descripcion_Sat.text & Chr(9) _
                        & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                        & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                        & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & Cmb_Salidas.text & Chr(9) _
                        & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) & "" & Chr(9) & "NO" & Chr(9) & "SI" & Chr(9) & Posicion_OC
                End If
            Else
                Grid_Detalle_Factura.AddItem Txt_Cantidad.text & Chr(9) & Cmb_Unidad.text & Chr(9) & UCase(Trim(Mid(Cmb_Descripcion.text, 1, 45))) & Chr(9) & Cmb_Descripcion_Sat.text & Chr(9) _
                    & Format(Val(Txt_Precio.text), "#,##0.00") & Chr(9) _
                    & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                    & "" & Chr(9) & Cmb_Salidas.text & Chr(9) & Val(Text_Impuesto.text) & Chr(9) _
                    & Val(PG_Retencion_IVA) * ((Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(0, ",")))) & Chr(9) _
                    & Val(PG_Retencion_IVA) * ((Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(0, ",")))) & Chr(9) _
                    & "SI" & Chr(9) & "NO" & Chr(9) & Posicion_OC
        
            End If
            Formatea_Columnas_Grid
            'Cacula los totales
            Suma = 0
            'Hace el recorrido de los datos del grid para hacer la suma
            For Cont_Detalles = 1 To Grid_Detalle_Factura.Rows - 1
                Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 5), ","))
                Suma_IVA = Suma_IVA + Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 10), ",")), "#0.00")
            Next Cont_Detalles
            'Asigna los resultados a los totales
            Txt_Subtotal.text = Format(Suma, "#,##0.00")
            Txt_IVA.text = Format(Val(Suma_IVA), "#,##0.00")
            Txt_Total.text = Format(Val(Suma) + Val(Suma_IVA), "#,##0.00")
            Btn_Agregar.Default = False
            'Llamada para limpiar los campos de productos
            Limpia_Datos_Producto
            Cmb_Descripcion.text = ""
            Cmb_Unidad.ListIndex = -1
            Cmb_Descripcion_Sat.text = ""
            Txt_Cantidad.SetFocus
            Btn_Buscar.Enabled = False
        Else
            MsgBox "Faltan datos para agregar", vbExclamation
        End If
    End If
End Sub

Private Sub Btn_Agregar_Salida_Click()
Dim Rs_Consulta_Salidas_Detalles As rdoResultset
Dim Rs_Consulta_Unidad As rdoResultset
Dim Suma_IVA As Double              'Suma I.V.A.
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
Dim Unidad As String

    Cmb_Descripcion.Clear
    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Cmb_Salidas.text & "' and Facturado = 'NO' "
    Set Rs_Consulta_Salidas_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Salidas_Detalles.EOF Then
        If Grid_Detalle_Factura.Rows = 0 Then
            'Coloca el número de columnas
            Grid_Detalle_Factura.Cols = 13
            'Pone el encabezado en las columnas
            Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) _
                & "Importe" & Chr(9) & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) _
                & "No_Salida" & Chr(9) & "IVA" & Chr(9) & "Aplica_IVA" & Chr(9) & "Unidad" & Chr(9) & "Incluir"
        End If
        With Rs_Consulta_Salidas_Detalles
            While Not Rs_Consulta_Salidas_Detalles.EOF
                Mi_SQL = "SELECT * FROM Cat_Productos"
                Mi_SQL = Mi_SQL & " WHERE Producto_ID = '" & Rs_Consulta_Salidas_Detalles!Producto_ID & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    'consulta el nombre de la presentacion asignada
                    Mi_SQL = "SELECT Presentacion_ID, Nombre FROM Cat_Presentaciones"
                    Mi_SQL = Mi_SQL & " WHERE Presentacion_ID = '" & Rs_Consulta!Presentacion_ID & "'"
                    Set Rs_Consulta_Unidad = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Unidad.EOF Then
                            Unidad = Rs_Consulta_Unidad!Nombre
                        Else
                            Unidad = ""
                        End If
                    Rs_Consulta_Unidad.Close
                    If Trim(Rs_Consulta!Aplica_IVA) = "SI" And Trim(Cmb_Tipo_Documento.text) = "FACTURA" Then
                        'Agrega el dato en el grid
                        Grid_Detalle_Factura.AddItem Rs_Consulta_Salidas_Detalles!Cantidad & Chr(9) & _
                            UCase(Trim(Rs_Consulta_Salidas_Detalles!Descripcion)) & Chr(9) & _
                            Format((Rs_Consulta!Costo), "#,##0.00") & Chr(9) & _
                            Format(Val(Rs_Consulta!Costo) * Val(Rs_Consulta_Salidas_Detalles!Cantidad), "#,##0.00") & Chr(9) & _
                            Rs_Consulta_Salidas_Detalles!Producto_ID & Chr(9) & "" & Chr(9) & Rs_Consulta!Impuesto & Chr(9) & _
                            Cmb_Salidas.text & Chr(9) & (Val(Rs_Consulta_Salidas_Detalles!Cantidad) * Val(Rs_Consulta!Costo)) * Val(PG_Retencion_IVA) & Chr(9) _
                            & "SI" & Chr(9) & Unidad & Chr(9) & "SI"
                    Else
                        'Agrega el dato en el grid
                        Grid_Detalle_Factura.AddItem Rs_Consulta_Salidas_Detalles!Cantidad & Chr(9) & _
                            UCase(Trim(Rs_Consulta_Salidas_Detalles!Descripcion)) & Chr(9) & _
                            Format((Rs_Consulta!Costo), "#,##0.00") & Chr(9) & _
                            Format(Val(Rs_Consulta!Costo) * Val(Rs_Consulta_Salidas_Detalles!Cantidad), "#,##0.00") & Chr(9) & _
                            Rs_Consulta_Salidas_Detalles!Producto_ID & Chr(9) & "" & Chr(9) & Rs_Consulta!Impuesto & Chr(9) & _
                            Cmb_Salidas.text & Chr(9) & "" & Chr(9) & "NO" & Chr(9) & Unidad & Chr(9) & "SI"
                    End If
                Rs_Consulta.Close
                Rs_Consulta_Salidas_Detalles.MoveNext
            Wend
        End With
        Rs_Consulta_Salidas_Detalles.Close
        Grid_Detalle_Factura.FixedRows = 1
        Formatea_Columnas_Grid
        'Cacula los totales
        Suma = 0
        'Hace el recorrido de los datos del grid para hacer la suma
        For Cont_Detalles = 1 To Grid_Detalle_Factura.Rows - 1
            Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 3), ","))
            Suma_IVA = Suma_IVA + Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 8), ",")))
        Next Cont_Detalles
        'Asigna los resultados a los totales
        Txt_Subtotal.text = Format(Suma, "#,##0.00")
        Txt_IVA.text = Format(Val(Suma_IVA), "#,##0.00")
        Txt_Total.text = Format(Val(Suma) + Val(Suma_IVA), "#,##0.00")
    End If
End Sub

'Botón para buscar una factura
Private Sub Btn_Buscar_Click()
    Fra_Busqueda.Visible = True
    Fra_Busqueda.Enabled = True
    Fra_Busqueda_Con_Controles.Enabled = True
End Sub

Private Sub Btn_Cancelar_Click()
    Resultado = MsgBox("¿Seguro de Cancelar el Documento?", vbYesNo + vbQuestion)
    If Resultado = 6 Then
        Cancela_Factura
    End If
End Sub

Private Sub Btn_Cerrar_Click()
    Fra_Busqueda.Visible = False
    Fra_Busqueda.Enabled = False
    Fra_Busqueda_Con_Controles.Enabled = False
End Sub

'Botón para eliminar un registro del grid de detalles de producto
Private Sub Btn_Eliminar_Click()
Dim Cont_Detalles As Integer            'Usada para contar detalles del grid
Dim Suma As Double                      'Usada para sumar el importe y manejo del I.V.A.
Dim Resp As Integer
Dim Suma_IVA As Double

    If Grid_Detalle_Factura.Rows > 1 Then
  
        Resp = MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbExclamation)
        If Resp = 6 Then
            'Si la respuesta es afirmativa elimina el registro seleccionado
            If Grid_Detalle_Factura.Rows = 2 Then
                Grid_Detalle_Factura.FixedRows = 0
                'Quita el item del grid
                Grid_Detalle_Factura.RemoveItem (Grid_Detalle_Factura.RowSel + 1)
                Btn_Buscar.Enabled = True
            Else
                If Grid_Detalle_Factura.Rows > 2 Then
                    Grid_Detalle_Factura.RemoveItem (Grid_Detalle_Factura.RowSel)
                End If
            End If
            Suma = 0
            Productos = 0
            
            'Hace el recorrido de los datos del grid para hacer la suma
            For Cont_Detalles = 1 To Grid_Detalle_Factura.Rows - 1
                Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 4), ","))
                Suma_IVA = Suma_IVA + Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 10), ",")), "#0.00")
            Next Cont_Detalles
            'Asigna los resultados a los totales
            Txt_Subtotal.text = Format(Suma, "#,##0.00")
            Txt_IVA.text = Format(Val(Suma_IVA), "#,##0.00")
            Txt_Total.text = Format(Val(Suma) + Val(Suma_IVA), "#,##0.00")
            Btn_Agregar.Default = False
            
            'Llamada para limpiar los campos de productos
            Call Limpia_Datos_Producto
            Cmb_Descripcion.text = ""
            Txt_Cantidad.SetFocus
            Btn_Buscar.Enabled = False
        End If
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprimir_Facturas
'DESCRIPCIÓN            : Imprime la factura de acuerdo a la configuración dada
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 30-Diciembre-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Imprimir_Facturas()
Dim Mi_SQL As String                            'Cadena para general las consultas
Dim Rs_Formato  As rdoResultset                 'Manejo de registro para la tabla Cfg_Formatos
Dim Rs_Formato_Generales  As rdoResultset       'Manejo de registro para la tabla Cfg_Formatos_Detalles
Dim Rs_Formato_Detalles As rdoResultset         'Manejo de registro para la tabla Cfg_Formatos_Detalles
Dim Rs_Consulta_Facturas As rdoResultset        'Manejo de registro para la tabla Descripcion_Facturas
Dim Rs_Consulta_Clave_Producto As rdoResultset  'Manejo de registro para la tabla Descripcion_Facturas
Dim Total_Le As String
Dim Longitud As Integer
Dim Inicio As Integer
Dim Salto As Double
Dim Cont_Renglon As Double
Dim Total_Vale As Double
Dim Precio As String
Dim Importe As String
Dim Rs_Consulta_Factura_Generales As rdoResultset
    
    For i = 1 To 1
        'Consulta para la configuración de facturas
        Mi_SQL = "SELECT * FROM Cfg_Formatos"
        Mi_SQL = Mi_SQL & " WHERE Nombre = 'FACTURA'"
        Set Rs_Formato = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Consulta para la configuración general de facturas
        Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Nombre = 'FACTURA'"
        Mi_SQL = Mi_SQL & " AND Tipo = 'General'"
        Set Rs_Formato_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Consulta para la configuración a detalle de facturas
        Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Nombre = 'FACTURA'"
        Mi_SQL = Mi_SQL & " AND Tipo = 'Detalle' ORDER BY Campo"
        Set Rs_Formato_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Impresión de la factura
        If Not Rs_Formato.EOF Then
            'Configura la fuente de la factura para generales
            With Rs_Formato
                Printer.ScaleMode = vbCentimeters
                Printer.FontSize = .rdoColumns("Tamaño_Generales")
                Printer.Font = .rdoColumns("Letra_Generales")
                If .rdoColumns("Estilo_Generales") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
            End With
            'Imprime los datos del cliente
            With Rs_Formato_Generales
                While Not .EOF
                    Printer.CurrentX = .rdoColumns("X")
                    Printer.CurrentY = .rdoColumns("Y")
                    Longitud = .rdoColumns("Longitud")
                    'Consulta el almacen de la factura
                    Mi_SQL = " SELECT  Almacen FROM Cat_Clientes "
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000") & "'"
                    Set Rs_Consulta_Factura_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Factura_Generales.EOF Then
                        If Not IsNull(Rs_Consulta_Factura_Generales!Almacen) And Len(Rs_Consulta_Factura_Generales!Almacen) > 0 Then
                            
                            If .rdoColumns("Campo") = "Cliente" Then
                               Printer.Print Cmb_Nombre_Cliente.text
                            End If
                            If Trim(.rdoColumns("Campo")) = "Domicilio" Then
                                Printer.Print Mid(Rs_Consulta_Factura_Generales!Almacen, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "CP" Then
                                 Printer.Print Mid("COL: " & Txt_Colonia_Cliente.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "Ciudad" Then
                                Printer.Print Mid(Txt_Ciudad_Cliente.text, 1, Longitud) & "  C.P. " & Mid(Txt_Codigo_Postal.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "Fecha_Dia" Then
                               Printer.Print Format(DTP_Fecha_Factura.Value, "d") & " de " & Format(DTP_Fecha_Factura.Value, "MMMM") & " del " & Format(DTP_Fecha_Factura.Value, "yyyy")
                            End If
                            If .rdoColumns("Campo") = "RFC" Then
                               Printer.Print Mid(Txt_RFC_Cliente.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "CANTIDAD_PAGARE" Then
                               Printer.Print "$" & Conectar_Ayudante.Alinea_Derecha(Format(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","), "#,##0.00"), 13) & " " & Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                            End If
                            If .rdoColumns("Campo") = "COLONIA" Then
                               Printer.Print Mid(Txt_Direccion_Cliente.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "Almacen_Entrega" Then
                               Printer.Print Mid(Txt_Ciudad_Cliente.text, 1, Longitud) & "  C.P. " & Mid(Txt_Codigo_Postal.text, 1, Longitud)
                            End If
                        Else 'IMPRESION NORMAL SIN EL CAMPO ALMACEN DE ENTREGA
                            If .rdoColumns("Campo") = "Cliente" Then
                               Printer.Print Cmb_Nombre_Cliente.text
                            End If
                            If Trim(.rdoColumns("Campo")) = "Domicilio" Then
                               Printer.Print Mid(Txt_Direccion_Cliente.text, 1, Longitud);
                            End If
                            If .rdoColumns("Campo") = "CP" Then
                               Printer.Print Mid(Txt_Ciudad_Cliente.text, 1, Longitud) & "  C.P. " & Mid(Txt_Codigo_Postal.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "Ciudad" Then
                               Printer.Print Mid(Txt_Ciudad_Cliente.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "Fecha_Dia" Then
                               Printer.Print Format(DTP_Fecha_Factura.Value, "d") & " de " & Format(DTP_Fecha_Factura.Value, "MMMM") & " del " & Format(DTP_Fecha_Factura.Value, "yyyy")
                            End If
                            If .rdoColumns("Campo") = "RFC" Then
                               Printer.Print Mid(Txt_RFC_Cliente.text, 1, Longitud)
                            End If
                            If .rdoColumns("Campo") = "CANTIDAD_PAGARE" Then
                               Printer.Print "$" & Conectar_Ayudante.Alinea_Derecha(Format(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","), "#,##0.00"), 13) & " " & Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                            End If
                            If .rdoColumns("Campo") = "COLONIA" Then
                               Printer.Print Mid("COL: " & Txt_Colonia_Cliente.text, 1, Longitud)
                            End If
                        End If
                    End If
                    Rs_Consulta_Factura_Generales.Close
                    'Imprime los totales
                    If .rdoColumns("Campo") = "Cantidad_Letras" Then
                        Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                    End If
                    If .rdoColumns("Campo") = "Subtotal" Then
                        Printer.Print "SUBTOTAL" & " $" & Chr(9) & Conectar_Ayudante.Alinea_Derecha(Format(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","), "#,##0.00"), 13)
                    End If
                    If .rdoColumns("Campo") = "IVA" Then
                        Printer.Print "IVA  16 %" & Chr(9) & "$" & Chr(9) & Conectar_Ayudante.Alinea_Derecha(Format(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","), "#,##0.00"), 13)
                    End If
                    If .rdoColumns("Campo") = "Total" Then
                        Printer.Print "TOTAL" & Chr(9) & "$" & Chr(9) & Conectar_Ayudante.Alinea_Derecha(Format(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","), "#,##0.00"), 13)
                    End If
                    If .rdoColumns("Campo") = "Fecha_Pagare" Then
                       Printer.Print Format(DTP_Fecha_Factura.Value, "dd MMM yyyy")
                    End If
                    If .rdoColumns("Campo") = "Cantidad_Letra_Pagare" Then
                        Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                    End If
                    .MoveNext
                Wend
            End With
            'Configura la fuente de la factura para detalles
            With Rs_Formato
                Printer.FontSize = .rdoColumns("Tamaño_Detalles")
                Printer.Font = .rdoColumns("Letra_Detalles")
                If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
            End With
            'Consulta de la tabla Adm_Descripcion_Facturas con el número de facturas
            Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(Txt_No_Factura.text, "0000000000") & "'"
            Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            'Configura la fuente para empresión de los detalles
            If Not Rs_Formato.EOF Then
                With Rs_Formato
                    Printer.ScaleMode = vbCentimeters
                    Printer.FontSize = .rdoColumns("Tamaño_Detalles")
                    Printer.Font = .rdoColumns("Letra_Detalles")
                    If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                        Printer.FontBold = True
                    Else
                        Printer.FontBold = False
                    End If
                    Salto = .rdoColumns("Separacion_Detalles")
                End With
                Cont_Renglon = 0
                'Imprime la factura
                While Not Rs_Consulta_Facturas.EOF
                    Cont_Renglon = Cont_Renglon + Salto
                    While Not Rs_Formato_Detalles.EOF
                        Printer.CurrentX = Rs_Formato_Detalles.rdoColumns("X")
                        Printer.CurrentY = Rs_Formato_Detalles.rdoColumns("Y") + Cont_Renglon
                        Longitud = Rs_Formato_Detalles.rdoColumns("Longitud")
                        
                        If Rs_Formato_Detalles.rdoColumns("Campo") = "Cantidad" Then
                            If Val(Rs_Consulta_Facturas.rdoColumns("Cantidad")) > 0 Then
                                Printer.Print Conectar_Ayudante.Alinea_Derecha(Rs_Consulta_Facturas.rdoColumns("Cantidad"), 5)
                            End If
                        End If
                        ''If Rs_Formato_Detalles.rdoColumns("Campo") = "Producto" Then Printer.Print Mid(Rs_Consulta_Facturas.rdoColumns("Descripcion"), 1, 45)
                        If Rs_Formato_Detalles.rdoColumns("Campo") = "CLAVE" Then
                            Mi_SQL = " SELECT Clave FROM Cat_Productos WHERE Producto_ID='" & Rs_Consulta_Facturas!Producto_ID & "'"
                            Set Rs_Consulta_Clave_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Clave_Producto.EOF Then
                                Printer.Print Rs_Consulta_Clave_Producto.rdoColumns("Clave")
                            End If
                            Rs_Consulta_Clave_Producto.Close
                        End If
                        If Val(Rs_Consulta_Facturas.rdoColumns("Cantidad")) > 0 Then
                            If Rs_Formato_Detalles.rdoColumns("Campo") = "Precio" Then
                                Precio = "$" & Format(Rs_Consulta_Facturas.rdoColumns("Precio"), "#,##0.00")
                                Printer.Print Conectar_Ayudante.Alinea_Derecha(Precio, 13)
                            End If
                            If Rs_Formato_Detalles.rdoColumns("Campo") = "Importe" Then
                                Importe = "$" & Format(Rs_Consulta_Facturas.rdoColumns("Importe"), "#,##0.00")
                                Printer.Print Conectar_Ayudante.Alinea_Derecha(Importe, 13)
                            End If
                        End If
                        If Rs_Formato_Detalles.rdoColumns("Campo") = "Producto" Then
                            Cont_Renglon = Imprime_Varias_Lineas(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Facturas.rdoColumns("Descripcion"), Chr(13)), Longitud, Rs_Formato_Detalles.rdoColumns("X"), Rs_Formato_Detalles.rdoColumns("Y"), Cont_Renglon, Salto + 0.1)
                        End If
                        Rs_Formato_Detalles.MoveNext
                    Wend
                    Rs_Formato_Detalles.MoveFirst
                    Rs_Consulta_Facturas.MoveNext
                Wend
                Rs_Consulta_Facturas.Close
            End If
            Printer.EndDoc
        End If
        Rs_Formato.Close
        Rs_Formato_Generales.Close
        Rs_Formato_Detalles.Close
    Next i
'    MsgBox "Factura enviada a Impresión", vbInformation
End Sub

Public Function Quitar_Enter(Texto As String, Caracter As String) As String
Dim posicion As Integer
posicion = 1
While posicion <> 0
    posicion = InStr(1, Texto, Caracter, vbTextCompare)
    If posicion <> 0 Then
        Texto = Mid(Texto, 1, posicion - 1) & " " & Mid(Texto, posicion + 2, Len(Texto))
    End If
Wend
Quitar_Enter = Texto
End Function

Function Imprime_Varias_Lineas(Real As String, Tamaño As Integer, X, Y, Contador_Renglon, Salto_Linea) As Double
Dim Ultima_Posicion As Integer
Dim Aux_Espacio As Integer
Dim Espacio As Integer
Dim Cadena As String
Dim Cortada As String

Ultima_Posicion = 1
Espacio = 1
Aux_Espacio = 1
Real = Real & Chr(13)
Printer.CurrentX = X
Printer.CurrentY = Y
                
Cadena = Mid(Real, Ultima_Posicion, Tamaño)
While Cadena <> ""
  ' Debug.Print Cadena
    Espacio = 0
    Aux_Espacio = 1
    While Aux_Espacio > 0
        Espacio = Aux_Espacio
        Aux_Espacio = InStr(Espacio + 1, Cadena, Chr(13), vbTextCompare)
        If Aux_Espacio = 0 Then
            Aux_Espacio = InStr(Espacio + 1, Cadena, " ", vbTextCompare)
        Else
            Espacio = Aux_Espacio + 1
            Aux_Espacio = 0
            Cadena = Mid(Cadena, 1, Espacio - 2)
        End If
    Wend
    If Espacio > 0 Then
        Printer.CurrentY = Y + Contador_Renglon
        Printer.CurrentX = X
        Printer.Print Mid(Cadena, 1, Espacio)
        Contador_Renglon = Contador_Renglon + Salto_Linea
    End If
    Ultima_Posicion = Ultima_Posicion + Espacio
    Cadena = Mid(Real, Ultima_Posicion, Tamaño)
Wend
Imprime_Varias_Lineas = Contador_Renglon
End Function

Private Sub Btn_Eliminar_Rel_Click()
    
End Sub

'Botón para imprimir, llama a la función Imprimir_Factura
Private Sub Btn_Imprimir_Click()
Dim Copias As Integer
Dim Impresiones As Integer
    
    Impresiones = 0
    Select Case Trim(Cmb_Tipo_Documento.text)
        Case "FACTURA"
            If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                Call Muestra_PDF("FACTURA")
            Else
                If MsgBox("¿Desea Imprimir la factura?", vbYesNo + vbInformation) = vbYes Then
                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
                    If Val(Impresiones) > 0 Then
                        For Copias = 1 To Impresiones
                            Call Imprimir_Facturas
                        Next
                    Else
                        MsgBox "Parámetro inválido", vbExclamation
                        Exit Sub
                    End If
                    MsgBox "Factura(s) enviada(s) a impresión", vbInformation
                End If
            End If
            If Cmb_Salidas.text <> "" Then
                Call Muestra_PDF("REMISION")
''                If MsgBox("¿Desea Imprimir la Remision?", vbYesNo + vbInformation) = vbYes Then
''                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
''                    If Val(Impresiones) > 0 Then
''                        For Copias = 1 To Impresiones
''                            Call Imprimir_Remision
''                        Next
''                        MsgBox "Remisión(es) enviada(s) a impresión", vbInformation
''                    Else
''                        MsgBox "Parámetro inválido", vbExclamation
''                        Exit Sub
''                    End If
''
''                End If
            End If
            Case "NOTA CARGO"
            If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                Call Muestra_PDF("NOTA CARGO")
            Else
                If MsgBox("¿Desea Imprimir la nota de cargo?", vbYesNo + vbInformation) = vbYes Then
                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
                    If Val(Impresiones) > 0 Then
                        For Copias = 1 To Impresiones
                            Call Imprimir_Facturas
                        Next
                    Else
                        MsgBox "Parámetro inválido", vbExclamation
                        Exit Sub
                    End If
                    MsgBox "Factura(s) enviada(s) a impresión", vbInformation
                End If
            End If
            If Cmb_Salidas.text <> "" Then
                Call Muestra_PDF("REMISION")
            End If
        Case "REMISION"
            
            Call Muestra_PDF("REMISION")
''            Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
''            If Val(Impresiones) > 0 Then
''                For Copias = 0 To Impresiones
''                    Call Imprimir_Remision
''                Next
''                MsgBox "Remisión(es) enviada(s) a impresión", vbInformation
''            Else
''                MsgBox "Parámetro inválido", vbExclamation
''                Exit Sub
''            End If
    End Select
End Sub

'Botón para agregar una nueva factura
Private Sub Btn_Nuevo_Click()
Set Conectar_Ayudante = New Ayudante  'Manejador del ayudante
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        'Habilita los controles para poder capturar la información y deshabilita otros
        Fra_Datos_Cliente.Enabled = True
        Fra_Datos_Factura.Enabled = True
        Fra_Detalle_Factura.Enabled = True
        Fra_Comentarios.Enabled = True
        Cmb_Nombre_Cliente.Enabled = True
        Cmb_Nombre_Cliente.text = ""
        Cmb_Nombre_Cliente_KeyPress 13
        Cmb_Nombre_Cliente.SetFocus
        Btn_Imprimir.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Nuevo.Visible = True
        Btn_Buscar.Enabled = False
        Btn_Cancelar.Enabled = False
        Btn_Enviar_Email.Enabled = False
        Btn_Salir.Caption = "Cancelar"
        Lbl_Facturacion.Caption = "Estatus"
        Btn_Nuevo.Caption = "Dar de Alta"       'Cambia el texto del botón
        DTP_Fecha_Pago.Value = Now              'Asigna la fecha al día actual
        DTP_Fecha_Factura.Value = Now
        Grid_Detalle_Factura.Rows = 0
        Call Conectar_Ayudante.Limpiar_Textos(Frm_Adm_Clientes_Facturas)
'        Txt_No_Factura.Text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
        Cmb_Tipo_Documento.ListIndex = 0
        Cmb_Tipo_Factura.ListIndex = 0
        Fra_Busqueda.Visible = False
        Fra_Busqueda.Enabled = False
        Fra_Busqueda_Con_Controles.Enabled = False
        Lbl_Dias_Credito.Caption = "Días de Crédito"
        Btn_Agregar.Enabled = True
        Btn_Eliminar.Enabled = True
        Chc_Adenda.Value = 0
        Chc_Relacionados.Value = 0
        lbl_estatus_cancel.Visible = False
        Chc_Adenda.Value = 0
    Else
        'Validacion de que esten todos los datos requeridos para dar de alta la factura
        If Grid_Detalle_Factura.Rows > 1 Then
            If Txt_No_Factura.text <> "" Then
                If Cmb_Nombre_Cliente.ListIndex > -1 Then
                    If Me.Caption = "FACTURAS CLIENTES" Then
                        'Se validan los datos completos del receptor
                        If Txt_Codigo_Postal.text <> "" Then
                            If Trim(Len(Txt_Codigo_Postal.text)) < 5 Then
                                MsgBox "El código postal no contiene la longitud requerida, favor de verificar", vbExclamation
                                Exit Sub
                            End If
                        Else
                            MsgBox "Falta el código postal del cliente, favor de verificar", vbExclamation
                            Exit Sub
                        End If
                        If Txt_Ciudad_Cliente.text = "" Then
                            MsgBox "Falta la ciudad del cliente, favor de verificar", vbExclamation
                            Exit Sub
                        End If
                        If Txt_Estado.text = "" Then
                            MsgBox "Falta el estado del cliente, favor de verificar", vbExclamation
                            Exit Sub
                        End If
                        If Txt_Pais.text = "" Then
                            MsgBox "Falta el país del cliente, favor de verificar", vbExclamation
                            Exit Sub
                        End If
                        If Cmb_Metodo_Pago.ListIndex = -1 Then
                            MsgBox "Favor de indicar el método de pago", vbExclamation
                            Cmb_Metodo_Pago.SetFocus
                            Exit Sub
                        End If
                        If Chc_Relacionados.Value = 1 Then
'                            Busca_UUID
                            If Cmb_Relacionados.ListIndex = -1 Then
                                MsgBox "Favor de indicar el tipo de relación", vbExclamation
                                Cmb_Relacionados.SetFocus
                                Exit Sub
                            End If
                            If Grid_Relacionados.Rows < 1 Then
                                MsgBox "Favor de agregar facturas relacionadas", vbExclamation
                                Exit Sub
                            End If
                        End If
                        If Cmb_Uso_CFDI.ListIndex = -1 Then
                                MsgBox "Favor de indicar el uso de CFDI", vbExclamation
                                Cmb_Uso_CFDI.SetFocus
                                Exit Sub
                            End If
                        If Txt_Cuenta_Pago.text <> "" Then
                           If Trim(Len(Txt_Cuenta_Pago.text)) < 4 Then
                                MsgBox "Debe registrar al menos los últimos 4 dígitos de la cuenta de pago, favor de verificar", vbExclamation
                                Txt_Cuenta_Pago.SetFocus
                                Exit Sub
                            End If
                        End If
                        If Chc_Adenda.Value = 1 Then
                            If Cmb_Tipo_Adenda.ListIndex = -1 Then
                                MsgBox "Debe seleccionar un tipo de adenda si es requerida.", vbExclamation
                                Cmb_Tipo_Adenda.SetFocus
                                Exit Sub
                            End If
                            If Cmb_Tipo_Adenda.ListIndex > 0 And Txt_Orden_Compra = "" Then
                                MsgBox "Favor de introducir la orden de compra de la factura", vbExclamation
                                Txt_Orden_Compra.SetFocus
                                Exit Sub
                            End If
                            If Cmb_Tipo_Adenda.text = "NADRO" And Txt_Plazo_Pago.text = "" Then
                                MsgBox "Favor de introducir el plazo de pago de la factura", vbExclamation
                                Txt_Plazo_Pago.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    If Cmb_Forma_Pago.ListIndex > -1 Then
                        'Alta en la base de datos de la factura
                        Alta_Factura
                    Else
                        MsgBox "Seleccione la forma de pago", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                        Exit Sub
                    End If
                Else
                    MsgBox "Seleccione el cliente para dar de alta el documento", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    Cmb_Nombre_Cliente.SetFocus
                End If
            Else
                MsgBox "Faltan el numero de documento", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        Else
            MsgBox "Faltan detalles del documento", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Alta_Factura
'DESCRIPCIÓN            : Hace el alta de la factura o una remisión en la base de datos
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21 de Enero del 2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Alta_Factura()

Dim Rs_Alta_Factura_Clientes As rdoResultset                            'Manejo del registro de Adm_Factura_Clientes
Dim Rs_Alta_Descripcion_Facturas As rdoResultset                        'Manejo del registro de Adm_Descripcion_Facturas
Dim Rs_Alta_Remision_Clientes As rdoResultset                           'Manejo del registro de Adm_Remision_Clientes
Dim Rs_Alta_Descripcion_Remision As rdoResultset                        'Manejo del registro de Adm_Descripcion_Remision
Dim Rs_Modifica_Alm_Salidas_Almacen_Detalles As rdoResultset            'Manejo del registro de Adm_Descripcion_Remision
Dim Rs_Modifica_Alm_Salidas_Almacen As rdoResultset
Dim Rs_Alta_Relacionados As rdoResultset
Dim Fecha_Xml As Date                           'Almacena la fecha del xml
Dim Str_Cadena_Original As String               'Almacena la cadena original de la factura
Dim Str_Cadena_UTF As String                    'Almacena la cadena en formato utf
Dim Str_Cadena_MD5 As String                    'Almacena la cadena en formato md5
Dim Str_Cadena_Sello As String                  'Almacena la cadena del sello digital
Dim Cont_Detalles_Factura As Integer            'Almacena el conteo de las partidas
Dim Contador As Integer                         'Almacena el número de partidas del grid
Dim Fecha_Timbrado As Date                      'Almacena la fecha del timbrado
Dim Grupo_Fecha_Timbrado() As String            'Almacena la fecha del timbrado separada por T
Dim Rs_Modifica_Factura As rdoResultset
Dim Impuesto As Double
Dim Grupo_Fecha() As String
Dim Unidades() As String
Dim Fecha_Generacion As Date
Dim Copias As Integer
Dim Impresiones As Integer
Dim Rs_Consulta_Regimen As rdoResultset
Dim Adenda As String
On Error GoTo handler
    Conexion_Base.BeginTrans
        CFD_Generales.No_Salida = Format(Cmb_Salidas.text, "0000000000")
        If Trim(Cmb_Tipo_Documento.text) = "FACTURA" Or Trim(Cmb_Tipo_Documento.text) = "NOTA CARGO" Then  'DA DE ALTA LA FACTURA
            If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
             If Trim(Cmb_Tipo_Documento.text) = "FACTURA" Then
                Txt_Factura_ID.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
                Txt_No_Factura.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas WHERE Forma_Factura = 'E'", "No_Factura_Electronica")
             Else
                Txt_Factura_ID.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
                Txt_No_Factura.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas WHERE Forma_Factura = 'E' AND Serie='NCA'", "No_Factura_Electronica")
             End If
                'Asigna los valores a las variable para la generación de la factura electrónica
                CFD_Generales.Version = "3.3"
                CFD_Generales.Serie = Trim(Txt_Serie.text)
                CFD_Generales.Folio = Val(Txt_No_Factura.text)
                CFD_Generales.Factura_ID = Format(Txt_Factura_ID.text, "0000000000")
                'Asigna la fecha del xml
                Fecha_Xml = Format(DTP_Fecha_Factura.Value, "dd/MM/yyyy") & " " & Format(Now, "HH:mm:ss")
                CFD_Generales.Fecha = Format(DTP_Fecha_Factura.Value, "yyyy-MM-dd") & "T" & Format(DTP_Fecha_Factura.Value, "HH:mm:ss")
                'CFD_Generales.Forma_Pago = CFD_Elimina_Espacios("pago en una sola exhibicion")
                
               CFD_Generales.Forma_Pago = Cmb_Forma_Pago.text
                'CFD_Generales.Forma_Pago = Trim(Cmb_Forma_Pago.ItemData(Cmb_Forma_Pago.ListIndex))
                'If Opt_Contado = True Then
                    'CFD_Generales.Forma_Pago_Credito_Contado = "CONTADO"
                    'CFD_Generales.Condiciones_Pago = "CONTADO"
                'Else
                  '  CFD_Generales.Forma_Pago_Credito_Contado = "CREDITO"
                  '  CFD_Generales.Condiciones_Pago = "CREDITO"
        '         '  CFD_Generales.Fecha_Vencimiento = DateAdd("d", Val(Txt_Factura_Remision_Dias_Credito.text), Format(DTP_Fecha_Venta.Value, "MM/dd/yyyy"))
                'End If
                CFD_Generales.SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
                CFD_Generales.Total = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                Impuesto = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
                CFD_Generales.Descuento = 0 'Val(Conectar_Ayudante.Quitar_Caracter(Lbl_Descuento.Caption, ","))
                CFD_Generales.Tipo_Moneda = "MXN"
                CFD_Generales.Tipo_Documento = "NORMAL"
                CFD_Generales.Tipo_Comprobante = "I"
                cp = Txt_Codigo_Postal.text
                
                CFD_Generales.Fecha_Vencimiento = Format(DTP_Fecha_Pago.Value, "dd-MMMM-yyyy")
                CFD_Generales.Metodo_Pago = Cmb_Metodo_Pago.text
                'CFD_Generales.Forma_Pago = CFD_Elimina_Espacios(Cmb_Forma_Pago.ItemData(Cmb_Forma_Pago.ListIndex))
                CFD_Generales.Cuenta_Pago = Trim(Txt_Cuenta_Pago.text)
                ReDim CFD_Relacionados_Conceptos(0)
                If Chc_Relacionados.Value = 1 Then
                    CFD_Relacionados.Existe = True
                    CFD_Relacionados.Relacionados = Cmb_Relacionados.text
                    ReDim CFD_Relacionados_Conceptos(Grid_Relacionados.Rows - 1)
                    
                    For i = 1 To Grid_Relacionados.Rows - 1
                        CFD_Relacionados_Conceptos(i).Serie = Grid_Relacionados.TextMatrix(i, 0)
                        CFD_Relacionados_Conceptos(i).Folio = Grid_Relacionados.TextMatrix(i, 1)
                        CFD_Relacionados_Conceptos(i).UUID = Grid_Relacionados.TextMatrix(i, 2)
                    Next i
                Else
                    CFD_Relacionados.Existe = False
                End If
                CFD_Generales.Uso_CFDI = Cmb_Uso_CFDI.text
                CFD_Generales.Condiciones_Pago = Txt_Condicion_Pago.text
                CFD_Generales.Tipo_Factura = Cmb_Tipo_Factura.text
                  
                'Asigna los datos del EMISOR al CFD para generar el xml
                CFD_Emisor.Nombre = CFD_Elimina_Espacios(Nombre_Emisor)
                CFD_Emisor.RFC = CFD_Elimina_Espacios(RFC_Emisor)
                CFD_Emisor.Expedido_En = CFD_Elimina_Espacios(Expedida_En)
                CFD_Emisor.Calle = CFD_Elimina_Espacios(Calle_Emisor)
                CFD_Emisor.No_Exterior = CFD_Elimina_Espacios(No_Exterior_Emisor)
                CFD_Emisor.No_Interior = CFD_Elimina_Espacios(No_Interior_Emisor)
                CFD_Emisor.Colonia = CFD_Elimina_Espacios(Colonia_Emisor)
                CFD_Emisor.cp = CFD_Elimina_Espacios(Codigo_Postal_Emisor)
        '        CFD_Emisor.Localidad = CFD_Elimina_Espacios(Municipio_Emisor)
                CFD_Emisor.Municipio = CFD_Elimina_Espacios(Municipio_Emisor)
                CFD_Emisor.Estado = CFD_Elimina_Espacios(Estado_Emisor)
                CFD_Emisor.Pais = CFD_Elimina_Espacios(Pais_Emisor)
                CFD_Emisor.Referencia = ""
                Mi_SQL = "SELECT Mensaje_Factura FROM Cat_Parametros_Factura_Electronica"
                Set Rs_Consulta_Regimen = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                CFD_Emisor.Regimen_Fiscal = Mid(Rs_Consulta_Regimen.rdoColumns("Mensaje_Factura"), 1, 3)
                Rs_Consulta_Regimen.Close
                                
                'Asigna los datos del RECEPTOR al CFD para generar el xml
                CFD_Receptor.Nombre = CFD_Elimina_Espacios(Cmb_Nombre_Cliente.text)
                CFD_Receptor.RFC = CFD_Elimina_Espacios(Txt_RFC_Cliente.text)
                CFD_Receptor.RFC = Conectar_Ayudante.Quitar_Caracter(CFD_Receptor.RFC, "-")
                CFD_Receptor.Calle = CFD_Elimina_Espacios(Txt_Direccion_Cliente.text)
                CFD_Receptor.No_Exterior = CFD_Elimina_Espacios(Txt_No_Exterior.text)
                CFD_Receptor.No_Interior = CFD_Elimina_Espacios(Txt_No_Interior.text)
                CFD_Receptor.Colonia = CFD_Elimina_Espacios(Txt_Colonia_Cliente.text)
                CFD_Receptor.cp = CFD_Elimina_Espacios(Txt_Codigo_Postal.text)
        '        CFD_Receptor.Localidad = CFD_Elimina_Espacios(Txt_Factura_Remision_Ciudad.text)
                CFD_Receptor.Municipio = CFD_Elimina_Espacios(Txt_Ciudad_Cliente.text)
                CFD_Receptor.Estado = CFD_Elimina_Espacios(Txt_Estado.text)
                CFD_Receptor.Pais = "MEXICO"
                CFD_Receptor.Referencia = ""
                CFD_Receptor.Uso_CFDI = CFD_Elimina_Espacios(Cmb_Uso_CFDI.text)
                
                Contador = 0
                'Valida que el numero de partidas
                For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
                    Contador = Contador + 1
                Next
                'Asigna el conteo de partidas al arreglo
                ReDim CFD_Conceptos(Contador)
                'Recorre las partidas del grid
                Contador = 0
                For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
                    
                    Contador = Contador + 1
                    Unidades = Split(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 1), "-")
                    CFD_Conceptos(Contador).Unidad = Unidades(1)
                    CFD_Conceptos(Contador).Unidad_Medida = Unidades(0)
                    CFD_Conceptos(Contador).Cod_prod = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3)
                    CFD_Conceptos(Contador).No_Identificacion = "" 'Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 12)
                    CFD_Conceptos(Contador).Cantidad = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 0), ","))
                    CFD_Conceptos(Contador).Descripcion = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2))
                    CFD_Conceptos(Contador).Valor_Unitario = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4), ",")), "#0.00")
                    CFD_Conceptos(Contador).Importe = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 5), ",")), "#0.00")
                    'CFD_Conceptos(Contador).Unidad = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 10))
                    CFD_Conceptos(Contador).Aplica_IVA = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 11))
                    CFD_Conceptos(Contador).Impuesto = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 8), ",")) * 0.16, "#0.00")
                    If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 11) = "SI" Then
                        CFD_Conceptos(Contador).IVA_Producto = True
                    Else
                        CFD_Conceptos(Contador).IVA_Producto = False
                    End If
                    CFD_Conceptos(Contador).Posicion_OC = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 13)
                Next
                
                'Asigna los datos de los IMPUESTO al CFD para generar el xml segun el tipo de factura
                ReDim CFD_Impuestos(0)
                ReDim CFD_Impuestos_Retenidos(0)
                ReDim CFD_Impuestos_Locales(0)
                If Val(Txt_IVA) > 0 Then
                    ReDim CFD_Impuestos(1)
                    CFD_Impuestos(1).Impuesto = "002"
                    CFD_Impuestos(1).Tasa = PG_Retencion_IVA 'Val(PG_Retencion_IVA * 100)
                    CFD_Impuestos(1).Importe = Impuesto
                    IVA_EXENTO = False
                Else
                    'ReDim CFD_Impuestos(1)
                    'CFD_Impuestos(1).Impuesto = "002"
                    'CFD_Impuestos(1).Tasa = "0"
                    'CFD_Impuestos(1).Importe = 0
                    IVA_EXENTO = True
                End If
                If Chc_Adenda.Value = 1 Then
                    If Cmb_Tipo_Adenda.text = "NADRO" Then
                        Adenda = "Adenda_Nadro"
                        CFD_Generales.Plazo = Trim(Txt_Plazo_Pago.text)
                    Else
                        Adenda = "Adenda"
                    End If
                    CFD_Generales.Orden_Compra = Trim(Txt_Orden_Compra.text)
                Else
                    Adenda = ""
                End If
                'Crea el sello digital con toda la informacion
                CFD_Generales.No_Certificado = CFD_Consulta_Serie_Certificado(Ruta_Certificado)
                CFD_Generales.Certificado = CFD_Consulta_Certificado(Ruta_Certificado)
                Str_Cadena_Original = CFD_Cadena_Original(Cmb_Tipo_Factura.text)
                Str_Cadena_UTF = CFD_Valida_Caracteres_UTF(Str_Cadena_Original)
                Str_Cadena_MD5 = CFD_Genera_MD5(Str_Cadena_UTF)
                Str_Cadena_Sello = CFD_Genera_Sello(Str_Cadena_UTF, Ruta_Llave_Privada)
                CFD_Generales.Cadena_Original = Str_Cadena_UTF

                CFD_Generales.Sello = Str_Cadena_Sello
                CFD_Generales.Importe_Letra = Conectar_Ayudante.Convierte_Cantidad_Letras(Format(CStr(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))), "#0.00"))
            End If
            If Chc_Relacionados.Value = 1 Then
                Set Rs_Alta_Relacionados = Conectar_Ayudante.Recordset_Agregar("Ope_Relacionados")
                For i = 1 To Grid_Relacionados.Rows - 1
                    With Rs_Alta_Relacionados
                        .AddNew
                            .rdoColumns("Relacionado_ID") = Format(Conectar_Ayudante.Maximo_Catalogo("Ope_Relacionados", "Relacionado_ID"), "0000000000")
                            .rdoColumns("Serie") = CFD_Generales.Serie
                            .rdoColumns("No_Factura_Electronica") = Format(CFD_Generales.Folio, "0000000000")
                            .rdoColumns("Tipo_Relacion") = Cmb_Relacionados.text
                            .rdoColumns("UUID_Relacion") = Grid_Relacionados.TextMatrix(i, 2)
                            .rdoColumns("Factura_Rel") = Format(Grid_Relacionados.TextMatrix(i, 1), "0000000000")
                            .rdoColumns("Serie_Rel") = Grid_Relacionados.TextMatrix(i, 0)
                        .Update
                    End With
                Next i
            End If
            Set Rs_Alta_Factura_Clientes = Conectar_Ayudante.Recordset_Agregar("Adm_Clientes_Facturas")
            With Rs_Alta_Factura_Clientes
                .AddNew
                    If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                        .rdoColumns("No_Factura") = Format(Txt_Factura_ID.text, "0000000000")
                        .rdoColumns("No_Factura_Electronica") = Format(Txt_No_Factura.text, "0000000000")
                        .rdoColumns("Serie") = Trim(Txt_Serie.text)
                        'Convierte la fecha de timbrado
                        Grupo_Fecha = Split(CFD_Generales.Fecha, "T")
                        Fecha_Generacion = Grupo_Fecha(0) & " " & Grupo_Fecha(1)
                        .rdoColumns("Fecha_Creo_XML") = Format(Fecha_Generacion, "MM/dd/yyyy HH:mm:ss")
                    Else
                        .rdoColumns("No_Factura") = Format(Txt_No_Factura.text, "0000000000")
                    End If
                    .rdoColumns("Cliente_ID") = Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000")
                    .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Pago") = Format(DTP_Fecha_Pago.Value, "MM/dd/yyyy")
                    .rdoColumns("Forma_Factura") = Mid(Cmb_Tipo_Factura.text, 1, 1)
                    'If Opt_Contado = True Then .rdoColumns("Tipo_Pago") = "CONTADO"
                    'If Opt_Credito = True Then .rdoColumns("Tipo_Pago") = "CREDITO"
                    .rdoColumns("Tipo_Pago") = Cmb_Metodo_Pago.text
                    .rdoColumns("Subtotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
                    .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
                    .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                    .rdoColumns("Cancelada") = "N"
                    .rdoColumns("Comentarios") = Txt_Comentarios.text
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now()
                    If Cmb_Salidas.ListIndex > -1 Then .rdoColumns("No_Salida") = Format(Cmb_Salidas.text, "0000000000")
                    .rdoColumns("Abono") = 0
                    .rdoColumns("Saldo") = CDbl(Txt_Total.text)
                    .rdoColumns("Pagada") = "N"
                    .rdoColumns("Tipo_Documento") = Cmb_Tipo_Documento.text
'                    If Cmb_Metodo_Pago.ListIndex = 0 Then
'                        .rdoColumns("Metodo_Pago") = CFD_Elimina_Espacios("PPD")
'                    Else
                    .rdoColumns("Metodo_Pago") = Cmb_Metodo_Pago.text
'                    End If
                    .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                    .rdoColumns("Uso_CFDI") = Mid(Cmb_Uso_CFDI.text, 1, 3)
                    .rdoColumns("Condiciones_Pago") = Trim(Txt_Condicion_Pago.text)
                    .rdoColumns("No_Cuenta_Pago") = Trim(Txt_Cuenta_Pago.text)
                    If Chc_Relacionados.Value = 1 Then
                        .rdoColumns("Relacionado") = "S"
                        .rdoColumns("Tipo_Relacion") = Cmb_Relacionados.text
                        .rdoColumns("UUID_Relacion") = Trim(Txt_UUID_Relacion.text)
                    Else
                        .rdoColumns("Relacionado") = "N"
                        .rdoColumns("Tipo_Relacion") = ""
                        .rdoColumns("UUID_Relacion") = ""
                    End If
                    If Chc_Adenda.Value = 1 Then
                        .rdoColumns("Orden_Compra") = Trim(Txt_Orden_Compra.text)
                    End If
                .Update
            End With
            
            Set Rs_Alta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Agregar("Adm_Clientes_Facturas_Detalles")
            For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
                'Llena la tabla de Adm_Clientes_Facturas_Detalles con los datos contenidos en el grid
                With Rs_Alta_Descripcion_Facturas
                    .AddNew
                        If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                            .rdoColumns("No_Factura") = Format(Txt_Factura_ID.text, "0000000000")
                            .rdoColumns("No_Factura_Electronica") = Format(Txt_No_Factura.text, "0000000000")
                        Else
                            .rdoColumns("No_Factura") = Format(Txt_No_Factura.text, "0000000000")
                        End If
                        If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 6) <> "" Then
                            .rdoColumns("Producto_ID") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 6)
                        Else
                            .rdoColumns("Producto_ID") = Null
                        End If
                        .rdoColumns("Descripcion") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2)
                        .rdoColumns("Cantidad") = Val(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 0))
                        .rdoColumns("Precio") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4), ","))
                        .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 5), ","))
                        .rdoColumns("Unidad") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 1)
                        .rdoColumns("Clave_SAT") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3)
                    .Update
                End With
                If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7) <> "" Then
                    'MODIFICA LA SALIDA LA PONE COMO FACTURADA
                    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7), "0000000000") & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 6), "00000") & "' AND Facturado ='NO'"
                    Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
                        With Rs_Modifica_Alm_Salidas_Almacen_Detalles
                            .Edit
                                .rdoColumns("Facturado") = "SI"
                                .rdoColumns("Precio_Venta") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2), ","))
                                .rdoColumns("Importe") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3), ","))
                                .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                            .Update
                        End With
                    End If
                    Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
                    'EL ESTATUS DE LA SALIDA QUEDA COMO FACTURADA
                    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen WHERE No_Salida ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7), "0000000000") & "' "
                    Set Rs_Modifica_Alm_Salidas_Almacen = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Alm_Salidas_Almacen.EOF Then
                        With Rs_Modifica_Alm_Salidas_Almacen
                            .Edit
                                .rdoColumns("Estatus") = "FACTURADA"
                                .rdoColumns("Comentarios_Remision") = Trim(Txt_Comentarios.text)
                            .Update
                        End With
                    End If
                    Rs_Modifica_Alm_Salidas_Almacen.Close
                End If
            Next Cont_Detalles_Factura
            Rs_Alta_Factura_Clientes.Close
            Rs_Alta_Descripcion_Facturas.Close
            
            'Si se generó una factura y electronica valida si se enviara Emial y crea el txt
            If Cmb_Tipo_Factura.ListIndex = 0 Then
                'Crea el xml con los datos de la factura
                Call CFD_Crea_Xml("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text), "PAGOS", Adenda)
                
                'Convierte la fecha de timbrado
                Grupo_Fecha_Timbrado = Split(Timbrado_FechaTimbrado, "T")
                Fecha_Timbrado = Grupo_Fecha_Timbrado(0) & " " & Grupo_Fecha_Timbrado(1)
                
                'Actualiza la factura con el timbrado
                Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas"
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(Val(Txt_Factura_ID.text), "0000000000") & "'"
                Set Rs_Modifica_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Modifica_Factura
                        If Not .EOF Then
                            .Edit
                                .rdoColumns("No_Certificado") = CFD_Generales.No_Certificado
                                .rdoColumns("Timbre_Version") = Timbrado_VersionSat
                                .rdoColumns("Timbre_UUID") = Timbrado_UUID
                                .rdoColumns("Timbre_Fecha_Timbrado") = Format(Fecha_Timbrado, "MM/dd/yyyy HH:mm:ss")
                                .rdoColumns("Timbre_selloCFD") = Timbrado_selloCFD
                                .rdoColumns("Timbre_noCertificadoSAT") = Timbrado_noCertificadoSAT
                                .rdoColumns("Timbre_selloSAT") = Timbrado_selloSAT
                                .rdoColumns("Ruta_Codigo_BD") = Ruta_Pdfs & "\CFDI_" & Trim(Txt_Serie.text) & "_" & Trim(CFD_Generales.Folio) & ".bmp"
                                CFD_Generales.Imagen_BMP = Ruta_Pdfs & "\CFDI_" & Trim(Txt_Serie.text) & "_" & Trim(CFD_Generales.Folio) & ".bmp"
                            .Update
                        End If
                    End With
                Rs_Modifica_Factura.Close
            End If
        Else 'DA DE ALTA LA REMISION
            Set Rs_Alta_Factura_Clientes = Conectar_Ayudante.Recordset_Agregar("Adm_Clientes_Remisiones")
            With Rs_Alta_Factura_Clientes
                .AddNew
                    .rdoColumns("No_Remision") = Format(Txt_No_Factura.text, "0000000000")
                    .rdoColumns("Cliente_ID") = Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000")
                    .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Pago") = Format(DTP_Fecha_Pago.Value, "MM/dd/yyyy")
                    'If Opt_Contado = True Then .rdoColumns("Tipo_Pago") = "CONTADO"
                    'If Opt_Credito = True Then .rdoColumns("Tipo_Pago") = "CREDITO"
                    .rdoColumns("Subtotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
                    .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
                    .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                    .rdoColumns("Cancelada") = "N"
                    .rdoColumns("Comentarios") = Txt_Comentarios.text
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now()
                    If Cmb_Salidas.ListIndex > -1 Then .rdoColumns("No_Salida") = Format(Cmb_Salidas.text, "0000000000")
                    .rdoColumns("Abono") = 0
                    .rdoColumns("Saldo") = CDbl(Txt_Total.text)
                    .rdoColumns("Pagada") = "N"
                    .rdoColumns("Tipo_Documento") = "REMISION"
                    .rdoColumns("Facturada") = "NO"
                .Update
            End With
            Set Rs_Alta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Agregar("Adm_Clientes_Remisiones_Detalles")
            For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
                'Llena la tabla de Adm_Clientes_Facturas_Detalles con los datos contenidos en el grid
                With Rs_Alta_Descripcion_Facturas
                    .AddNew
                        .rdoColumns("No_Remision") = Format(Txt_No_Factura.text, "0000000000")
                        If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4) <> "" Then
                            .rdoColumns("Producto_ID") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4)
                        Else
                            .rdoColumns("Producto_ID") = Null
                        End If
                        .rdoColumns("Descripcion") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2)
                        .rdoColumns("Cantidad") = CDbl(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 0))
                        .rdoColumns("Precio") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2), ","))
                        .rdoColumns("Importe") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3), ","))
                    .Update
                End With
                If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7) <> "" Then
                    'MODIFICA LA SALIDA LA PONE COMO FACTURADA
                    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7), "0000000000") & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4), "00000") & "' AND Facturado ='NO'"
                    Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
                        With Rs_Modifica_Alm_Salidas_Almacen_Detalles
                            .Edit
                                .rdoColumns("Facturado") = "SI"
                                .rdoColumns("Precio_Venta") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2), ","))
                                .rdoColumns("Importe") = CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3), ","))
                                .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                            .Update
                        End With
                    End If
                    Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
                    'EL ESTATUS DE LA SALIDA QUEDA COMO REMISIONADA
                    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen WHERE No_Salida ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 7), "0000000000") & "'"
                    Set Rs_Modifica_Alm_Salidas_Almacen = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Alm_Salidas_Almacen.EOF Then
                        With Rs_Modifica_Alm_Salidas_Almacen
                            .Edit
                                .rdoColumns("Estatus") = "REMISIONADA"
                                .rdoColumns("Comentarios_Remision") = Trim(Txt_Comentarios.text)
                            .Update
                        End With
                    End If
                    Rs_Modifica_Alm_Salidas_Almacen.Close
                End If
            Next Cont_Detalles_Factura
            Rs_Alta_Factura_Clientes.Close
            Rs_Alta_Descripcion_Facturas.Close
        End If
    Conexion_Base.CommitTrans
    
    Select Case Trim(Cmb_Tipo_Documento.text)
        Case "FACTURA"
            If Cmb_Tipo_Factura.ListIndex = 0 Then
                Me.MousePointer = 0
                MsgBox "La factura ha sido dada de alta", vbInformation
                Me.MousePointer = 11
                Call Valida_Termino_Folios_Activos("FACTURA", Trim(Txt_Serie.text), Trim(Txt_No_Factura.text))
                Call CFD_Crea_PDF("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text), "FACTURA", "NORMAL", Year(Fecha_Xml))
                Call Muestra_PDF("FACTURA")
                If MsgBox("¿Desea realizar el envío de la factura por correo", vbYesNo + vbQuestion, "ENVÍO DE CORREO") = vbYes Then
                    Btn_Enviar_Email_Click
                End If
            Else
                If MsgBox("La Factura ha sido dado de alta" & Chr(13) & "¿Desea Imprimir la factura?", vbYesNo + vbInformation) = vbYes Then
                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
                    If Val(Impresiones) > 0 Then
                        For Copias = 1 To Impresiones
                            Call Imprimir_Facturas
                        Next
                    Else
                        MsgBox "Parámetro inválido", vbExclamation
                        Exit Sub
                    End If
                    MsgBox "Factura(s) enviada(s) a impresión", vbInformation
                End If
            End If
            If Cmb_Salidas.text <> "" Then
                If MsgBox("¿Desea Imprimir la Remision?", vbYesNo + vbInformation) = vbYes Then
                    Call CFD_Crea_PDF("REMISION_" & CFD_Generales.No_Salida, "REMISION", "", Year(Fecha_Xml))
                    Call Muestra_PDF("REMISION")
''                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
''                    If Val(Impresiones) > 0 Then
''                        For Copias = 1 To Impresiones
''                            Call Imprimir_Remision
''                        Next
''                    Else
''                        MsgBox "Parámetro inválido", vbExclamation
''                        Exit Sub
''                    End If
''                    MsgBox "Remision(es) enviada(s) a impresión", vbInformation
                End If
            End If
        Case "NOTA CARGO"
            If Cmb_Tipo_Factura.ListIndex = 0 Then
                Me.MousePointer = 0
                MsgBox "La nota de cargo ha sido dada de alta", vbInformation
                Me.MousePointer = 11
                Call Valida_Termino_Folios_Activos("FACTURA", Trim(Txt_Serie.text), Trim(Txt_No_Factura.text))
                Call CFD_Crea_PDF("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text), "NOTA CARGO", "NORMAL", Year(Fecha_Xml))
                Call Muestra_PDF("FACTURA")
                If MsgBox("¿Desea realizar el envío de la nota de cargo por correo", vbYesNo + vbQuestion, "ENVÍO DE CORREO") = vbYes Then
                    Btn_Enviar_Email_Click
                End If
            Else
                If MsgBox("La nota de cargo ha sido dado de alta" & Chr(13) & "¿Desea Imprimir la nota de cargo?", vbYesNo + vbInformation) = vbYes Then
                    Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
                    If Val(Impresiones) > 0 Then
                        For Copias = 1 To Impresiones
                            Call Imprimir_Facturas
                        Next
                    Else
                        MsgBox "Parámetro inválido", vbExclamation
                        Exit Sub
                    End If
                    MsgBox "Nota(s) enviada(s) a impresión", vbInformation
                End If
            End If
            
        Case "REMISION"
            If MsgBox("La remision ha sido dado de alta" & Chr(13) & " ¿Desea enviarla a Imprimir?", vbYesNo + vbInformation) = vbYes Then
                Call CFD_Crea_PDF("REMISION_" & CFD_Generales.No_Salida, "REMISION", "", Year(Fecha_Xml))
                Call Muestra_PDF("REMISION")
''                Impresiones = InputBox("Ingrese el número de copias que requiere", "COPIAS DE IMPRESION")
''                If Val(Impresiones) > 0 Then
''                    For Copias = 1 To Impresiones
''                        Call Imprimir_Remision
''                    Next
''                Else
''                    MsgBox "Parámetro inválido", vbExclamation
''                    Exit Sub
''                End If
''                MsgBox "Remision(es) enviada(s) a impresión", vbInformation
            End If
    End Select
    Fra_Datos_Cliente.Enabled = False
    Fra_Datos_Factura.Enabled = False
    Fra_Detalle_Factura.Enabled = False
    Fra_Comentarios.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Imprimir.Enabled = True
    Btn_Cancelar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    If Cmb_Tipo_Factura.ListIndex = 0 Then
        Btn_Enviar_Email.Enabled = True
    Else
        Btn_Enviar_Email.Enabled = False
    End If
    Me.MousePointer = 0
    Exit Sub
handler:
    Me.MousePointer = 0
    MsgBox Err.Description
    Conexion_Base.RollbackTrans
    If Err.Number <> 0 Then
        MsgBox Err.Description
    Else
        For Each Er In rdoErrors
            MsgBox Er.Description
        Next Er
    End If
End Sub

Public Sub Codigo_Unidades(Unidad As String)
    Dim Unidades As String
    Mi_SQL = "SELECT Clave_Unidad FROM Cat_Unidades_Medida WHERE Clave = '" & Format(Unidad, "#00000") & "'"
            Set Rs_Consulta_Unidad = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Unidades = Rs_Consulta_Unidad.rdoColumns("Clave_Unidad")
            Rs_Consulta_Unidad.Close
            UM = Unidades
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Muestra_PDF
'DESCRIPCIÓN: Valida si existe el PDF de la factura para mostrarlo en pantalla,
'             de lo contrario llama la funcion para volver a crear el PDF
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 14-Agosto-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'******************************************************************************'
Private Sub Muestra_PDF(Tipo As String)
Dim Nombre_Archivo As String   'Almacena el nombre del archivo

On Error GoTo errorHandler

    If Txt_No_Factura.text <> "" Then
        MDIFrm_Apl_Principal.MousePointer = 11
        If Tipo = "FACTURA" Then
            'Asigna el nombre del archivo
            Nombre_Archivo = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text) & ".pdf"
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Pdfs & "\" & Nombre_Archivo, "ARCHIVO") = True Then
                'Envia para abrir el archivo
                ShellExecute ByVal 0&, "open", Ruta_Pdfs & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
            Else 'Regenera el pdf
                Regenerar_PDF_Factura
                'Envia para abrir el archivo
                ShellExecute ByVal 0&, "open", Ruta_Pdfs & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
            End If
        Else
            'Asigna el nombre del archivo
            Nombre_Archivo = "REMISION_" & Cmb_Salidas.text & ".pdf"
            'Valida que exista el PDF de la remisión para eliminarlo
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Remisiones & "\" & Nombre_Archivo, "ARCHIVO") = True Then
                Kill Ruta_Remisiones & "\" & Nombre_Archivo
            End If
            CFD_Generales.No_Salida = Format(Cmb_Salidas.text, "0000000000")
            'Regenera el pdf
            Call CFD_Crea_PDF("REMISION_" & Cmb_Salidas.text, "REMISION", "", Year(Fecha_Xml))
            'Envia para abrir el archivo
            ShellExecute ByVal 0&, "open", Ruta_Remisiones & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
        End If
    Else
        MsgBox "Seleccione la factura", vbExclamation
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    If Err.Number = 70 Then
        MsgBox "El archivo PDF se encuentra abierto actualmente, favor de verificar", vbExclamation
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Regenerar_PDF_Factura
'DESCRIPCIÓN: Regenera el PDF la factura electronica seleccionada
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 14-Agosto-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'******************************************************************************'
Private Sub Regenerar_PDF_Factura()
Dim Rs_Consulta_Timbrado As rdoResultset    'Variable para el manejo de la tabla
Dim Rs_Consulta_Fecha_XML As rdoResultset   'Variable para el manejo de la tabla
Dim Fecha_Xml As Date                       'Almacena la fecha del xml
Dim Fecha_Timbrado As Date                  'Almacena la fecha del timbrado
Dim Str_Cadena_Original As String           'Almacena la cadena original de la factura
Dim Str_Cadena_UTF As String                'Almacena la cadena en formato utf
Dim Str_Cadena_MD5 As String                'Almacena la cadena en formato md5
Dim Str_Cadena_Sello As String              'Almacena la cadena del sello digital
Dim Cont_Detalles_Factura As Integer        'Almacena el conteo de las partidas
Dim Impuesto As Double                      'Almacena el importe por concepto de impuestos
Dim SubTotal As Double                      'Almacena el importe por concepto de subtotal
Dim Contador As Integer

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Consulta la fecha de generación del xml y datos del timbrado
    Mi_SQL = "SELECT No_Factura, Subtotal, IVA, Total, Fecha_Creo_XML, No_Certificado, Timbre_Version, Timbre_UUID, Timbre_Fecha_Timbrado,"
    Mi_SQL = Mi_SQL & " Timbre_SelloCFD, Timbre_noCertificadoSAT, Timbre_selloSAT, Ruta_Codigo_BD"
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
    Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(Val(Txt_Factura_ID.text), "0000000000") & "'"
    Set Rs_Consulta_Fecha_XML = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Fecha_XML.EOF Then
            'Asigna la fecha del xml
            Fecha_Xml = Rs_Consulta_Fecha_XML.rdoColumns("Fecha_Creo_XML")
            CFD_Generales.SubTotal = Format(Rs_Consulta_Fecha_XML.rdoColumns("Subtotal"), "#0.00")
            CFD_Generales.Impuestos = Format(Rs_Consulta_Fecha_XML.rdoColumns("IVA"), "#0.00")
            CFD_Generales.Total = Format(Rs_Consulta_Fecha_XML.rdoColumns("Total"), "#0.00")
            CFD_Generales.No_Certificado = Trim(Rs_Consulta_Fecha_XML.rdoColumns("No_Certificado"))
            Timbrado_VersionSat = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Timbre_Version"))
            Timbrado_UUID = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Timbre_UUID"))
            Fecha_Timbrado = Rs_Consulta_Fecha_XML.rdoColumns("Timbre_Fecha_Timbrado")
            Timbrado_selloCFD = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Timbre_selloCFD"))
            Timbrado_noCertificadoSAT = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Timbre_noCertificadoSAT"))
            Timbrado_selloSAT = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Timbre_selloSAT"))
            CFD_Generales.Imagen_BMP = Trim(Rs_Consulta_Fecha_XML.rdoColumns("Ruta_Codigo_BD"))
        Else
            Fecha_Xml = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy") & " " & Format(Now, "HH:mm:ss")
            Fecha_Timbrado = Fecha_Xml
        End If
    Rs_Consulta_Fecha_XML.Close
    
    'Asigna los valores a las variable para la generación de la factura electrónica
    CFD_Generales.Version = "3.2"
    CFD_Generales.Serie = Trim(Txt_Serie.text)
    CFD_Generales.Folio = Val(Txt_No_Factura.text)
    CFD_Generales.Factura_ID = Format(Txt_Factura_ID.text, "0000000000")
    'Asigna la fecha del xml
    Fecha_Xml = Format(DTP_Fecha_Factura.Value, "dd/MM/yyyy") & " " & Format(Now, "HH:mm:ss")
    CFD_Generales.Fecha = Format(DTP_Fecha_Factura.Value, "yyyy-MM-dd") & "T" & Format(DTP_Fecha_Factura.Value, "HH:mm:ss")
    CFD_Generales.Forma_Pago = CFD_Elimina_Espacios(Cmb_Forma_Pago.ItemData(Cmb_Forma_Pago.ListIndex))
    'If Opt_Contado = True Then
    '    CFD_Generales.Forma_Pago_Credito_Contado = "CONTADO"
    '    CFD_Generales.Condiciones_Pago = "CONTADO"
    'Else
    '    CFD_Generales.Forma_Pago_Credito_Contado = "CREDITO"
    '    CFD_Generales.Condiciones_Pago = "CREDITO"
'            CFD_Generales.Fecha_Vencimiento = DateAdd("d", Val(Txt_Factura_Remision_Dias_Credito.text), Format(DTP_Fecha_Venta.Value, "MM/dd/yyyy"))
    'End If
    CFD_Generales.SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
    CFD_Generales.Total = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
    Impuesto = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
    CFD_Generales.Descuento = 0 'Val(Conectar_Ayudante.Quitar_Caracter(Lbl_Descuento.Caption, ","))
    CFD_Generales.Tipo_Moneda = "MXN"
    CFD_Generales.Tipo_Documento = "NORMAL"
    CFD_Generales.Tipo_Comprobante = "I"
    
    CFD_Generales.Fecha_Vencimiento = Format(DTP_Fecha_Pago.Value, "dd-MMMM-yyyy")
    If Cmb_Metodo_Pago.ListIndex = 0 Then
        CFD_Generales.Metodo_Pago = CFD_Elimina_Espacios("PPD")
    Else
        CFD_Generales.Metodo_Pago = CFD_Elimina_Espacios("PUE")
    End If
    CFD_Generales.Forma_Pago = CFD_Elimina_Espacios(Cmb_Forma_Pago.ItemData(Cmb_Forma_Pago.ListIndex))
    CFD_Generales.Cuenta_Pago = Trim(Txt_Cuenta_Pago.text)
    CFD_Generales.Uso_CFDI = CFD_Elimina_Espacios(Cmb_Uso_CFDI.text)
    CFD_Generales.Condiciones_Pago = CFD_Elimina_Espacios(Txt_Condicion_Pago.text)
    If Chc_Relacionados.Value = 1 Then
        CFD_Generales.Relacionado = "S"
        CFD_Generales.Tipo_Relacion = CFD_Elimina_Espacios(Cmb_Relacionados.text)
        CFD_Generales.UUID_Relacion = CFD_Elimina_Espacios(Txt_UUID_Relacion.text)
    Else
        CFD_Generales.Relacionado = "N"
    End If
            
    'Asigna los datos del EMISOR al CFD para generar el xml
    CFD_Emisor.Nombre = CFD_Elimina_Espacios(Nombre_Emisor)
    CFD_Emisor.RFC = CFD_Elimina_Espacios(RFC_Emisor)
    CFD_Emisor.Expedido_En = CFD_Elimina_Espacios(Expedida_En)
    CFD_Emisor.Calle = CFD_Elimina_Espacios(Calle_Emisor)
    CFD_Emisor.No_Exterior = CFD_Elimina_Espacios(No_Exterior_Emisor)
    CFD_Emisor.No_Interior = CFD_Elimina_Espacios(No_Interior_Emisor)
    CFD_Emisor.Colonia = CFD_Elimina_Espacios(Colonia_Emisor)
    CFD_Emisor.cp = CFD_Elimina_Espacios(Codigo_Postal_Emisor)
'        CFD_Emisor.Localidad = CFD_Elimina_Espacios(Municipio_Emisor)
    CFD_Emisor.Municipio = CFD_Elimina_Espacios(Municipio_Emisor)
    CFD_Emisor.Estado = CFD_Elimina_Espacios(Estado_Emisor)
    CFD_Emisor.Pais = CFD_Elimina_Espacios(Pais_Emisor)
    CFD_Emisor.Referencia = ""
    
                    
    'Asigna los datos del RECEPTOR al CFD para generar el xml
    CFD_Receptor.Nombre = CFD_Elimina_Espacios(Cmb_Nombre_Cliente.text)
    CFD_Receptor.RFC = CFD_Elimina_Espacios(Txt_RFC_Cliente.text)
    CFD_Receptor.RFC = Conectar_Ayudante.Quitar_Caracter(CFD_Receptor.RFC, "-")
    CFD_Receptor.Calle = CFD_Elimina_Espacios(Txt_Direccion_Cliente.text)
    CFD_Receptor.No_Exterior = CFD_Elimina_Espacios(Txt_No_Exterior.text)
    CFD_Receptor.No_Interior = CFD_Elimina_Espacios(Txt_No_Interior.text)
    CFD_Receptor.Colonia = CFD_Elimina_Espacios(Txt_Colonia_Cliente.text)
    CFD_Receptor.cp = CFD_Elimina_Espacios(Txt_Codigo_Postal.text)
'        CFD_Receptor.Localidad = CFD_Elimina_Espacios(Txt_Factura_Remision_Ciudad.text)
    CFD_Receptor.Municipio = CFD_Elimina_Espacios(Txt_Ciudad_Cliente.text)
    CFD_Receptor.Estado = CFD_Elimina_Espacios(Txt_Estado.text)
    CFD_Receptor.Pais = "MEXICO"
    CFD_Receptor.Referencia = ""
    
    Contador = 0
    'Valida que el numero de partidas
    For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
        Contador = Contador + 1
    Next
    'Asigna el conteo de partidas al arreglo
    ReDim CFD_Conceptos(Contador)
    'Recorre las partidas del grid
    Contador = 0
    For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
        Contador = Contador + 1
'            If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2) <> "" Then
'                CFD_Conceptos(Contador).No_Identificacion = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2))
'            End If
        CFD_Conceptos(Contador).Cantidad = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 0), ","))
        CFD_Conceptos(Contador).Descripcion = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 1))
        CFD_Conceptos(Contador).Valor_Unitario = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2), ",")), "#0.00")
        CFD_Conceptos(Contador).Importe = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3), ",")), "#0.00")
        CFD_Conceptos(Contador).Unidad = CFD_Elimina_Espacios(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 10))
    Next
    
    'Asigna los datos de los IMPUESTO al CFD para generar el xml segun el tipo de factura
    ReDim CFD_Impuestos(0)
    ReDim CFD_Impuestos_Retenidos(0)
    ReDim CFD_Impuestos_Locales(0)
    If Impuesto > 0 Then
        ReDim CFD_Impuestos(1)
        CFD_Impuestos(1).Impuesto = "IVA"
        CFD_Impuestos(1).Tasa = Val(PG_Retencion_IVA * 100)
        CFD_Impuestos(1).Importe = Impuesto
    Else
        ReDim CFD_Impuestos(1)
        CFD_Impuestos(1).Impuesto = "IVA"
        CFD_Impuestos(1).Tasa = "0"
        CFD_Impuestos(1).Importe = 0
    End If
           
    'Crea el sello digital con toda la informacion
    Str_Cadena_Original = CFD_Cadena_Original("")
    Str_Cadena_UTF = CFD_Valida_Caracteres_UTF(Str_Cadena_Original)
    Str_Cadena_MD5 = CFD_Genera_MD5(Str_Cadena_UTF)
    Str_Cadena_Sello = CFD_Genera_Sello(Str_Cadena_UTF, Ruta_Llave_Privada)
    CFD_Generales.Cadena_Original = Str_Cadena_UTF
    CFD_Generales.No_Certificado = CFD_Consulta_Serie_Certificado(Ruta_Certificado)
    CFD_Generales.Certificado = CFD_Consulta_Certificado(Ruta_Certificado)
    CFD_Generales.Sello = Str_Cadena_Sello
    CFD_Generales.Importe_Letra = Conectar_Ayudante.Convierte_Cantidad_Letras(Format(CStr(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))), "#0.00"))
            
    'Crea el PDF con los datos de la factura
    Call CFD_Crea_PDF("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text), "FACTURA", "NORMAL", Year(Fecha_Xml))
    
    MDIFrm_Apl_Principal.MousePointer = 0
    
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    'Obtiene el error
    If Err.Number = 7777 Or Err.Number = -255 Then
        MsgBox Err.Description
    Else
        If Err.Number = 76 Then
            MsgBox "No se encontró la ruta destino para almacenar el archivo PDF, favor de verificar", vbExclamation
        Else
            For Each Er In rdoErrors
                MsgBox Er.Description
            Next
        End If
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Cancela_Factura
'DESCRIPCIÓN                : pone como cancelada la factura
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 24 Enero 2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Public Sub Cancela_Factura()
Dim Rs_Cancela_Factura_Clientes As rdoResultset            'Manejo del registro de Adm_Factura_Clientes
Dim Rs_Cancela_Factura_Clientes_Detalles As rdoResultset
Dim Resultado As Integer
Dim Mi_SQL As String
Dim Cont_Fila As Integer
Dim Codigo_UUID As String                   'Almacena el codigo fiscal sat
Dim Motivo As String
Dim Mensaje As String
Dim Cancel As Boolean
On Error GoTo errorHandler
    Set Conectar_Ayudante = New Ayudante
    Motivo = InputBox("Ingrese el motivo de la cancelación", "CANCELACIÓN DE DOCUMENTOS")
    If Trim(Motivo) = "" Then
        MsgBox "Debe ingresar el motivo de la cancelación", vbExclamation
        Exit Sub
    End If
    Conexion_Base.BeginTrans
        MDIFrm_Apl_Principal.MousePointer = 11
        If Cmb_Tipo_Documento.text = "FACTURA" Or Cmb_Tipo_Documento.text = "NOTA CARGO" Then
            Mi_SQL = "SELECT No_Factura,Mensaje_Cancelado, Cancelada, Usuario_Cancelo, Fecha_Cancelo, Motivo_Cancelo, Timbre_UUID"
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
            If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(Txt_Factura_ID.text, "0000000000") & "'"
            Else
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(Txt_No_Factura.text, "0000000000") & "'"
            End If
            If Cmb_Tipo_Documento.text = "NOTA CARGO" Then
                Mi_SQL = Mi_SQL & " AND Serie = 'NCA'"
            End If
            Set Rs_Cancela_Factura_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                'Llena la tabla de Adm_Clientes_Facturas con los datos contenidos en las cajas de textos
                With Rs_Cancela_Factura_Clientes
                    .Edit
                        If Not IsNull(.rdoColumns("Timbre_UUID")) Then
                            Codigo_UUID = .rdoColumns("Timbre_UUID")
                        End If
                        .rdoColumns("Fecha_Cancelo") = Now()
                        .rdoColumns("Motivo_Cancelo") = Trim(Motivo)
                        .rdoColumns("Usuario_Cancelo") = Nombre_Usuario
                        If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                            'Genera la estructura del xml y cancelacion del timbrado
                            If Lbl_Facturacion.Caption = "CANCELACIÓN EN PROCESO" Then
                                Mensaje = CFD_Cancela_Xml(Codigo_UUID, True)
                            Else
                                Mensaje = CFD_Cancela_Xml(Codigo_UUID, False)
                            End If
                            If Mensaje Like "*Cancelación exitosa*" Or Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                                Cancel = True
                            Else
                                Cancel = False
                            End If
                            lbl_estatus_cancel.Caption = Mensaje
                            lbl_estatus_cancel.Visible = True
                        End If
                        If Cancel And Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                            .rdoColumns("Cancelada") = "S"
                        ElseIf Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                            .rdoColumns("Cancelada") = "PC"
                        End If
                        .rdoColumns("Mensaje_Cancelado") = Mensaje
                    .Update
                End With
            Rs_Cancela_Factura_Clientes.Close
            If Cancel Or Cmb_Tipo_Factura.text <> "ELECTRONICA" Then
                For Cont_Fila = 1 To Grid_Detalle_Factura.Rows - 1 Step 1
                    'MODIFICA LA SALIDA
                    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Cmb_Salidas.text & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Fila, 4), "00000") & "' "
                    Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
                        With Rs_Modifica_Alm_Salidas_Almacen_Detalles
                            .Edit
                                .rdoColumns("Facturado") = "NO"
                                .rdoColumns("Precio_Venta") = 0
                                .rdoColumns("Importe") = 0
                                .rdoColumns("Total") = 0
                            .Update
                        End With
                    End If
                    Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
                Next
                'PONE COMO ACTIVA LA SALIDA
                Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen WHERE No_Salida ='" & Cmb_Salidas.text & "' "
                Set Rs_Modifica_Alm_Salidas_Almacen = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Not Rs_Modifica_Alm_Salidas_Almacen.EOF Then
                    With Rs_Modifica_Alm_Salidas_Almacen
                        .Edit
                            .rdoColumns("Estatus") = "RECEPCION"
                        .Update
                    End With
                End If
                Rs_Modifica_Alm_Salidas_Almacen.Close
            End If
        Else
            Mi_SQL = "SELECT No_Remision, Cancelada, Usuario_Cancelo, Fecha_Cancelo "
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones "
            Mi_SQL = Mi_SQL & " WHERE No_Remision = '" & Txt_No_Factura.text & "'"
            Set Rs_Cancela_Factura_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            'Llena la tabla de Adm_Clientes_Facturas con los datos contenidos en las cajas de textos
            With Rs_Cancela_Factura_Clientes
                .Edit
                    .rdoColumns("Fecha_Cancelo") = Now()
                    .rdoColumns("Usuario_Cancelo") = Nombre_Usuario
                    .rdoColumns("Cancelada") = "S"
                .Update
            End With
            Rs_Cancela_Factura_Clientes.Close
            For Cont_Fila = 1 To Grid_Detalle_Factura.Rows - 1 Step 1
                'MODIFICA LA SALIDA
                Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Cmb_Salidas.text & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Fila, 4), "00000") & "' "
                Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
                    With Rs_Modifica_Alm_Salidas_Almacen_Detalles
                        .Edit
                            .rdoColumns("Facturado") = "NO"
                            .rdoColumns("Precio_Venta") = 0
                            .rdoColumns("Importe") = 0
                            .rdoColumns("Total") = 0
                        .Update
                    End With
                End If
                Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
            Next
            'PONE COMO ACTIVA LA SALIDA
            Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen WHERE No_Salida ='" & Cmb_Salidas.text & "' "
            Set Rs_Modifica_Alm_Salidas_Almacen = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Modifica_Alm_Salidas_Almacen.EOF Then
                With Rs_Modifica_Alm_Salidas_Almacen
                    .Edit
                        .rdoColumns("Estatus") = "RECEPCION"
                    .Update
                End With
            End If
            Rs_Modifica_Alm_Salidas_Almacen.Close
        End If
    Conexion_Base.CommitTrans
    MDIFrm_Apl_Principal.MousePointer = 0
    If Cancel Then
        MsgBox "Documento Cancelado Existosamente", vbInformation
        Lbl_Facturacion.Caption = "CANCELADA"
    Else
        MsgBox "Documento en Proceso de Cancelación", vbInformation
        Lbl_Facturacion.Caption = "CANCELACION EN PROCESO"
    End If
    Btn_Cancelar.Enabled = False
    Btn_Imprimir.Enabled = False
    Btn_Enviar_Email.Enabled = False
    
    Fra_Datos_Cliente.Enabled = False
    Fra_Datos_Factura.Enabled = False
    Fra_Detalle_Factura.Enabled = False
    Fra_Comentarios.Enabled = False
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    'Obtiene el error
'    If Err.Number = 7777 Then
        MsgBox Err.Description
'    Else
'        For Each Rdo_Error In rdoErrors
'            MsgBox Rdo_Error.Description
'        Next
'    End If
End Sub

Private Sub Btn_Refacturar_Click()

    If Btn_Refacturar.Caption = "Refacturar" Then
        Fra_Datos_Cliente.Enabled = True
        Fra_Datos_Factura.Enabled = True
        Fra_Detalle_Factura.Enabled = True
        Fra_Comentarios.Enabled = True
        Btn_Imprimir.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Nuevo.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Cancelar.Enabled = False
        Btn_Salir.Caption = "Cancelar"
        Btn_Refacturar.Caption = "Actualizar"
        Cmb_Nombre_Cliente.Enabled = True
        Cmb_Nombre_Cliente.SetFocus
    Else
        Call Actualizar_Refacturacion
    End If
End Sub

Private Sub Btn_Relacionados_Click()

    If Cmb_FacRef.ListIndex > -1 Then
        Busca_UUID
        Grid_Relacionados.AddItem Cmb_Serie.text & Chr(9) & Cmb_FacRef.text & Chr(9) & Txt_UUID_Relacion.text
    End If
End Sub

'Botón para cerrar la forma
Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Grid_Detalle_Factura.Rows = 0
        Cmb_Nombre_Cliente.text = ""
        Cmb_Descripcion.text = ""
        Btn_Nuevo.Enabled = True
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Buscar.Enabled = True
        Btn_Salir.Caption = "Salir"
        Fra_Datos_Cliente.Enabled = False
        Fra_Datos_Factura.Enabled = False
        Fra_Detalle_Factura.Enabled = False
        Fra_Comentarios.Enabled = False
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Txt_Presio_Sin_IVA.Visible = False
        Chc_Adenda.Value = 0
        Grid_Relacionados.Enabled = False
        lbl_estatus_cancel.Visible = False
        Chc_Adenda.Value = 0
    End If
End Sub

Private Sub Btn_Sincronizar_Click()
    If Opt_Factura = True Or Opt_Remision = True Or Opt_Nota_Cargo = True Then
        Call Busca_Factura
        Fra_Busqueda.Visible = False
        Fra_Busqueda.Enabled = False
        Fra_Busqueda_Con_Controles.Enabled = False
    Else
        MsgBox "Seleccione un criterio de busqueda", vbInformation
    End If
End Sub


Private Sub Chc_Adenda_Click()
    Txt_Orden_Compra.text = ""
    Cmb_Tipo_Adenda.ListIndex = -1
    Txt_Plazo_Pago.text = ""
    If Chc_Adenda.Value > 0 Then
        Cmb_Tipo_Adenda.Enabled = True
    Else
        Cmb_Tipo_Adenda.Enabled = False
    End If
End Sub

Private Sub Chc_Relacionados_Click()
    Grid_Relacionados.Rows = 0
    Cmb_Relacionados.ListIndex = -1
    Cmb_FacRef.text = ""
    Cmb_Serie.ListIndex = -1
    If Chc_Relacionados.Value = 1 Then
        Cmb_Relacionados.Enabled = True
        Cmb_FacRef.Enabled = True
        Cmb_Serie.Enabled = True
        Grid_Relacionados.Enabled = True
    Else
        'Cmb_Relacionados.Index = -1
        'Txt_UUID_Relacion.text = ""
        Cmb_Relacionados.Enabled = False
        Cmb_FacRef.Enabled = False
        Cmb_Serie.Enabled = False
        Grid_Relacionados.Enabled = False
    End If
End Sub
Public Sub Carga_Facturas()
    Dim Rs_Consulta As rdoResultset
    Cmb_Serie.Clear
    Mi_SQL = "SELECT DISTINCT Serie FROM Adm_Clientes_Facturas WHERE Serie is not null"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta.EOF
         Cmb_Serie.AddItem (Rs_Consulta.rdoColumns("Serie"))
         Rs_Consulta.MoveNext
    Wend
    Rs_Consulta.Close
    Cmb_FacRef.Clear
    Mi_SQL = "SELECT No_Factura_Electronica FROM Adm_Clientes_Facturas WHERE Cliente_ID = '" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "#00000") & "' and Timbre_UUID is not null and Timbre_UUID<>'' "
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta.EOF
         Cmb_FacRef.AddItem (Rs_Consulta.rdoColumns("No_Factura_Electronica"))
         Rs_Consulta.MoveNext
    Wend
     Rs_Consulta.Close
End Sub
Public Sub Busca_UUID()
    Dim Rs_Consulta As rdoResultset
    Txt_UUID_Relacion.text = ""
    Mi_SQL = "SELECT Timbre_UUID FROM Adm_Clientes_Facturas WHERE Serie='" & Cmb_Serie.text & "' and No_Factura_Electronica='" & Cmb_FacRef.text & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
         Txt_UUID_Relacion.text = Rs_Consulta.rdoColumns("Timbre_UUID")
    End If
    Rs_Consulta.Close
End Sub



'Si cambia con el teclado la descrpcion se limpia los datos del producto
Private Sub Cmb_Descripcion_Change()
    'Call Limpia_Datos_Producto
End Sub

'Al dar clic en el combo descripción éste muestra los datos del producto seleccionado
Private Sub Cmb_Descripcion_Click()
Dim Mi_SQL As String                            'Obtiene los valores de la consulta
Dim Rs_Consulta_Alm_Salidas_Almacen_Detalles As rdoResultset   'Manejo de registro
Set Conectar_Ayudante = New Ayudante

On Error GoTo handler
    'Consulta
    Mi_SQL = "SELECT * FROM Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Producto_ID = '" & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & "'"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Revisa que no sea fin de archivo para llenar el grid
        With Rs_Consulta_Alm_Salidas_Almacen_Detalles
            If Not .EOF Then
                Txt_Cantidad.Enabled = True
                Txt_Precio.Enabled = True
                Txt_Importe.Enabled = False
                'Limpia_Pantalla_Facturacion (2)
                Cmb_Descripcion.text = .rdoColumns("Nombre")
'                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Presentacion_ID"), Cmb_Unidad)
                If Not IsNull(.rdoColumns("Precio_Venta")) Then
                    If Val(.rdoColumns("Precio_Venta")) = 0 Then
                        Txt_Precio.text = ""
                    Else
                        Txt_Precio.text = .rdoColumns("Costo")
                    End If
                    Txt_Aplica_IVA.text = .rdoColumns("Aplica_IVA")
                Else
                    Txt_Precio.text = ""
                End If
            Else
                Exit Sub
            End If
        End With
    Rs_Consulta_Alm_Salidas_Almacen_Detalles.Close
    Txt_Importe.text = (Val(Txt_Cantidad.text) * Val(Txt_Precio.text))
    Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'Al dar clic en el combo éste se llena con los datos del catálogo de productos
Private Sub Cmb_Descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Set Conectar_Ayudante = New Ayudante
        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Descripcion, 1, "Nombre")
    End If
End Sub

'Al dar clic en el combo descripción_sat éste muestra los datos del producto seleccionado
'Private Sub Cmb_Descripcion_Sat_Click()
'Dim Mi_SQL As String                            'Obtiene los valores de la consulta
'Dim Rs_Consulta_Alm_Salidas_Almacen_Detalles As rdoResultset   'Manejo de registro
'Set Conectar_Ayudante = New Ayudante

'On Error GoTo handler
'    'Consulta
'    Mi_SQL = "SELECT * FROM Cat_Productos_Servicios"
'    Mi_SQL = Mi_SQL & " WHERE Clave_Producto_Servicio = '" & Format(Cmb_Descripcion_Sat.ItemData(Cmb_Descripcion_Sat.ListIndex), "00000000") & "'"
'    'Le envía la consulta al ayudante para que realice la consulta
'    Set Rs_Consulta_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'        'Revisa que no sea fin de archivo para llenar el grid
'        With Rs_Consulta_Alm_Salidas_Almacen_Detalles
'            If Not .EOF Then
'                Cmb_Descripcion.text = .rdoColumns("Descripcion")
'            Else
'                Exit Sub
'            End If
'        End With
'    Rs_Consulta_Alm_Salidas_Almacen_Detalles.Close
'    Exit Sub
'handler:
'    For Each Er In rdoErrors
'        MsgBox Er.Description
'    Next Er
'End Sub

'Al dar clic en el combo éste se llena con los datos del catálogo de productos
Private Sub Cmb_Descripcion_Sat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Set Conectar_Ayudante = New Ayudante
        'Call Conectar_Ayudante.Llena_Combo_Item("Top 32000 Clave_Producto_Servicio,Descripcion", "Cat_Productos_Servicios", Cmb_Descripcion_Sat, 1, "Descripcion")
         Call Conectar_Ayudante.Llena_Combo_Item("Top 32000 Clave,Clave_Producto_Servicio + ' ' + Descripcion AS Cadena", "Cat_Productos_Servicios", Cmb_Descripcion_Sat, 1, "Descripcion")
         If Cmb_Descripcion_Sat.text <> "" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Top 32000 Clave,Clave_Producto_Servicio + ' ' + Descripcion AS Cadena", "Cat_Productos_Servicios", Cmb_Descripcion_Sat, 1, "Clave_Producto_Servicio")
         End If
    End If
End Sub






Private Sub Cmb_FacRef_KeyPress(KeyAscii As Integer)
Dim i As Integer
    If KeyAscii = 13 And Cmb_Nombre_Cliente.text <> "" Then
        If Grid_Relacionados.Rows = 0 Then
            Grid_Relacionados.AddItem "Serie" & Chr(9) & "Folio" & Chr(9) & "UUID"
            Grid_Relacionados.ColWidth(0) = 500        'Serie
            Grid_Relacionados.ColWidth(1) = 500        'Folio
            Grid_Relacionados.ColWidth(2) = 2000        'UUID
        End If
        If Cmb_FacRef.ListIndex > -1 Then
            Busca_UUID
            For i = 1 To Grid_Relacionados.Rows - 1
                If Txt_UUID_Relacion.text = Grid_Relacionados.TextMatrix(i, 2) Then
                    MsgBox "La factura referenciada ya fue agregada"
                    Exit Sub
                End If
            Next i
            If Txt_UUID_Relacion.text <> "" Then
                Grid_Relacionados.AddItem Cmb_Serie.text & Chr(9) & Val(Cmb_FacRef.text) & Chr(9) & Txt_UUID_Relacion.text
            Else
                MsgBox "No se encontró el UUID de la factura referenciada"
                Cmb_FacRef.text = ""
                Exit Sub
            End If
        End If
    End If
    Cmb_FacRef.text = ""
End Sub

Private Sub Cmb_Metodo_Pago_Click()
    'Txt_Cuenta_Pago.text = ""
    'Txt_Cuenta_Pago.Locked = True
    'Txt_Cuenta_Pago.TabStop = False
    'If Cmb_Metodo_Pago.ListIndex > -1 Then
    '    If Cmb_Metodo_Pago.ListIndex > 1 Then
    '        Txt_Cuenta_Pago.Locked = False
    '        Txt_Cuenta_Pago.TabStop = True
    '    End If
    'End If
End Sub

'Función para que cuando cambie el texto del combo limpie los campos de los clientes
Private Sub Cmb_Nombre_Cliente_Change()
    Limpia_Datos_Clientes
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Limpia_Datos_Clientes
'DESCRIPCIÓN: Limpia las cajas de texto al cambiar el nombre del cliente
'PARÁMETROS:
'CREO:
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Limpia_Datos_Clientes()
    Txt_Cliente_ID.text = ""
    Txt_RFC_Cliente.text = ""
    Txt_Direccion_Cliente.text = ""
    Txt_No_Exterior.text = ""
    Txt_No_Interior.text = ""
    Txt_Colonia_Cliente.text = ""
    Txt_Ciudad_Cliente.text = ""
    Txt_Estado.text = ""
    Txt_Pais.text = ""
    Txt_Dias_Credito.text = ""
    Txt_Telefono_Cliente.text = ""
    Txt_Codigo_Postal.text = ""
    Txt_Cuenta_Pago.text = ""
    Cmb_Metodo_Pago.ListIndex = -1
    Grid_Relacionados.Rows = 0
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Limpia_Datos_Producto
'DESCRIPCIÓN: Cambia a 0 las cajas de texto al cambiar el nombre del cliente
'PARÁMETROS:
'CREO:
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Limpia_Datos_Producto()
    Txt_Aplica_IVA.text = ""
    Txt_Cantidad.text = ""
    Txt_Precio.text = ""
    Txt_Importe.text = ""
End Sub

'Al dar clic en el combo éste muestra los datos del cliente seleccionado
Private Sub Cmb_Nombre_Cliente_Click()
Dim Mi_SQL As String                                                 'Guarda la consulta
Dim Rs_Consulta_Cat_Clientes As rdoResultset                         'Manejo del registro de la tabla Cat_Clientes
Dim Rs_Consulta_Salidas As rdoResultset
Dim Rs_Consulta_Salidas_Detalles As rdoResultset

On Error GoTo handler
    Grid_Relacionados.Rows = 0
    'Realiza la consulta para enviarla al recordset
    Mi_SQL = "SELECT Credito_Flexible,Cliente_ID, Cat_Clientes.Nombre as Nombre_C, Dias_Credito, RFC, Direccion, No_Ext,"
    Mi_SQL = Mi_SQL & " No_Int, Colonia, Ciudad, Estado, Pais, Telefono, CP, Metodo_Pago, Cuenta_Pago, Email, Tipo_Persona"
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000") & "'"
    'Le asigna la consulta al recordsert de consulta
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Clientes
        If Not .EOF Then
            'Le asigna los valores de la tabla a los campos en la forma
            Cmb_Nombre_Cliente.text = .rdoColumns("Nombre_C")
            If Not IsNull(.rdoColumns("RFC")) Then
                Txt_RFC_Cliente.text = .rdoColumns("RFC")
            End If
            If Not IsNull(.rdoColumns("Direccion")) Then Txt_Direccion_Cliente.text = .rdoColumns("Direccion")
            If Not IsNull(.rdoColumns("No_Ext")) Then
                Txt_No_Exterior.text = Trim(.rdoColumns("No_Ext"))
            Else
                Txt_No_Exterior.text = ""
            End If
            If Not IsNull(.rdoColumns("No_Int")) Then
                Txt_No_Interior.text = Trim(.rdoColumns("No_Int"))
            Else
                Txt_No_Interior.text = ""
            End If
            If Not IsNull(.rdoColumns("Colonia")) Then
                Txt_Colonia_Cliente.text = .rdoColumns("Colonia")
            Else
                Txt_Colonia_Cliente.text = ""
            End If
            If Not IsNull(.rdoColumns("Ciudad")) Then Txt_Ciudad_Cliente.text = .rdoColumns("Ciudad")
            If Not IsNull(.rdoColumns("Estado")) Then
                Txt_Estado.text = .rdoColumns("Estado")
            Else
                Txt_Estado.text = ""
            End If
            If Not IsNull(.rdoColumns("Pais")) Then
                Txt_Pais.text = .rdoColumns("Pais")
            Else
                Txt_Pais.text = ""
            End If
            If Not IsNull(.rdoColumns("Email")) Then
                Txt_Email.text = .rdoColumns("Email")
            Else
                Txt_Email.text = ""
            End If
            If Not IsNull(.rdoColumns("Telefono")) Then Txt_Telefono_Cliente.text = .rdoColumns("Telefono")
            If Not IsNull(.rdoColumns("Cliente_ID")) Then Txt_Cliente_ID.text = .rdoColumns("Cliente_ID")
            If Not IsNull(.rdoColumns("Dias_Credito")) Then Txt_Dias_Credito.text = .rdoColumns("Dias_Credito")
            If Not IsNull(.rdoColumns("CP")) Then Txt_Codigo_Postal.text = .rdoColumns("CP")
            If Not IsNull(.rdoColumns("Metodo_Pago")) Then
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Metodo_Pago"), Cmb_Metodo_Pago)
            Else
                Cmb_Metodo_Pago.ListIndex = -1
            End If
            If Not IsNull(.rdoColumns("Cuenta_Pago")) Then
                Txt_Cuenta_Pago.text = .rdoColumns("Cuenta_Pago")
            Else
                Txt_Cuenta_Pago.text = ""
            End If
            If Not Btn_Nuevo.Caption = "Nuevo" Then DTP_Fecha_Pago.Value = DateAdd("d", Val(Txt_Dias_Credito.text), DTP_Fecha_Factura.Value)
            If Not IsNull(.rdoColumns("Credito_Flexible")) Then
                If Trim(.rdoColumns("Credito_Flexible")) = "SI" Then
                    Txt_Dias_Credito.Locked = False
                    Lbl_Dias_Credito.Caption = "Cred Flex"
                Else
                    Txt_Dias_Credito.Locked = True
                    Lbl_Dias_Credito.Caption = "Días Créd"
                End If
            Else
                Txt_Dias_Credito.Locked = True
                Lbl_Dias_Credito.Caption = "Días Créd"
            End If
            
            Cmb_Uso_CFDI.Clear
            If .rdoColumns("Tipo_Persona") = "MORAL" Then
                Mi_SQL = "SELECT * FROM Cat_Uso_Comprobantes  ORDER BY Clave"  'Where Persona_Moral='SI' ORDER BY Clave"
                Set Rs_Consulta_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                While Not Rs_Consulta_Comprobantes.EOF
                    If Rs_Consulta_Comprobantes.rdoColumns("Persona_Moral") = "SI" Then
                        Cmb_Uso_CFDI.AddItem UCase(Rs_Consulta_Comprobantes.rdoColumns("Codigo_Uso_Comprobante") & " " & Rs_Consulta_Comprobantes.rdoColumns("Descripcion"))
                    End If
                    Rs_Consulta_Comprobantes.MoveNext
                Wend
                Else
                    Mi_SQL = "SELECT * FROM Cat_Uso_Comprobantes  ORDER BY Clave"  'Where Persona_Moral='SI' ORDER BY Clave"
                    Set Rs_Consulta_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    While Not Rs_Consulta_Comprobantes.EOF
                        Cmb_Uso_CFDI.AddItem UCase(Rs_Consulta_Comprobantes.rdoColumns("Codigo_Uso_Comprobante") & " " & Rs_Consulta_Comprobantes.rdoColumns("Descripcion"))
                        Rs_Consulta_Comprobantes.MoveNext
                    Wend
        
             End If
             Rs_Consulta_Comprobantes.Close
        Else
            Exit Sub
        End If
    End With
    'SE CONSULTAN LAS SALIDAS RELACIONADS CON EL CLIENTE
    Mi_SQL = "SELECT Distinct Alm_Salidas_Almacen.No_Salida FROM Alm_Salidas_Almacen,Alm_Salidas_Almacen_Detalles WHERE Alm_Salidas_Almacen.Cliente_ID ='" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Tipo_Salida = 'VENTA'"
    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Estatus = 'RECEPCION'"
    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.No_Salida = Alm_Salidas_Almacen.No_Salida "
    ''Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.Facturado ='NO'"
    Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Limpia el combo
    If Btn_Nuevo.Caption <> "Nuevo" Then
        Cmb_Salidas.Clear
        'Consulta las salidas
        If Not Rs_Consulta_Salidas.EOF Then
            With Rs_Consulta_Salidas
                While Not Rs_Consulta_Salidas.EOF
                    Cmb_Salidas.AddItem Rs_Consulta_Salidas!No_Salida
                    Cmb_Salidas.ItemData(Cmb_Salidas.NewIndex) = Format(Rs_Consulta_Salidas!No_Salida, "00000")
                    Rs_Consulta_Salidas.MoveNext
                Wend
            End With
        End If
    End If
    Rs_Consulta_Salidas.Close
    Carga_Facturas
    Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'Al dar enter en el combo se llena con los datos del catálogo
Private Sub Cmb_Nombre_Cliente_KeyPress(KeyAscii As Integer)
  Dim Despliega_Lista As Long

    If KeyAscii = 13 Then
        Set Conectar_Ayudante = New Ayudante
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Nombre_Cliente, 1, "Nombre")
        Cmb_Nombre_Cliente.text = ""
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Nombre_Cliente.hwnd, &H14F, True, 0)
    End If
End Sub







Private Sub Cmb_Tipo_Adenda_Click()
     If Cmb_Tipo_Adenda.text = "NADRO" Then
        Txt_Orden_Compra.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = True
        Label4.Enabled = True
        Txt_Plazo_Pago.Enabled = True
    Else
        Txt_Plazo_Pago.text = ""
        Txt_Orden_Compra.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = False
        Txt_Plazo_Pago.Enabled = False
    End If
End Sub

Private Sub Cmb_Tipo_Factura_Click()
Dim Rs_Consulta_Serie As rdoResultset

    If Btn_Nuevo.Caption = "Dar de Alta" Then
        If Cmb_Tipo_Factura.ListIndex > -1 Then
            If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                If Cmb_Tipo_Documento.text = "FACTURA" Then
                    'Consulta la serie del rango activo actual
                    Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
                    Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
                    Mi_SQL = Mi_SQL & " AND Tipo = 'FACTURA'"
                    Set Rs_Consulta_Serie = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Serie.EOF Then
                            Txt_Serie.text = Trim(Rs_Consulta_Serie.rdoColumns("Serie"))
                        End If
                    Rs_Consulta_Serie.Close
                    'Valida si aun existen folios de facturas disponibles para utilizar
                    Call Aviso_Termino_Folios("FACTURA")
                    'si la bandera esta habilitada muestra mensaje y cancela la operación
                    If Folios_Terminados = True Then
                        Txt_Factura_ID.text = ""
                        Txt_No_Factura.text = ""
                        MsgBox "No se encontraron folios de factura disponibles, favor de verificar", vbCritical
                        Btn_Salir_Click
                        Exit Sub
                    Else
                        Txt_Factura_ID.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
                        Txt_No_Factura.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas WHERE Forma_Factura = 'E'", "No_Factura_Electronica")
                    End If
                Else
                    'Consulta la serie del rango activo actual
                    Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
                    Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
                    Mi_SQL = Mi_SQL & " AND Tipo = 'FACTURA'"
                    Mi_SQL = Mi_SQL & " AND Serie = 'NCA'"
                    Set Rs_Consulta_Serie = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Serie.EOF Then
                            Txt_Serie.text = Trim(Rs_Consulta_Serie.rdoColumns("Serie"))
                        End If
                    Rs_Consulta_Serie.Close
                    'Valida si aun existen folios de facturas disponibles para utilizar
                    Call Aviso_Termino_Folios("NOTA CARGO")
                    'si la bandera esta habilitada muestra mensaje y cancela la operación
                    If Folios_Terminados = True Then
                        Txt_Factura_ID.text = ""
                        Txt_No_Factura.text = ""
                        MsgBox "No se encontraron folios de nota de cargo disponibles, favor de verificar", vbCritical
                        Btn_Salir_Click
                        Exit Sub
                    Else
                        Txt_Factura_ID.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
                        Txt_No_Factura.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas WHERE Forma_Factura = 'E' AND Serie='NCA'", "No_Factura_Electronica")
                    End If
                
                End If
                
            Else
                Txt_Factura_ID.text = ""
                Txt_Serie.text = ""
                Txt_No_Factura.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
            End If
        End If
    End If
End Sub

'Si elige que la factura se va a pagar de contado la fecha de pago es la actual
Private Sub Cmb_Tipo_Pago_Click()
    If Cmb_Tipo_Pago.text = "CONTADO" Then
        DTP_Fecha_Pago.Value = Now
    Else
        DTP_Fecha_Pago.Value = DateAdd("d", Val(Txt_Dias_Credito.text), DTP_Fecha_Factura.Value)
    End If
End Sub

Private Sub Cmb_Salidas_Click()
    Select Case Trim(Cmb_Tipo_Documento.text)
        Case "REMISION"
            Txt_No_Factura.text = Cmb_Salidas.text
    End Select
End Sub

Private Sub Cmb_Tipo_Documento_Click()
    Cmb_Tipo_Factura.ListIndex = -1
    Cmb_Tipo_Factura.Enabled = False
    Grid_Detalle_Factura.Rows = 0
    Txt_Factura_ID.text = ""
    Txt_No_Factura.text = ""
    Txt_Serie.text = ""
    If Cmb_Tipo_Documento.ListIndex > -1 Then
        If Cmb_Tipo_Documento.text = "REMISION" And Btn_Nuevo.Caption <> "Nuevo" Then
            Me.Caption = "REMISION CLIENTES"
            Txt_No_Factura.text = Cmb_Salidas.text
'            Grid_Detalle_Factura.Rows = 0
        Else
            If Cmb_Tipo_Documento.text = "FACTURA" And Btn_Nuevo.Caption <> "Nuevo" Then
                Me.Caption = "FACTURAS CLIENTES"
                Cmb_Tipo_Factura.Enabled = True
'                Txt_No_Factura.Text = Conectar_Ayudante.Maximo_Catalogo("Adm_Clientes_Facturas", "No_Factura")
'                Grid_Detalle_Factura.Rows = 0

            Else
                If Cmb_Tipo_Documento.text = "NOTA CARGO" And Btn_Nuevo.Caption <> "Nuevo" Then
                    Me.Caption = "NOTA DE CARGO CLIENTES"
'                    Cmb_Tipo_Factura.text = "ELECTRONICA"
                    Cmb_Tipo_Factura.ListIndex = 0
                    Cmb_Tipo_Factura.Enabled = True
                    
                End If
            End If
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Nombre_Cliente, 1, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Clave_Unidad + '-' + Nombre ", "Cat_Unidades_Medida", Cmb_Unidad, 1, "Nombre")
'    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Nombre", "Cat_Unidades_Medida", Cmb_Unidad, 1, "Nombre")
    'Call Conectar_Ayudante.Llena_Combo_Item("Clave,Descripcion", "Cat_Regimen_Fiscal", Cmb_Regimen_Fiscal, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Metodo_ID + '' + Descripcion ", "Cat_Metodo_Pago", Cmb_Metodo_Pago, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Codigo_Tipo_Relacion + '' + Descripcion", "Cat_Tipos_Relacion", Cmb_Relacionados, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Codigo_Uso_Comprobante + '' + Descripcion as Descripcion", "Cat_Uso_Comprobantes", Cmb_Uso_CFDI, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Clave + '' + Descripcion", "Cat_Formas_Pago", Cmb_Forma_Pago, 1, "Descripcion")
    Limpia_Variables
    Call Cmb_Nombre_Cliente_KeyPress(13)
    Call Cmb_Descripcion_KeyPress(13)
    Consulta_Cancelados
End Sub

Private Sub Grid_Detalle_Factura_Click()
    
    If Grid_Detalle_Factura.Rows > 0 And Grid_Detalle_Factura.ColSel > 2 Then
        Grid_Detalle_Factura.SelectionMode = flexSelectionFree
        Grid_Detalle_Factura.Refresh
        'En notas de cargo no utiliza el campo de modificacion de cantidad
        If Val(Grid_Detalle_Factura.TextMatrix(Grid_Detalle_Factura.RowSel, 0)) > 0 Then
            If Grid_Detalle_Factura.Rows > 1 Then
                Txt_Presio_Sin_IVA.text = Grid_Detalle_Factura.TextMatrix(Grid_Detalle_Factura.RowSel, 4)
            End If
            If Grid_Detalle_Factura.Rows = 1 Then
                Txt_Presio_Sin_IVA.text = Grid_Detalle_Factura.TextMatrix(Grid_Detalle_Factura.RowSel, 1)
            End If
            If Grid_Detalle_Factura.Rows <= 1 Then Exit Sub
            If ((Grid_Detalle_Factura.Col = 4)) Or ((Grid_Detalle_Factura.Col = 2)) Then
                Call Mover_Control_Grid_TextBox(Grid_Detalle_Factura, Txt_Presio_Sin_IVA)
            Else
                Txt_Presio_Sin_IVA.Visible = False
            End If
        Else
            If Grid_Detalle_Factura.ColSel = 2 Then
                MsgBox "La partida no tiene cantidad"
                Txt_Presio_Sin_IVA.Visible = False
            Else
               If Grid_Detalle_Factura.ColSel = 1 Then
                    Txt_Presio_Sin_IVA.text = Grid_Detalle_Factura.TextMatrix(Grid_Detalle_Factura.RowSel, 1)
                    If Grid_Detalle_Factura.Rows <= 1 Then Exit Sub
                    Call Mover_Control_Grid_TextBox(Grid_Detalle_Factura, Txt_Presio_Sin_IVA)
                    Txt_Presio_Sin_IVA.Visible = True
               End If
            End If
        End If
    End If
End Sub

Private Sub Grid_Detalle_Factura_EnterCell()
    If (Grid_Detalle_Factura.Col = 4 Or Grid_Detalle_Factura.Col = 2) And Grid_Detalle_Factura.Rows > 1 Then
        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Detalle_Factura, Txt_Presio_Sin_IVA)
    End If
End Sub

Private Sub Grid_Detalle_Factura_LeaveCell()
  Grid_Detalle_Factura.CellBackColor = vbWhite
End Sub



Private Sub Opt_Factura_Click()
    Cmb_Consulta_Tipo_Factura.Visible = False
    Cmb_Consulta_Tipo_Factura.ListIndex = -1
    If Opt_Factura.Value = True Then
        Cmb_Consulta_Tipo_Factura.Visible = True
        Cmb_Consulta_Tipo_Factura.ListIndex = 0
    End If
End Sub

Private Sub Opt_Nota_Cargo_Click()
    Cmb_Consulta_Tipo_Factura.Visible = False
    Cmb_Consulta_Tipo_Factura.ListIndex = -1
    If Opt_Nota_Cargo.Value = True Then
        Cmb_Consulta_Tipo_Factura.Visible = True
        Cmb_Consulta_Tipo_Factura.ListIndex = 0
    End If
End Sub

Private Sub Opt_Remision_Click()
    If Opt_Remision.Value = True Then
        Cmb_Consulta_Tipo_Factura.Visible = False
        Cmb_Consulta_Tipo_Factura.ListIndex = -1
    End If
End Sub

'Función para que cuando cambie el texto de la cantidad calcule automáticamente el importe
Private Sub Txt_Cantidad_Change()
    Txt_Importe.text = (Val(Txt_Cantidad.text) * Val(Txt_Precio.text))
End Sub

Private Sub Txt_Cantidad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cantidad.text, True)
End Sub

Private Sub Txt_Dias_Credito_Change()
    If Not Btn_Nuevo.Caption = "Nuevo" Then
        DTP_Fecha_Pago.Value = DateAdd("d", Val(Txt_Dias_Credito.text), DTP_Fecha_Factura.Value)
    End If
End Sub

'Función para que calcule la cantidad automáticamente al ingresar el precio manualmente
Private Sub Txt_Precio_Change()
    Txt_Importe.text = (Val(Txt_Cantidad.text) * Val(Txt_Precio.text))
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Busca_Factura
'DESCRIPCIÓN: Busca la factura de acuerdo al número proporcionado
'PARÁMETROS:
'CREO:
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Busca_Factura()
Dim Mi_SQL As String
Dim Rs_Consulta_Factura_Clientes As rdoResultset        'Manejo del registro para buscar facturas
Dim Rs_Consulta_Descripcion_Facturas As rdoResultset    'Manejo del registro para el detalle de la factura
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo del registro del catálogo de clientes
Dim Rs_Consulta_Movimientos_Facturas As rdoResultset    'Manejo del registro de movimientos
Dim RS_Consulta_Relacionados As rdoResultset
Dim No_Factura As String                                'Variable para capturar el número de la factura a buscar
Dim Suma As Double                  'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double              'Suma I.V.A.
Dim Electronica As Boolean
Dim Unidad As String
Dim i As Integer
    
    No_Factura = InputBox("Teclee el número de Documento a consultar", "Consulta de Documentos")
    If No_Factura <> "" Then
        Electronica = False
        If Opt_Factura = True Then 'SI LA OPCION ES FACTURA
            'Prepara el recordset para consultar el número de factura de la tabla Adm_Clientes_Facturas
            Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas"
            If Mid(Cmb_Consulta_Tipo_Factura.text, 1, 1) = "E" Then
                Mi_SQL = Mi_SQL & " WHERE No_Factura_Electronica = '" & Format(No_Factura, "0000000000") & "'"
                Electronica = True
            Else
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(No_Factura, "0000000000") & "'"
            End If
            Set Rs_Consulta_Factura_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Else
            If Opt_Nota_Cargo = True Then 'SI LA OPCION ES NOTA DE CARGO
                'Prepara el recordset para consultar el número de factura de la tabla Adm_Clientes_Facturas
                Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas"
                If Mid(Cmb_Consulta_Tipo_Factura.text, 1, 1) = "E" Then
                    Mi_SQL = Mi_SQL & " WHERE No_Factura_Electronica = '" & Format(No_Factura, "0000000000") & "'"
                    Electronica = True
                Else
                    Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(No_Factura, "0000000000") & "'"
                End If
                Mi_SQL = Mi_SQL & " AND Serie = 'NCA'"
                Set Rs_Consulta_Factura_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Else
                'SI LA OPCION ES REMISION
                'Prepara el recordset para consultar el número de factura de la tabla Adm_Clientes_Facturas
                Mi_SQL = "SELECT * FROM Adm_Clientes_Remisiones WHERE No_Remision = '" & Format(No_Factura, "0000000000") & "'"
                Set Rs_Consulta_Factura_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            End If
        End If
        
        'Llena los controles con los datos de la consulta
        If Not Rs_Consulta_Factura_Clientes.EOF Then
            If Not IsNull(Rs_Consulta_Factura_Clientes!Tipo_Documento) Then
                Select Case Trim(Rs_Consulta_Factura_Clientes!Tipo_Documento)
                Case "FACTURA"
                    Cmb_Tipo_Documento.text = "FACTURA"
                Case "NOTA CARGO"
                    Cmb_Tipo_Documento.text = "NOTA CARGO"
                Case "REMISION"
                    Cmb_Tipo_Documento.text = "REMISION"
                End Select
            Else
                Cmb_Tipo_Documento.text = "FACTURA"
            End If
            
            If Opt_Factura = True Or Opt_Nota_Cargo = True Then
                If Electronica = True Then
                    Txt_No_Factura.text = Val(Rs_Consulta_Factura_Clientes!No_Factura_Electronica)
                    Txt_Factura_ID.text = Rs_Consulta_Factura_Clientes!No_Factura
                    Txt_Serie.text = Rs_Consulta_Factura_Clientes!Serie
                    Cmb_Tipo_Factura.text = "ELECTRONICA"
                    If Not IsNull(Rs_Consulta_Factura_Clientes!Mensaje_Cancelado) Then
                        lbl_estatus_cancel.Caption = Rs_Consulta_Factura_Clientes!Mensaje_Cancelado
                        lbl_estatus_cancel.Visible = True
                    End If
                Else
                    Txt_No_Factura.text = Rs_Consulta_Factura_Clientes!No_Factura
                    Txt_Factura_ID.text = ""
                    Txt_Serie.text = ""
                    Cmb_Tipo_Factura.text = "PAPEL"
                End If
                DTP_Fecha_Factura.Value = Rs_Consulta_Factura_Clientes!Fecha
                DTP_Fecha_Pago.Value = Rs_Consulta_Factura_Clientes!Fecha_Pago
                'If Trim(Rs_Consulta_Factura_Clientes!Tipo_Pago) = "CONTADO" Then
                '    Opt_Contado = True
                'Else
                '    Opt_Credito = True
                'End If
            Else
                Txt_No_Factura.text = Rs_Consulta_Factura_Clientes!No_Remision
                DTP_Fecha_Factura.Value = Rs_Consulta_Factura_Clientes!Fecha
                DTP_Fecha_Pago.Value = Rs_Consulta_Factura_Clientes!Fecha_Pago
                'If Trim(Rs_Consulta_Factura_Clientes!Tipo_Pago) = "CONTADO" Then
                '    Opt_Contado = True
                'Else
                '    Opt_Credito = True
                'End If
            End If
            If Not IsNull(Rs_Consulta_Factura_Clientes!Orden_Compra) Then
                Chc_Adenda.Value = 1
                Txt_Orden_Compra.text = Rs_Consulta_Factura_Clientes!Orden_Compra
            Else
                Chc_Adenda.Value = 0
            End If
            
            
            If Not IsNull(Rs_Consulta_Factura_Clientes!No_Salida) Then
                Cmb_Salidas.text = Rs_Consulta_Factura_Clientes!No_Salida
            Else
                Cmb_Salidas.text = ""
            End If
            If Not IsNull(Rs_Consulta_Factura_Clientes!Tipo_Factura) Then
                If Rs_Consulta_Factura_Clientes!Tipo_Factura = "NORMAL" Then
                    Cmb_Tipo_Factura.ListIndex = 0
                Else
                    Cmb_Tipo_Factura.ListIndex = 1
                End If
            End If
            
            
            'Consulta del cliente de la factura seleccionada
            Mi_SQL = "SELECT Cliente_ID, Cat_Clientes.Nombre as Nombre_C, Dias_Credito, RFC, Direccion, Colonia, Ciudad, Telefono, Estado, CP"
            Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID ='" & Format(Rs_Consulta_Factura_Clientes!Cliente_ID, "00000") & "'"
            Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                'Llena los controles con los datos de la búsqueda
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Call Conectar_Ayudante.Asigna_Item_Combo(Rs_Consulta_Cat_Clientes!Nombre_C, Cmb_Nombre_Cliente)
                    ''Cmb_Nombre_Cliente.Text = Rs_Consulta_Cat_Clientes!Nombre_C
                    If Not IsNull(Rs_Consulta_Cat_Clientes!RFC) Then Txt_RFC_Cliente.text = Rs_Consulta_Cat_Clientes!RFC
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Direccion) Then Txt_Direccion_Cliente.text = Rs_Consulta_Cat_Clientes!Direccion
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Colonia) Then Txt_Colonia_Cliente.text = Rs_Consulta_Cat_Clientes!Colonia
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Ciudad) Then Txt_Ciudad_Cliente.text = Rs_Consulta_Cat_Clientes!Ciudad & ", " & Rs_Consulta_Cat_Clientes!Estado
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Dias_Credito) Then Txt_Dias_Credito.text = Rs_Consulta_Cat_Clientes!Dias_Credito
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Telefono) Then Txt_Telefono_Cliente.text = Rs_Consulta_Cat_Clientes!Telefono
                    If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("CP")) Then Txt_Codigo_Postal.text = Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                End If
                Txt_Cliente_ID.text = Rs_Consulta_Factura_Clientes!Cliente_ID
                If Not IsNull(Rs_Consulta_Factura_Clientes!Comentarios) Then Txt_Comentarios.text = Rs_Consulta_Factura_Clientes!Comentarios
            Rs_Consulta_Cat_Clientes.Close
            'Asigna datos a combos
            For i = 0 To Cmb_Uso_CFDI.ListCount - 1
                If Mid(Cmb_Uso_CFDI.List(i), 1, 3) = Mid(Rs_Consulta_Factura_Clientes!Uso_CFDI, 1, 3) Then
                    Cmb_Uso_CFDI.ListIndex = i
                    Exit For
                End If
            Next i

            For i = 0 To Cmb_Forma_Pago.ListCount - 1
                If Mid(Cmb_Forma_Pago.List(i), 1, 2) = Mid(Rs_Consulta_Factura_Clientes!Forma_Pago, 1, 2) Then
                    Cmb_Forma_Pago.ListIndex = i
                    Exit For
                End If
            Next i
            
            For i = 0 To Cmb_Metodo_Pago.ListCount - 1
                If Mid(Cmb_Metodo_Pago.List(i), 1, 3) = Mid(Rs_Consulta_Factura_Clientes!Tipo_Pago, 1, 3) Then
                    Cmb_Metodo_Pago.ListIndex = i
                    Exit For
                End If
            Next i
            
            'Si la factura está pagada muestra el botón de imprimir
            If Trim(Rs_Consulta_Factura_Clientes!Pagada) = "S" Then
                Btn_Imprimir.Caption = "Imprimir"
                Btn_Cancelar.Enabled = False
                Lbl_Facturacion.Caption = "PAGADA"
                Btn_Enviar_Email.Enabled = False
            Else
                If Trim(Rs_Consulta_Factura_Clientes!cancelada) = "N" And Trim(Rs_Consulta_Factura_Clientes!cancelada) <> "PC" Then
                   'Consulta para ver si hay movimientos aplicados a las facturas
                    Mi_SQL = "SELECT No_Movimiento,No_Factura, Estatus, Cantidad "
                    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos "
                    If Electronica = True Then
                        Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Txt_Factura_ID.text & "'"
                    Else
                        Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Txt_No_Factura.text & "'"
                    End If
                    
                    Mi_SQL = Mi_SQL & " AND Estatus = 'A'"
                    Set Rs_Consulta_Movimientos_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Movimientos_Facturas.EOF Then
                            Btn_Cancelar.Enabled = False
                            Lbl_Facturacion.Caption = "ACTIVA CON PAGOS"
                        Else
                            Btn_Cancelar.Enabled = True
                            Lbl_Facturacion.Caption = "ACTIVA"
                        End If
                    Rs_Consulta_Movimientos_Facturas.Close
                    If Cmb_Tipo_Factura.text = "ELECTRONICA" Then
                        Btn_Enviar_Email.Enabled = True
                    Else
                        Btn_Enviar_Email.Enabled = False
                    End If
                ElseIf Trim(Rs_Consulta_Factura_Clientes!cancelada) = "S" Then
                    Btn_Cancelar.Enabled = False
                    Lbl_Facturacion.Caption = "CANCELADA"
                    Btn_Enviar_Email.Enabled = False
                ElseIf Trim(Rs_Consulta_Factura_Clientes!cancelada) = "PC" Then
                    Btn_Cancelar.Enabled = True
                    Lbl_Facturacion.Caption = "CANCELACIÓN EN PROCESO"
                    Btn_Enviar_Email.Enabled = False
                End If
                Btn_Imprimir.Caption = "Reimprimir"
            End If
            Btn_Imprimir.Enabled = True
            
            If Opt_Factura = True Or Opt_Nota_Cargo = True Then ' SI LA OPCION ES FACTURA
                'Prepara el recordset para consultar el número de factura de la tabla Adm_Descripcion_Facturas
                Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas_Detalles WHERE No_Factura ='" & Rs_Consulta_Factura_Clientes!No_Factura & "'"
                Set Rs_Consulta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Grid_Detalle_Factura.Rows = 0
                    Grid_Detalle_Factura.Cols = 13
                    'Pone el encabezado en las columnas
'                    Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) _
'                        & "Importe" & Chr(9) & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) _
'                        & "No_Salida" & Chr(9) & "IVA" & Chr(9) & "Aplica_IVA" & Chr(9) & "Unidad" & Chr(9) & "Incluye"
                        
                    Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Unidad" & Chr(9) & "Descripción" & Chr(9) & "Descripción SAT" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) _
                        & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) & "" & Chr(9) & "IVA" & Chr(9) _
                        & "Aplica_IVA" & Chr(9) & "Incluir"
    '            Grid_Detalle_Factura.Cols = 7
    '            Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) & "Producto"
                'Llenado del grid de la factura consultada
                While Not Rs_Consulta_Descripcion_Facturas.EOF
                    If Not IsNull(Rs_Consulta_Descripcion_Facturas!Unidad) Then
                        Unidad = Trim(Rs_Consulta_Descripcion_Facturas!Unidad)
                    Else
                        Unidad = ""
                    End If
                    Grid_Detalle_Factura.AddItem Rs_Consulta_Descripcion_Facturas!Cantidad & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Unidad & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Descripcion & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Clave_SAT & _
                        Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Precio, "###,##0.00") & _
                        Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Importe, "###,##0.00") & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Producto_ID & Chr(9) & "" & Chr(9) & "" & _
                        Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                    Grid_Detalle_Factura.FixedRows = 1
                    Rs_Consulta_Descripcion_Facturas.MoveNext
                Wend
                Formatea_Columnas_Grid
            Else 'SI LA OPCION ES REMISION
                Mi_SQL = "SELECT * FROM Adm_Clientes_Remisiones_Detalles WHERE No_Remision = '" & Rs_Consulta_Factura_Clientes!No_Remision & "'"
                Set Rs_Consulta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Grid_Detalle_Factura.Rows = 0
                Grid_Detalle_Factura.Cols = 7
                Grid_Detalle_Factura.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) & "Producto"
                'Llenado del grid de la remision consultada
                While Not Rs_Consulta_Descripcion_Facturas.EOF
                    Grid_Detalle_Factura.AddItem Rs_Consulta_Descripcion_Facturas!Cantidad & Chr(9) & Rs_Consulta_Descripcion_Facturas!Descripcion & Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Precio, "###,##0.00") & Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Importe, "###,##0.00") & Chr(9) & Rs_Consulta_Descripcion_Facturas!Producto_ID
                    Grid_Detalle_Factura.FixedRows = 1
                    Rs_Consulta_Descripcion_Facturas.MoveNext
                Wend
                'Configura el grid
                Grid_Detalle_Factura.ColWidth(0) = 800
                Grid_Detalle_Factura.ColWidth(1) = 6600
                Grid_Detalle_Factura.ColAlignment(1) = 1
                Grid_Detalle_Factura.ColWidth(2) = 1200
                Grid_Detalle_Factura.ColWidth(3) = 1200
                Grid_Detalle_Factura.ColWidth(4) = 0
                Grid_Detalle_Factura.ColWidth(5) = 0
                Grid_Detalle_Factura.ColWidth(6) = 0
            End If
            
            
            Mi_SQL = "SELECT * FROM Ope_Relacionados WHERE No_Factura_Electronica ='" & Format(No_Factura, "0000000000") & "'"
            Set RS_Consulta_Relacionados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not RS_Consulta_Relacionados.EOF Then
                Chc_Relacionados.Value = 1
                Grid_Relacionados.Rows = 0
                Grid_Relacionados.AddItem "Serie" & Chr(9) & "Folio" & Chr(9) & "UUID"
                Grid_Relacionados.ColWidth(0) = 500        'Serie
                Grid_Relacionados.ColWidth(1) = 500        'Folio
                Grid_Relacionados.ColWidth(2) = 2000        'UUID
                Call Conectar_Ayudante.Asigna_Item_Combo(RS_Consulta_Relacionados!Tipo_Relacion, Cmb_Relacionados)
                While Not RS_Consulta_Relacionados.EOF
                    Grid_Relacionados.AddItem RS_Consulta_Relacionados!Serie_Rel & Chr(9) & Val(RS_Consulta_Relacionados!Factura_Rel) & Chr(9) & RS_Consulta_Relacionados!UUID_Relacion
                    RS_Consulta_Relacionados.MoveNext
                Wend
                Grid_Relacionados.Enabled = True
            End If
                    
                    
            'Hace el recorrido de los datos del grid para hacer la suma
    '''        For Cont_Detalles = 1 To Grid_Detalle_Factura.Rows - 1
    '''            Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 3), ","))
    '''            Suma_IVA = Suma_IVA + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 6), ","))
    '''        Next Cont_Detalles
            
    '''        'Asigna los resultados a los totales
    '''        Txt_Subtotal.Text = Format(Suma, "#,##0.00")
    '''        Txt_IVA.Text = Format(Val(Suma) * Val(PG_Retencion_IVA), "#,##0.00")
    '''        Txt_Total.Text = Format(Val(Suma) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ",")), "#,##0.00")
            
            'Coloca la información consultada en los controles correspondientes
            Txt_Subtotal.text = Format(Rs_Consulta_Factura_Clientes!SubTotal, "#,##0.00")
            Txt_IVA.text = Format(Rs_Consulta_Factura_Clientes!Iva, "#,##0.00")
            Txt_Total.text = Format(Rs_Consulta_Factura_Clientes!Total, "#,##0.00")
            
            Rs_Consulta_Factura_Clientes.Close
            Rs_Consulta_Descripcion_Facturas.Close
            Fra_Datos_Cliente.Enabled = False
            Fra_Datos_Factura.Enabled = False
            Fra_Detalle_Factura.Enabled = True
            Btn_Agregar.Enabled = False
            Btn_Eliminar.Enabled = False
        Else
            MsgBox "Documento inexistente", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    End If
End Sub

Private Sub Txt_Precio_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Precio.text, True)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Actualizar_Refacturacion
'DESCRIPCIÓN            : Actualiza los datos de la factura al refacturar
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 18- Nov - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************'
Public Sub Actualizar_Refacturacion()
Dim Rs_Actualiza_Factura_Clientes As rdoResultset                       'Manejo del registro de Adm_Factura_Clientes
Dim Rs_Actualiza_Descripcion_Facturas As rdoResultset                   'Manejo del registro de Adm_Descripcion_Facturas
Dim Mi_SQL As String
Dim Rs_Elimina_Detalles As rdoResultset                           'Manejo del registro de Adm_Remision_Clientes


Dim Rs_Alta_Remision_Clientes As rdoResultset                           'Manejo del registro de Adm_Remision_Clientes
Dim Rs_Alta_Descripcion_Remision As rdoResultset                        'Manejo del registro de Adm_Descripcion_Remision
Dim Rs_Modifica_Alm_Salidas_Almacen_Detalles As rdoResultset            'Manejo del registro de Adm_Descripcion_Remision
Dim Cont_Detalles_Factura As Integer                                    'Contador para agregar los datos del grid en la base de datos

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Actualiza de Factura
    Mi_SQL = " SELECT * FROM Adm_Clientes_Facturas WHERE No_Factura='" & Format(Txt_No_Factura.text, "0000000000") & "'"
    Set Rs_Actualiza_Factura_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Adm_Factura_Clientes con los datos contenidos en las cajas de textos
    With Rs_Actualiza_Factura_Clientes
        .Edit
            .rdoColumns("Cliente_ID") = Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000")
            .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Pago") = Format(DTP_Fecha_Pago.Value, "MM/dd/yyyy")
            .rdoColumns("Tipo_Pago") = Cmb_Tipo_Pago.text
            .rdoColumns("Tipo_Factura") = Cmb_Tipo_Factura.text
            .rdoColumns("Subtotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
            .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
            .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
            .rdoColumns("Cancelada") = "N"
            .rdoColumns("Comentarios") = Txt_Comentarios.text
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now()
            'Valida si la factura se va a pagar a crédito para asignar los valores
            If Cmb_Tipo_Pago.text = "CREDITO" Then
                .rdoColumns("Abono") = 0
                .rdoColumns("Saldo") = CDbl(Txt_Total.text)
                .rdoColumns("Pagada") = "N"
            Else
                .rdoColumns("Abono") = CDbl(Txt_Total.text)
                .rdoColumns("Saldo") = 0
                .rdoColumns("Pagada") = "S"
            End If
        .Update
    End With
    
    'SE MODIFICA LA SALIDA PARA PODERACTUALIZAR
    Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Trim(Cmb_Salidas.text) & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4), "00000") & "'"
    Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
        With Rs_Modifica_Alm_Salidas_Almacen_Detalles
            .Edit
                .rdoColumns("Facturado") = "NO"
            .Update
        End With
    End If
    Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
    
    'SE ELIMINAN LOS DETALLES PARA PODER ACTUALIZARLAS
    Mi_SQL = " SELECT * FROM Adm_Clientes_Facturas_Detalles WHERE No_Factura='" & Trim(Txt_No_Factura.text) & "'"
    Set Rs_Elimina_Detalles = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
    While Not Rs_Elimina_Detalles.EOF
        With Rs_Elimina_Detalles
            .Delete
        End With
        Rs_Elimina_Detalles.MoveNext
    Wend
    Rs_Elimina_Detalles.Close
    
    Set Rs_Actualiza_Descripcion_Facturas = Conectar_Ayudante.Recordset_Agregar("Adm_Clientes_Facturas_Detalles")
    For Cont_Detalles_Factura = 1 To Grid_Detalle_Factura.Rows - 1
        'Llena la tabla de Adm_Clientes_Facturas_Detalles con los datos contenidos en el grid
        With Rs_Actualiza_Descripcion_Facturas
            .AddNew
                .rdoColumns("No_Factura") = Txt_No_Factura.text
                If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4) <> "" Then
                    .rdoColumns("Producto_ID") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4)
                Else
                    .rdoColumns("Producto_ID") = Null
                End If
                .rdoColumns("Descripcion") = Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 1)
                .rdoColumns("Cantidad") = Val(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 0))
                .rdoColumns("Precio") = Val(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2))
                .rdoColumns("Importe") = Val(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 3))
            .Update
        End With

        'MODIFICA LA SALIDA
        Mi_SQL = "SELECT * FROM Alm_Salidas_Almacen_Detalles WHERE No_Salida ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 5), "0000000000") & "' and Producto_ID ='" & Format(Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 4), "00000") & "'"
        Set Rs_Modifica_Alm_Salidas_Almacen_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Modifica_Alm_Salidas_Almacen_Detalles.EOF Then
            With Rs_Modifica_Alm_Salidas_Almacen_Detalles
                .Edit
                    .rdoColumns("Facturado") = "SI"
                .Update
            End With
        End If
        Rs_Modifica_Alm_Salidas_Almacen_Detalles.Close
    Next Cont_Detalles_Factura
    
    'Cierra los manejadores del registro
    Rs_Actualiza_Factura_Clientes.Close
    Rs_Actualiza_Descripcion_Facturas.Close
    Conexion_Base.CommitTrans
    If MsgBox("La factura ha sido actualizada" & Chr(13) & " ¿Desea enviarla a Imprimir?", vbYesNo + vbInformation) = vbYes Then
        Imprimir_Facturas
    End If
    'Deshabilita controles y habilita los necesarios
    Fra_Datos_Cliente.Enabled = False
    Fra_Datos_Factura.Enabled = False
    Fra_Detalle_Factura.Enabled = False
    Fra_Comentarios.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
     Btn_Nuevo.Enabled = True
    Btn_Imprimir.Enabled = True
    Btn_Cancelar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Refacturar.Caption = "Refacturar"
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Txt_Presio_Sin_IVA_Change()
With Grid_Detalle_Factura
        If .RowSel > 0 Then
            If .ColSel = 4 Then 'PRECIO VENTA E IMPORTE
                If Val(Txt_Presio_Sin_IVA.text) >= 0 Then
                    .TextMatrix(.RowSel, 4) = Txt_Presio_Sin_IVA.text
                    .TextMatrix(.RowSel, 5) = Val(.TextMatrix(.RowSel, 4)) * Val(.TextMatrix(.RowSel, 0))
                    If Trim(.TextMatrix(.RowSel, 11)) = "SI" Then
                        .TextMatrix(.RowSel, 8) = PG_Retencion_IVA 'IMPUESTO
                        .TextMatrix(.RowSel, 10) = (Val(.TextMatrix(.RowSel, 4)) * Val(.TextMatrix(.RowSel, 0))) * Val(PG_Retencion_IVA) 'IVA
                    End If
                    'Hace el recorrido de los datos del grid para hacer la suma
                    For Cont_Detalles = 1 To Grid_Detalle_Factura.Rows - 1
                        Suma = Val(Suma) + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 5), ","))
                        Suma_IVA = Val(Suma_IVA) + Val(Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Factura.TextMatrix(Cont_Detalles, 10), ",")), "#0.00"))
                    Next Cont_Detalles
                    Txt_Subtotal.text = ""
                    Txt_IVA.text = ""
                    Txt_Total.text = ""
                    'Asigna los resultados a los totales
                    Txt_Subtotal.text = Format(Suma, "#,##0.00")
                    Txt_IVA.text = Format(Val(Suma_IVA), "#,##0.00")
                    Txt_Total.text = Format(Val(Suma) + Val(Suma_IVA), "#,##0.00")
                End If
            End If
            If .ColSel = 2 Then 'DESCRIPCION
                .TextMatrix(.RowSel, 2) = Txt_Presio_Sin_IVA.text
            End If
        End If
    End With
End Sub

Private Sub Txt_Presio_Sin_IVA_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
        If KeyCode > 37 Then Grid_Detalle_Factura.SetFocus
        If KeyCode = 37 Then
            If Txt_Presio_Sin_IVA.SelStart = 0 Then
                Grid_Detalle_Factura.SetFocus
                ''Grid_Detalle_Factura.Col = Grid_Detalle_Factura.ColSel - 1
            End If
        End If
        If Grid_Detalle_Factura.Row > 1 Then
            If KeyCode = 38 Then Grid_Detalle_Factura.Row = Grid_Detalle_Factura.RowSel - 1
            If KeyCode = 40 Then
                If Grid_Detalle_Factura.Row < Grid_Detalle_Factura.Rows - 1 Or Grid_Detalle_Factura.Row = 1 Then
                     Grid_Detalle_Factura.Row = Grid_Detalle_Factura.RowSel + 1
                Else
                    Txt_Presio_Sin_IVA.Visible = False
                    Exit Sub
                End If
            End If
            If Grid_Detalle_Factura.Col = 9 Then
                If KeyCode = 39 Then Grid_Detalle_Factura.Col = Grid_Detalle_Factura.ColSel + 1
            End If
        Else
            If KeyCode = 40 Then
                If Grid_Detalle_Factura.Row < Grid_Detalle_Factura.Rows - 1 Or Grid_Detalle_Factura.Row <> 1 Then
                    Grid_Detalle_Factura.Row = Grid_Detalle_Factura.RowSel + 1
                Else
                    Txt_Presio_Sin_IVA.Visible = False
                    Exit Sub
                End If
            End If
            If Grid_Detalle_Factura.Col = 2 Then
                If KeyCode = 39 Then Grid_Detalle_Factura.Col = Grid_Detalle_Factura.ColSel + 1
            End If
        End If
        If Txt_Presio_Sin_IVA.Visible = True Then
            Txt_Presio_Sin_IVA.SetFocus
            'SendKeys "{Home}+{End}"
        End If
    End If
    If KeyCode = 13 Then
        Txt_Presio_Sin_IVA.Visible = False
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprimir_Remision
'DESCRIPCIÓN            : Imprime la Remision o vale de salida
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 06-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Imprimir_Remision()
Dim Mi_SQL As String                            'Cadena para general las consultas
Dim Rs_Formato  As rdoResultset                 'Manejo de registro para la tabla Cfg_Formatos
Dim Rs_Formato_Generales  As rdoResultset       'Manejo de registro para la tabla Cfg_Formatos_Generales
Dim Rs_Formato_Detalles As rdoResultset         'Manejo de registro para la tabla Cfg_Formatos_Detalles
Dim Rs_Generales_Salida As rdoResultset
Dim Rs_Detalles_Salida As rdoResultset
Dim Clave As String
Dim Cantidad As Integer
Dim ARTICULO As String
Dim Nombre As String
Dim DOMICILIO As String
Dim Ciudad As String
Dim Fecha As String
Dim CoordenadaX As Double
Dim CoordenadaY As Double
Dim Rs_Consulta_Cliente As rdoResultset
        
    On Error GoTo handler
    
    'Consulta para la configuración de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Set Rs_Formato = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta para la configuración general de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'General'"
    Set Rs_Formato_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta para la configuración a detalle de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'Detalle' ORDER BY Campo"
    Set Rs_Formato_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Impresión de la factura
    If Not Rs_Formato.EOF Then
    
        'Crea el encabezado de la remisión
        Printer.ScaleMode = vbCentimeters
        Printer.FontSize = 16
        Printer.Font = "Arial"
        Printer.FontBold = True
        Printer.CurrentX = 5
        Printer.CurrentY = 0.5
        Printer.Print Nombre_Emisor
        Printer.CurrentX = 18
        Printer.CurrentY = 0.5
        Printer.Print "REMISION"
        Printer.CurrentX = 18.5
        Printer.CurrentY = 1.5
        Printer.FontSize = 11
        Printer.Print "No. " & Val(Cmb_Salidas.text)
        Printer.CurrentX = 3.8
        Printer.CurrentY = 1.5
        Printer.FontSize = 8
        Printer.Print Calle_Emisor & " No. " & No_Exterior_Emisor & "    COL. " & Colonia_Emisor & "  CP " & Codigo_Postal_Emisor & "  " & Municipio_Emisor & ",GTO."
        Printer.CurrentX = 6.5
        Printer.CurrentY = 2
        Printer.Print "TEL. y FAX  01(462)633-20-32      633-20-33 y 633-20-34"
        Printer.CurrentX = 6.8
        Printer.CurrentY = 2.5
        Printer.Print "alcesa@prodigy.net.mx            www.alcesa.com.mx"
        
        'Configura la fuente de la factura para generales
        With Rs_Formato
            Printer.ScaleMode = vbCentimeters
            Printer.FontSize = .rdoColumns("Tamaño_Generales")
            Printer.Font = .rdoColumns("Letra_Generales")
            If .rdoColumns("Estilo_Generales") = "Negrita" Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
        End With
        'SE CONSULTAN LOS DATOS GENERALES DE LA SALIDA
        Mi_SQL = " SELECT Alm_Salidas_Almacen.*,Cat_Clientes.*"
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen, Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen.No_Salida = '" & Format(Cmb_Salidas.text, "0000000000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Cliente_ID = Cat_Clientes.Cliente_ID"
        Set Rs_Generales_Salida = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Imprime los datos del cliente
        With Rs_Formato_Generales
            While Not .EOF
                    Printer.CurrentX = .rdoColumns("X")
                    Printer.CurrentY = .rdoColumns("Y")
                    Longitud = .rdoColumns("Longitud")
                    If .rdoColumns("Campo") = "NOMBRE" Then
                       Printer.Print Mid(Rs_Generales_Salida!Nombre, 1, Longitud)
                       CoordenadaY = .rdoColumns("Y")
                    End If
                    If .rdoColumns("Campo") = "DOMICILIO" Then
                       Printer.Print Mid(Rs_Generales_Salida.rdoColumns("Direccion_Remision") & " " & Rs_Generales_Salida.rdoColumns("Colonia_Remision"), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "CIUDAD" Then
                       Printer.Print Mid(Rs_Generales_Salida.rdoColumns("Ciudad_Remision") & "," & Rs_Generales_Salida.rdoColumns("Estado_Remision"), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "FECHA" Then
                       Printer.Print Format(Rs_Generales_Salida.rdoColumns("Fecha_Salida"), "dd/MM/yyyy")
                    End If
                    If .rdoColumns("Campo") = "NOMBRE_ALMACEN" Then
                       Printer.Print Mid(UCase(Txt_Nombre_Almacenista.text), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "NOMBRE_RECIBE" Then
                       Printer.Print Mid(UCase(Txt_Recibe.text), 1, Longitud)
                    End If
                .MoveNext
            Wend
        End With
        Rs_Generales_Salida.Close
        'Configura la fuente de la factura para detalles
        With Rs_Formato
            Printer.FontSize = .rdoColumns("Tamaño_Detalles")
            Printer.Font = .rdoColumns("Letra_Detalles")
            If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
        End With
        'Consulta de la tabla Adm_Descripcion_Facturas con el número de facturas
        Mi_SQL = "SELECT Alm_Salidas_Almacen_Detalles.*,Cat_Productos.Clave FROM Cat_Productos, Alm_Salidas_Almacen_Detalles"
        Mi_SQL = Mi_SQL & " WHERE No_Salida = '" & Format(Cmb_Salidas.text, "0000000000") & "'"
        Mi_SQL = Mi_SQL & " AND  Cat_Productos.Producto_ID = Alm_Salidas_Almacen_Detalles.Producto_ID "
        Set Rs_Detalles_Salida = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Configura la fuente para empresión de los detalles
        If Not Rs_Formato.EOF Then
            With Rs_Formato
                Printer.ScaleMode = vbCentimeters
                Printer.FontSize = .rdoColumns("Tamaño_Detalles")
                Printer.Font = .rdoColumns("Letra_Detalles")
                If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
                Salto = .rdoColumns("Separacion_Detalles")
            End With
            Cont_Renglon = 0
            Mi_SQL = " SELECT * FROM Cat_Clientes WHERE Cliente_ID='" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000") & "'"
            Set Rs_Consulta_Cliente = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            'Imprime los detalles
            While Not Rs_Detalles_Salida.EOF
                Cont_Renglon = Cont_Renglon + Salto
                While Not Rs_Formato_Detalles.EOF
                    Printer.CurrentX = Rs_Formato_Detalles.rdoColumns("X")
                    Printer.CurrentY = Rs_Formato_Detalles.rdoColumns("Y") + Cont_Renglon
                    Longitud = Rs_Formato_Detalles.rdoColumns("Longitud")
                    If Rs_Formato_Detalles.rdoColumns("Campo") = "CLAVE" Then
                        Printer.Print Trim(Rs_Detalles_Salida.rdoColumns("Clave"))
                        If CoordenadaX = 0 Then CoordenadaX = Rs_Formato_Detalles.rdoColumns("X")
                    End If
                    If Trim(Rs_Formato_Detalles.rdoColumns("Campo")) = "CANTIDAD" Then
                        Printer.Print Conectar_Ayudante.Alinea_Derecha(Trim(Rs_Detalles_Salida.rdoColumns("Cantidad")), 7)
                    End If
                    If Rs_Formato_Detalles.rdoColumns("Campo") = "ARTICULO" Then
                        Printer.Print Trim(Rs_Detalles_Salida.rdoColumns("Descripcion"))
                    End If
                    'REVISA SI LA REMISION SE VA A IMPRIMIR CON PRECIOS
                    If Not IsNull(Rs_Consulta_Cliente!Remision_Con_Presio) Then
                        If Trim(Rs_Consulta_Cliente!Remision_Con_Presio) = "SI" Then
                            If Trim(Rs_Formato_Detalles.rdoColumns("Campo")) = "PRECIO" Then
                                Printer.Print "$" & Conectar_Ayudante.Alinea_Derecha(Trim(Rs_Detalles_Salida.rdoColumns("Precio_Venta")), 7)
                            End If
                            If Trim(Rs_Formato_Detalles.rdoColumns("Campo")) = "IMPORTE" Then
                                Printer.Print "$" & Conectar_Ayudante.Alinea_Derecha(Trim(Rs_Detalles_Salida.rdoColumns("Importe")), 7)
                            End If
                            If Trim(Rs_Formato_Detalles.rdoColumns("Campo")) = "TOTAL" Then
                                Printer.Print "TOTAL" & Chr(9) & "$ " & Conectar_Ayudante.Alinea_Derecha(Trim(Rs_Detalles_Salida.rdoColumns("Total")), 7)
                            End If
                        End If
                    End If
                    Rs_Formato_Detalles.MoveNext
                Wend
                Rs_Formato_Detalles.MoveFirst
                Rs_Detalles_Salida.MoveNext
            Wend
            Rs_Detalles_Salida.Close
        End If
        Printer.PaintPicture Pic_Logotipo.Picture, 0.5, 0.5
        Printer.EndDoc
    End If
    Rs_Formato.Close
    Rs_Formato_Generales.Close
    Rs_Formato_Detalles.Close
    Rs_Consulta_Cliente.Close
'    MsgBox "Remisión enviada a Impresión", vbInformation
    Exit Sub
handler:
    MsgBox Err.Description
    Printer.EndDoc
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Btn_Enviar_Email_Click
'DESCRIPCIÓN: Realiza el proceso de envío de la factura electrónica por correo
'PARÁMETROS:
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  17-Agosto-2012
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Btn_Enviar_Email_Click()
Dim Archivo_XML As String       'Almacena el nombre del archivo xml
Dim Archivo_PDF As String       'Almacena el nombre del archivo pdf
Dim Archivos_Adjuntos As String
Dim Correo_Para As String       'Almacena la dirección del correo destino
Dim Correo_De As String         'Almacena la dirección de correo origen
Dim Enviar As Boolean           'Bandera que se habilita si se encuentran los archivos a adjuntar en el correo
Dim Mensaje As String           'Cadena que será el cuerpo del correo
Dim Copias As String

    If Txt_No_Factura.text <> "" Then
        'Por default se deshabilita la bandera
        Enviar = False
        Archivos_Adjuntos = ""
        'Valida que existan los archivos a adjuntar
        'Asgna a las variables los nombres de los archivos
        Archivo_PDF = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text) & ".pdf"
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Pdfs & "\" & Archivo_PDF, "ARCHIVO") = True Then
            Archivo_XML = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Factura.text) & ".xml"
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Xmls & "\" & Archivo_XML, "ARCHIVO") = True Then
                'Si se encuentran ambos archivos, habilita la bandera para indicar que se puede continuar con el proceso
                Enviar = True
            Else
                'No se encuentro el archivo XMl, manda mensaje indicando la situación y deshabilita la bandera para no continuar con el proceso
                MsgBox "No se encontró el archivo XML, favor de verificar", vbExclamation
                Enviar = False
            End If
        Else
            'No se encuentro el archivo XMl, manda mensaje indicando la situación y deshabilita la bandera para no continuar con el proceso
            MsgBox "No se encontró el archivo PDF, favor de verificar", vbExclamation
            Enviar = False
        End If

        If Enviar = True Then
            MDIFrm_Apl_Principal.MousePointer = 11
            'Revisa que se tenga la direccione de correo destino
            If Txt_Email.text = "" Then
                'Si no se tiene la dirección del cliente, la pregunta
                Correo_Para = InputBox("Ingrese la dirección de correo del cliente para enviar el email")
                If Correo_Para = "" Then
                    MDIFrm_Apl_Principal.MousePointer = 0
                    MsgBox "Ingrese una dirección de correo válida", vbExclamation
                    Exit Sub
                Else
                    If Conectar_Ayudante.Validar_Email(Trim(Correo_Para)) = True Then
                        'asigna la dirección al text
                        Txt_Email.text = Correo_Para
                    Else
                        MsgBox "La dirección ingresada no es válida, favor de verificar", vbExclamation
                        Exit Sub
                    End If
                End If
            End If

            'Si no hay pregunta si desea enviar el correo sin copias
            If MsgBox("¿Desea realizar el envío con alguna copia del correo?", vbQuestion + vbYesNo) = vbYes Then
                'Si no se tiene la dirección del cliente, la pregunta
                Copias = InputBox("Ingrese la(s) dirección(es) de correo para copia del email")
                If Copias = "" Then
                    MDIFrm_Apl_Principal.MousePointer = 0
                    MsgBox "Ingrese una dirección de correo válida", vbExclamation
                    Exit Sub
                End If
            End If
                
            Frm_Apl_Enviando_Correo.Show
            'Realiza el envío del correo
            Archivos_Adjuntos = Ruta_Pdfs & "\" & Archivo_PDF & "|" & Ruta_Xmls & "\" & Archivo_XML
            
            Mensaje = "<HTML>" & vbNewLine & _
                "<BODY>" & vbNewLine & _
                    "<P>Estimado cliente: <BR></P>" & vbNewLine & _
                    "<P>Por medio de este correo, se le hace entrega de un Comprobante Fiscal Digital por Internet(CFDI).<BR></P>" & vbNewLine & _
                    "<P>Adjunto a este correo encontrará un archivo PDF y un archivo XML correspondientes a su factura.</P>" & vbNewLine & _
                    "<P>Atentamente: <BR></P>" & vbNewLine & _
                    "<P>" & Nombre_Emisor & "<BR></P>" & vbNewLine & vbNewLine & _
                "</BODY>" & vbNewLine & _
                "</HTML>"
                
            If Enviar_Correo_Documentos("", Nombre_Emisor, Txt_Email.text, "Facturación Electrónica ALCESA", Mensaje, True, True, Archivos_Adjuntos, Copias) = True Then
                Unload Frm_Apl_Enviando_Correo
                MDIFrm_Apl_Principal.MousePointer = 0
                MsgBox "Correo enviado satisfactoriamente", vbInformation
            Else
                MDIFrm_Apl_Principal.MousePointer = 0
                Unload Frm_Apl_Enviando_Correo
            End If
        End If
    Else
        MsgBox "Seleccione una factura para poder hacer el envío por correo", vbExclamation
    End If
End Sub


