VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_Cat_Clientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "CATÁLOGO DE CLIENTES"
   ClientHeight    =   7380
   ClientLeft      =   5460
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   Tag             =   "Resize=TLH"
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
      Left            =   8175
      Picture         =   "Frm_Cat_Clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   65
      Tag             =   "A"
      Top             =   6660
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
      Left            =   6150
      Picture         =   "Frm_Cat_Clientes.frx":36FF
      Style           =   1  'Graphical
      TabIndex        =   64
      Tag             =   "B"
      Top             =   6660
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
      Left            =   4140
      Picture         =   "Frm_Cat_Clientes.frx":6CB9
      Style           =   1  'Graphical
      TabIndex        =   63
      Tag             =   "C"
      Top             =   6660
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
      Left            =   2130
      Picture         =   "Frm_Cat_Clientes.frx":A245
      Style           =   1  'Graphical
      TabIndex        =   62
      Tag             =   "M"
      Top             =   6660
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
      Picture         =   "Frm_Cat_Clientes.frx":D976
      Style           =   1  'Graphical
      TabIndex        =   61
      Tag             =   "A"
      Top             =   6660
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Alta_Toma_Inventario 
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
      Height          =   675
      Left            =   3990
      TabIndex        =   113
      Top             =   8100
      Width           =   1065
   End
   Begin VB.CommandButton Btn_Loyaut_Toma_Inventario 
      Caption         =   "Toma de Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1785
      TabIndex        =   112
      Top             =   8115
      Width           =   2190
   End
   Begin VB.CommandButton Btn_Layout_Proveedores 
      Caption         =   "Layout Proveedores"
      Height          =   375
      Left            =   6600
      TabIndex        =   111
      Top             =   8445
      Width           =   1935
   End
   Begin VB.CommandButton Btn_Alta_Loyaut_Proveedores 
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
      Height          =   375
      Left            =   8520
      TabIndex        =   110
      Top             =   8445
      Width           =   465
   End
   Begin VB.CommandButton Btn_Alta_Layout_Clientes 
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
      Height          =   375
      Left            =   8520
      TabIndex        =   109
      Top             =   8085
      Width           =   465
   End
   Begin VB.CommandButton Btn_Layout_Carga_Clientes 
      Caption         =   "Layout Clientes"
      Height          =   375
      Left            =   6600
      TabIndex        =   108
      Top             =   8085
      Width           =   1935
   End
   Begin VB.CommandButton Btn_Dar_Alta_Layout 
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
      Left            =   7560
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Btn_Cargar_Layout 
      Caption         =   "Layout"
      Height          =   330
      Left            =   7440
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Btn_Cargar_Presentaciones 
      Caption         =   "Cargar Productos"
      Height          =   330
      Left            =   6600
      TabIndex        =   101
      Top             =   7725
      Width           =   1935
   End
   Begin VB.Data Dt_Excel 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7005
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   8520
      TabIndex        =   100
      Top             =   7725
      Width           =   465
   End
   Begin MSComDlg.CommonDialog Cdg_Exel 
      Left            =   8505
      Top             =   6930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_Clientes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6630
      Left            =   0
      ScaleHeight     =   6630
      ScaleWidth      =   15540
      TabIndex        =   58
      Top             =   0
      Width           =   15540
      Begin VB.Frame Fra_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   165
         TabIndex        =   59
         Top             =   3600
         Width           =   9420
         Begin MSFlexGridLib.MSFlexGrid Grid_Clientes 
            Height          =   2670
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   4710
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin TabDlg.SSTab Tab_Clientes 
         Height          =   2950
         Left            =   165
         TabIndex        =   126
         Top             =   600
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   5212
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "Frm_Cat_Clientes.frx":10EAD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Datos_Generales_clientes"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos de Facturación"
         TabPicture(1)   =   "Frm_Cat_Clientes.frx":10EC9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Datos_Factura"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Datos de Remisión"
         TabPicture(2)   =   "Frm_Cat_Clientes.frx":10EE5
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Fra_Remisiones"
         Tab(2).ControlCount=   1
         Begin VB.Frame Fra_Remisiones 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   2535
            Left            =   -74880
            TabIndex        =   153
            Top             =   340
            Width           =   9165
            Begin VB.TextBox Txt_CP_Remision 
               Height          =   285
               Left            =   5925
               TabIndex        =   26
               Top             =   527
               Width           =   3060
            End
            Begin VB.TextBox Txt_Ciudad_Remision 
               Height          =   285
               Left            =   1125
               TabIndex        =   27
               Top             =   855
               Width           =   3060
            End
            Begin VB.TextBox Txt_Direccion_Remision 
               Height          =   285
               Left            =   1125
               TabIndex        =   24
               Top             =   200
               Width           =   7860
            End
            Begin VB.TextBox Txt_Estado_Remision 
               Height          =   285
               Left            =   5925
               TabIndex        =   28
               Top             =   855
               Width           =   3060
            End
            Begin VB.TextBox Txt_Colonia_Remision 
               Height          =   285
               Left            =   1125
               TabIndex        =   25
               Top             =   527
               Width           =   3060
            End
            Begin VB.Label Label25 
               BackColor       =   &H8000000E&
               Caption         =   "Estado"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5190
               TabIndex        =   158
               Top             =   870
               Width           =   675
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CP"
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
               Left            =   5190
               TabIndex        =   157
               Top             =   555
               Width           =   225
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ciudad"
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
               Left            =   150
               TabIndex        =   156
               Top             =   885
               Width           =   510
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dirección"
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
               Left            =   150
               TabIndex        =   155
               Top             =   225
               Width           =   690
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Colonia"
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
               Left            =   150
               TabIndex        =   154
               Top             =   585
               Width           =   540
            End
         End
         Begin VB.Frame Fra_Datos_Generales_clientes 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   2535
            Left            =   120
            TabIndex        =   138
            Top             =   340
            Width           =   9165
            Begin VB.TextBox Txt_Nombre_Cliente 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1365
               TabIndex        =   3
               Top             =   450
               Width           =   7620
            End
            Begin VB.TextBox Txt_Telefono_Cliente 
               Height          =   285
               Left            =   1365
               TabIndex        =   4
               Top             =   765
               Width           =   1635
            End
            Begin VB.TextBox Txt_Celular_Cliente 
               Height          =   285
               Left            =   3765
               TabIndex        =   5
               Top             =   765
               Width           =   1635
            End
            Begin VB.TextBox Txt_Almacen 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1365
               TabIndex        =   11
               Top             =   1755
               Width           =   4035
            End
            Begin VB.TextBox Txt_Fax_Cliente 
               Height          =   285
               Left            =   6885
               TabIndex        =   6
               Top             =   765
               Width           =   2115
            End
            Begin VB.TextBox Txt_Comentarios_Cliente 
               Height          =   405
               Left            =   1365
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   2050
               Width           =   7620
            End
            Begin VB.TextBox Txt_E_Mail_Cliente 
               Height          =   285
               Left            =   1365
               TabIndex        =   7
               Top             =   1095
               Width           =   4035
            End
            Begin VB.TextBox Txt_Dias_Credito 
               Height          =   285
               Left            =   6885
               MaxLength       =   3
               TabIndex        =   8
               Top             =   1095
               Width           =   2115
            End
            Begin VB.ComboBox Cmb_Clasificacion_Clientes 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F01
               Left            =   1365
               List            =   "Frm_Cat_Clientes.frx":10F03
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1395
               Width           =   4035
            End
            Begin VB.ComboBox Cmb_Credoto_Flexible 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F05
               Left            =   6885
               List            =   "Frm_Cat_Clientes.frx":10F0F
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   1395
               Width           =   2115
            End
            Begin VB.ComboBox Cmb_Remision_Con_Precios 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F1B
               Left            =   6885
               List            =   "Frm_Cat_Clientes.frx":10F25
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   1740
               Width           =   2115
            End
            Begin VB.TextBox Txt_RFC 
               Height          =   285
               Left            =   6885
               TabIndex        =   2
               Top             =   120
               Width           =   2115
            End
            Begin VB.ComboBox Cmb_Status_Cliente 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F31
               Left            =   3765
               List            =   "Frm_Cat_Clientes.frx":10F3B
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   120
               Width           =   1635
            End
            Begin VB.TextBox Txt_Cliente_ID 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1365
               Locked          =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   135
               Width           =   1635
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Almacen Entrega"
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
               TabIndex        =   152
               Top             =   1785
               Width           =   1230
            End
            Begin VB.Label Lbl_Fax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
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
               Left            =   5445
               TabIndex        =   151
               Top             =   795
               Width           =   285
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
               Left            =   90
               TabIndex        =   150
               Top             =   2145
               Width           =   915
            End
            Begin VB.Label Lbl_Celular 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Celular"
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
               Left            =   3060
               TabIndex        =   149
               Top             =   795
               Width           =   510
            End
            Begin VB.Label Lbl_Telefono 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono"
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
               TabIndex        =   148
               Top             =   795
               Width           =   630
            End
            Begin VB.Label Lbl_E_Mail 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
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
               TabIndex        =   147
               Top             =   1125
               Width           =   450
            End
            Begin VB.Label Lbl_Dias_Credito 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Días Crédito"
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
               Left            =   5445
               TabIndex        =   146
               Top             =   1125
               Width           =   900
            End
            Begin VB.Label Lbl_Clasificacion_Clientes 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Clasificación "
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
               TabIndex        =   145
               Top             =   1440
               Width           =   960
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credito Flexible"
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
               Left            =   5445
               TabIndex        =   144
               Top             =   1440
               Width           =   1125
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Remisión c/precios"
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
               Left            =   5445
               TabIndex        =   143
               Top             =   1785
               Width           =   1380
            End
            Begin VB.Label Lbl_RFC 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RFC"
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
               Left            =   5445
               TabIndex        =   142
               Top             =   150
               Width           =   345
            End
            Begin VB.Label Lbl_Status_Cliente 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
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
               Left            =   3045
               TabIndex        =   141
               Top             =   165
               Width           =   660
            End
            Begin VB.Label Lbl_Nombre_Cliente 
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
               TabIndex        =   140
               Top             =   480
               Width           =   660
            End
            Begin VB.Label Lbl_Cliente_ID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente ID"
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
               TabIndex        =   139
               Top             =   165
               Width           =   855
            End
         End
         Begin VB.Frame Fra_Datos_Factura 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   2535
            Left            =   -74880
            TabIndex        =   127
            Top             =   360
            Width           =   9165
            Begin VB.ComboBox Cmb_Tipo_Persona 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F52
               Left            =   1365
               List            =   "Frm_Cat_Clientes.frx":10F5C
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   1870
               Width           =   4020
            End
            Begin VB.TextBox Txt_Cliente_Pais 
               Height          =   285
               Left            =   6555
               TabIndex        =   21
               Top             =   1142
               Width           =   2430
            End
            Begin VB.TextBox Txt_Cliente_No_Int 
               Height          =   285
               Left            =   3555
               TabIndex        =   16
               Top             =   521
               Width           =   1830
            End
            Begin VB.TextBox Txt_Cuenta_Pago 
               Height          =   315
               Left            =   6555
               TabIndex        =   23
               Top             =   1485
               Width           =   2430
            End
            Begin VB.ComboBox Cmb_Metodo_Pago 
               Height          =   315
               ItemData        =   "Frm_Cat_Clientes.frx":10F6F
               Left            =   1365
               List            =   "Frm_Cat_Clientes.frx":10F88
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   1485
               Width           =   4020
            End
            Begin VB.TextBox Txt_Cliente_No_Ext 
               Height          =   285
               Left            =   1365
               TabIndex        =   15
               Top             =   514
               Width           =   1830
            End
            Begin VB.TextBox Txt_Colonia_Cliente 
               Height          =   285
               Left            =   1365
               TabIndex        =   18
               Top             =   828
               Width           =   4020
            End
            Begin VB.TextBox Txt_Estado_Cliente 
               Height          =   285
               Left            =   1365
               TabIndex        =   20
               Top             =   1142
               Width           =   4020
            End
            Begin VB.TextBox Txt_Direccion_Cliente 
               Height          =   285
               Left            =   1365
               TabIndex        =   14
               Top             =   200
               Width           =   7620
            End
            Begin VB.TextBox Txt_Ciudad_Cliente 
               Height          =   285
               Left            =   6555
               TabIndex        =   19
               Top             =   828
               Width           =   2430
            End
            Begin VB.TextBox Txt_Codigo_Postal_Cliente 
               Height          =   285
               Left            =   6555
               TabIndex        =   17
               Top             =   521
               Width           =   2430
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Tipo Persona"
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
               TabIndex        =   161
               Top             =   1880
               Width           =   975
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Método de Pago"
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
               TabIndex        =   137
               Top             =   1530
               Width           =   1185
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "No Cta. Pago"
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
               Left            =   5535
               TabIndex        =   136
               Top             =   1530
               Width           =   975
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "País"
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
               Left            =   5535
               TabIndex        =   135
               Top             =   1170
               Width           =   315
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Int."
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
               Left            =   3240
               TabIndex        =   134
               Top             =   551
               Width           =   225
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. Ext."
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
               TabIndex        =   133
               Top             =   551
               Width           =   585
            End
            Begin VB.Label Lbl_Colonia_Cliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Colonia"
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
               TabIndex        =   132
               Top             =   858
               Width           =   540
            End
            Begin VB.Label Lbl_Direccion 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dirección"
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
               TabIndex        =   131
               Top             =   230
               Width           =   690
            End
            Begin VB.Label Lbl_Ciudad_Cliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ciudad"
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
               Left            =   5535
               TabIndex        =   130
               Top             =   855
               Width           =   510
            End
            Begin VB.Label Lbl_CP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CP"
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
               Left            =   5535
               TabIndex        =   129
               Top             =   555
               Width           =   225
            End
            Begin VB.Label Lbl_Estado_Cliente 
               BackColor       =   &H8000000E&
               Caption         =   "Estado"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   90
               TabIndex        =   128
               Top             =   1149
               Width           =   675
            End
         End
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   3270
         Picture         =   "Frm_Cat_Clientes.frx":10FF9
         Top             =   45
         Width           =   360
      End
      Begin VB.Label Lbl_CLIENTES 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTES"
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
         Left            =   3885
         TabIndex        =   60
         Top             =   -60
         Width           =   1905
      End
   End
   Begin VB.PictureBox Pic_Cat_Productos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6630
      Left            =   0
      ScaleHeight     =   6630
      ScaleWidth      =   13560
      TabIndex        =   85
      Top             =   0
      Width           =   13560
      Begin VB.Frame Fra_Comentario 
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
         Height          =   690
         Left            =   120
         TabIndex        =   119
         Top             =   2640
         Width           =   9390
         Begin VB.TextBox Txt_Comentarios_Cat_Productos 
            Height          =   420
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   120
            Top             =   195
            Width           =   9165
         End
      End
      Begin VB.Frame Fra_Almacen_Cat_Productos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacén"
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
         Height          =   510
         Left            =   6645
         TabIndex        =   116
         Top             =   2115
         Width           =   2865
         Begin VB.TextBox Txt_Existencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Lbl_Existencia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Existencia"
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
            TabIndex        =   118
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Frame Fra_Costos_Cat_Productos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Costos"
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
         Height          =   510
         Left            =   120
         TabIndex        =   94
         Top             =   2115
         Width           =   6495
         Begin VB.TextBox Txt_Costo_Con_IVA 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7065
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   210
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox Txt_IVA 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4635
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   165
            Width           =   1680
         End
         Begin VB.TextBox Txt_Costo_Cat_Productos 
            Height          =   285
            Left            =   1560
            TabIndex        =   34
            Top             =   165
            Width           =   2175
         End
         Begin VB.TextBox Txt_Precio_Venta 
            Height          =   285
            Left            =   8025
            TabIndex        =   37
            Top             =   210
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Precio Venta"
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
            Left            =   6780
            TabIndex        =   98
            Top             =   240
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Lbl_Costo_Con_IVA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costo"
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
            Left            =   6555
            TabIndex        =   97
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Lbl_IVA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
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
            Left            =   4155
            TabIndex        =   96
            Top             =   195
            Width           =   300
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costo sin IVA"
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
            Top             =   195
            Width           =   990
         End
      End
      Begin VB.Frame Fra_Generales_Productos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   87
         Top             =   300
         Width           =   9375
         Begin VB.TextBox Txt_Cantidad_Cajas 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8055
            TabIndex        =   124
            Top             =   1095
            Width           =   1230
         End
         Begin VB.ComboBox Cmb_Cajas 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":146F1
            Left            =   6030
            List            =   "Frm_Cat_Clientes.frx":146FB
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   1080
            Width           =   810
         End
         Begin VB.ComboBox Cmb_Estatus_Producto 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":14707
            Left            =   6030
            List            =   "Frm_Cat_Clientes.frx":14711
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   180
            Width           =   3255
         End
         Begin VB.TextBox Txt_Nombre_Cat_Productos 
            Height          =   285
            Left            =   1530
            TabIndex        =   104
            Top             =   525
            Width           =   7755
         End
         Begin VB.ComboBox Cmb_Aplica_IVA 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":14728
            Left            =   3705
            List            =   "Frm_Cat_Clientes.frx":14732
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1410
            Width           =   810
         End
         Begin VB.ComboBox Cmb_Cat_Productos_Categorias 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":1473E
            Left            =   1530
            List            =   "Frm_Cat_Clientes.frx":14740
            TabIndex        =   31
            Top             =   810
            Width           =   7755
         End
         Begin VB.ComboBox Cmb_Presentaciones_Cat_Productos 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":14742
            Left            =   6030
            List            =   "Frm_Cat_Clientes.frx":14744
            TabIndex        =   33
            Top             =   1410
            Width           =   3255
         End
         Begin VB.TextBox Txt_Clave_Cat_Productos 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1530
            TabIndex        =   30
            Top             =   1425
            Width           =   1230
         End
         Begin VB.TextBox Txt_Producto_ID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   195
            Width           =   2985
         End
         Begin VB.ComboBox Cmb_Cat_Producto_Tipo 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":14746
            Left            =   1530
            List            =   "Frm_Cat_Clientes.frx":14748
            TabIndex        =   32
            Top             =   1125
            Width           =   2985
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Cajas"
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
            Left            =   6885
            TabIndex        =   125
            Top             =   1125
            Width           =   1110
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aplica Cajas"
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
            Left            =   4665
            TabIndex        =   123
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Lbl_Estatus_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
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
            Left            =   4665
            TabIndex        =   115
            Top             =   225
            Width           =   660
         End
         Begin VB.Label LblNombre_CAt_Productos 
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
            Left            =   60
            TabIndex        =   105
            Top             =   555
            Width           =   660
         End
         Begin VB.Label Lbl_Aplica_IVA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aplica IVA"
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
            Left            =   2880
            TabIndex        =   103
            Top             =   1455
            Width           =   735
         End
         Begin VB.Label Lbl_Cat_Productos_Categorias 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Categoria"
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
            Left            =   60
            TabIndex        =   99
            Top             =   855
            Width           =   840
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clave"
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
            Left            =   60
            TabIndex        =   92
            Top             =   1455
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   4665
            TabIndex        =   91
            Top             =   1455
            Width           =   1350
         End
         Begin VB.Label Lbl_Producto_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto ID"
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
            Left            =   60
            TabIndex        =   89
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label Lbl_Tipo_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Tipo Producto"
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
            Left            =   60
            TabIndex        =   88
            Top             =   1170
            Width           =   1215
         End
      End
      Begin VB.Frame Fra_Detalles_Productos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Left            =   120
         TabIndex        =   86
         Top             =   3300
         Width           =   9375
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Productos 
            Height          =   2340
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   4128
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTOS"
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
         Left            =   3720
         TabIndex        =   90
         Top             =   -90
         Width           =   2445
      End
   End
   Begin VB.PictureBox Pic_Proveedores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   9855
      TabIndex        =   66
      Top             =   0
      Width           =   9855
      Begin VB.Frame Fra_Generales_Proveedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General"
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
         Height          =   3105
         Left            =   30
         TabIndex        =   68
         Top             =   420
         Width           =   9375
         Begin VB.ComboBox Cmd_Tipo_Pago 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":1474A
            Left            =   4342
            List            =   "Frm_Cat_Clientes.frx":14754
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2025
            Width           =   1395
         End
         Begin VB.TextBox Txt_Dias_Credio_Proveedor 
            Height          =   285
            Left            =   4342
            TabIndex        =   49
            Top             =   2385
            Width           =   1395
         End
         Begin VB.TextBox Txt_Estado_Proveedor 
            Height          =   285
            Left            =   1327
            TabIndex        =   45
            Top             =   2055
            Width           =   2085
         End
         Begin VB.TextBox Txt_Celular_Proveedores 
            Height          =   285
            Left            =   7005
            TabIndex        =   53
            Top             =   2025
            Width           =   2175
         End
         Begin VB.ComboBox Cmb_Clasificacion_Proveedor 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":1476A
            Left            =   6997
            List            =   "Frm_Cat_Clientes.frx":1476C
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   990
            Width           =   2190
         End
         Begin VB.TextBox Txt_colonia_Proveedor 
            Height          =   285
            Left            =   1327
            TabIndex        =   43
            Top             =   1320
            Width           =   4410
         End
         Begin VB.TextBox Txt_Direccion_Proveedor 
            Height          =   285
            Left            =   1327
            TabIndex        =   42
            Top             =   960
            Width           =   4410
         End
         Begin VB.TextBox Txt_CP_Proveedores 
            Height          =   285
            Left            =   4342
            TabIndex        =   47
            Top             =   1680
            Width           =   1395
         End
         Begin VB.TextBox Txt_Ciudad 
            Height          =   285
            Left            =   1327
            TabIndex        =   44
            Top             =   1680
            Width           =   2085
         End
         Begin VB.TextBox Txt_comentarios_proveedores 
            Height          =   315
            Left            =   1327
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   2715
            Width           =   7860
         End
         Begin VB.TextBox Txt_PRFC_Proveedores 
            Height          =   285
            Left            =   7005
            TabIndex        =   51
            Top             =   1365
            Width           =   2175
         End
         Begin VB.TextBox Txt_Nombre_Proveedor 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1327
            MaxLength       =   100
            TabIndex        =   40
            Top             =   600
            Width           =   7860
         End
         Begin VB.TextBox Txt_Proveedor_ID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1327
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txt_Fax_Proveedores 
            Height          =   285
            Left            =   7005
            TabIndex        =   54
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox Txt_Correo_Proveedores 
            Height          =   285
            Left            =   1327
            TabIndex        =   46
            Top             =   2400
            Width           =   2085
         End
         Begin VB.TextBox Txt_Telefono_Proveedores 
            Height          =   285
            Left            =   7005
            TabIndex        =   52
            Top             =   1680
            Width           =   2175
         End
         Begin VB.ComboBox Cmb_Estatus_Proveedor 
            Height          =   315
            ItemData        =   "Frm_Cat_Clientes.frx":1476E
            Left            =   6997
            List            =   "Frm_Cat_Clientes.frx":14778
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   225
            Width           =   2190
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Pago"
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
            Left            =   3405
            TabIndex        =   121
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Credito"
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
            Left            =   3405
            TabIndex        =   93
            Top             =   2415
            Width           =   900
         End
         Begin VB.Label Lbl_Clasificacion_Proveedor 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Clasificación "
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
            Left            =   5760
            TabIndex        =   84
            Top             =   1035
            Width           =   1170
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Colonia"
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
            Left            =   150
            TabIndex        =   82
            Top             =   1350
            Width           =   645
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
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
            Left            =   150
            TabIndex        =   81
            Top             =   990
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad"
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
            Left            =   150
            TabIndex        =   80
            Top             =   1710
            Width           =   510
         End
         Begin VB.Label Label12 
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
            Left            =   150
            TabIndex        =   79
            Top             =   2760
            Width           =   915
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CP"
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
            Left            =   3405
            TabIndex        =   78
            Top             =   1710
            Width           =   225
         End
         Begin VB.Label Label10 
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
            Left            =   5760
            TabIndex        =   77
            Top             =   1395
            Width           =   390
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
            Left            =   150
            TabIndex        =   76
            Top             =   630
            Width           =   660
         End
         Begin VB.Label Lbl_Proveedor_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor ID"
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
            Left            =   150
            TabIndex        =   75
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Left            =   5760
            TabIndex        =   74
            Top             =   2430
            Width           =   285
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Celular"
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
            Left            =   5760
            TabIndex        =   73
            Top             =   2055
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
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
            Left            =   5760
            TabIndex        =   72
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
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
            Left            =   150
            TabIndex        =   71
            Top             =   2430
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
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
            Left            =   5760
            TabIndex        =   70
            Top             =   270
            Width           =   660
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   150
            TabIndex        =   69
            Top             =   2017
            Width           =   1170
         End
      End
      Begin VB.Frame Fra_Proveedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   30
         TabIndex        =   67
         Top             =   3480
         Width           =   9375
         Begin MSFlexGridLib.MSFlexGrid Grid_Proveedores 
            Height          =   2205
            Left            =   150
            TabIndex        =   56
            Top             =   240
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Proveedores 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDORES"
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
         Left            =   3405
         TabIndex        =   83
         Top             =   30
         Width           =   2955
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   2835
         Picture         =   "Frm_Cat_Clientes.frx":1478F
         Top             =   75
         Width           =   360
      End
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Método de Pago"
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
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "Frm_Cat_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Correcto As Integer         'Utilizada para validar que se modifique correctamente el catalogo sin salirse del sistema si es que manda error
Dim Mi_Ayudante As Ayudante
Dim Cambio As Boolean           'Bandera para poder realizar cambios al precio de venta

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Agregar_Presentaciones_Click
'DESCRIPCION            : Da de alta los productos contenidas en el grid de presentaciones
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''
Private Sub Btn_Agregar_Presentaciones_Click()
Dim Rs_Alta_Presentaciones As rdoResultset
Dim Rs_Consulta As rdoResultset
Dim Cont_Fila As Integer
Dim Producto_ID As String
Dim Mi_SQL As String
Dim Presentacion_ID As String
Dim Categoria_ID As String
Dim Consulta_Cat_Productos As rdoResultset
Dim Rs_Consulta_Proveedor As rdoResultset
Dim Rs_Alta_Proveedor As rdoResultset
Dim Proveedor_ID As String


On Error GoTo handler
    Conexion_Base.BeginTrans
    
    Set Rs_Alta_Presentaciones = Conectar_Ayudante.Recordset_Agregar("Cat_Productos")
    
        For Cont_Fila = 1 To Grid_Cat_Productos.Rows - 1 Step 1
            Mi_SQL = " SELECT * FROM Cat_Productos WHERE Clave ='" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 0)) & "'"
            Set Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Consulta_Cat_Productos.EOF Then
                Producto_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos", "Producto_ID"), "00000")
                Mi_SQL = " SELECT * FROM Cat_Presentaciones WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 3)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Presentacion_ID = Rs_Consulta!Presentacion_ID
                Else
                    Presentacion_ID = "00001"
                End If
                Rs_Consulta.Close
                
                Mi_SQL = " SELECT * FROM Cat_Categorias WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Categoria_ID = Rs_Consulta!Categoria_ID
                Else
                    Categoria_ID = "00001"
                End If
                Rs_Consulta.Close
                
                'DA DE ALTA EL PROVEEDOR
                Set Rs_Consulta_Proveedor = Conectar_Ayudante.Recordset_Consultar("SELECT * From Cat_Proveedores WHERE Nombre='" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 2)) & "' ")
                If Rs_Consulta_Proveedor.EOF Then
                    Set Rs_Alta_Proveedor = Conectar_Ayudante.Recordset_Agregar("Cat_Proveedores")
                    Proveedor_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Proveedores", "Proveedor_ID"), "00000")
                    With Rs_Alta_Proveedor
                        .AddNew
                            .rdoColumns("Proveedor_ID") = Trim(Proveedor_ID)
                            .rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 2))
                            .rdoColumns("Direccion") = " "
                            .rdoColumns("Colonia") = " "
                            .rdoColumns("Ciudad") = " "
                            .rdoColumns("Estado") = " "
                            .rdoColumns("Correo_Electronico") = " "
                            .rdoColumns("Comentarios") = " "
                            .rdoColumns("Estatus") = "ACTIVO"
                            .rdoColumns("RFC") = " "
                            .rdoColumns("Codigo_Postal") = " "
                            .rdoColumns("Telefono") = " "
                            .rdoColumns("Celular") = " "
                            .rdoColumns("Prioridad") = 0
                            .rdoColumns("Fax") = " "
                            .rdoColumns("Dias_Credito") = 0
                            .rdoColumns("Usuario_Creo") = Usuario
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                    End With
                Else
                    Proveedor_ID = Rs_Consulta_Proveedor!Proveedor_ID
                End If
                Rs_Consulta_Proveedor.Close
                
                With Rs_Alta_Presentaciones
                    .AddNew
                        .rdoColumns("Producto_ID") = Producto_ID
                        .rdoColumns("Proveedor_ID") = Proveedor_ID
                        .rdoColumns("Marca_ID") = "00001"
                        .rdoColumns("Tipo_ID") = "00002"
                        .rdoColumns("Clave") = Grid_Cat_Productos.TextMatrix(Cont_Fila, 0)
                        .rdoColumns("Descripcion") = Grid_Cat_Productos.TextMatrix(Cont_Fila, 1)
                        '.rdoColumns("Nivel") = Grid_Cat_Productos.TextMatrix(Cont_Fila, 2)
                        .rdoColumns("Presentacion_ID") = Presentacion_ID
                        .rdoColumns("Categoria_ID") = Categoria_ID
                        .rdoColumns("Estatus") = "ACTIVO"
                        .rdoColumns("Comentarios") = " "
                        .rdoColumns("Costo") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 4))
                        .rdoColumns("Precio_Venta") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 5))
                        .rdoColumns("Precio_Venta_Oficial") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 5))
                        .rdoColumns("Usuario_Creo") = Nombre_Usuario
                        .rdoColumns("Fecha_Creo") = Now
                    .Update
                End With
            End If
        Next
        
    Rs_Alta_Presentaciones.Close
    Conexion_Base.CommitTrans
    MsgBox "Registros dados de Alta"
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    MsgBox Er.Description
End Sub

Private Sub Btn_Cancelar_Click()

End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Alta_Layout_Clientes_Click
'DESCRIPCION            : Alta_Clientes
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 07 - Noviembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''''
Private Sub Btn_Alta_Layout_Clientes_Click()
Dim Rs_Alta_Cat_Clientes As rdoResultset 'Del manejo de registro

On Error GoTo handler
    Conexion_Base.BeginTrans
    
    For Cont_Fila = 1 To Grid_Clientes.Rows - 1 Step 1
        'Alta de cliente
        Set Rs_Alta_Cat_Clientes = Conectar_Ayudante.Recordset_Agregar("Cat_Clientes")
        'Llena la tabla de Cat_Clientes con los datos contenidos en las cajas de textos
        With Rs_Alta_Cat_Clientes
        .AddNew
            .rdoColumns("Cliente_ID") = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Clientes", "Cliente_ID"), "00000")
            .rdoColumns("Clasificacion_ID") = Format(1, "00000")
            .rdoColumns("Status") = "A"
            .rdoColumns("Nombre") = Grid_Clientes.TextMatrix(Cont_Fila, 0)
            .rdoColumns("RFC") = " "
            .rdoColumns("Direccion") = Grid_Clientes.TextMatrix(Cont_Fila, 2)
            .rdoColumns("Colonia") = Grid_Clientes.TextMatrix(Cont_Fila, 3)
            .rdoColumns("CP") = " "
            .rdoColumns("Ciudad") = Grid_Clientes.TextMatrix(Cont_Fila, 1)
            .rdoColumns("Estado") = "Guanajuato"
            .rdoColumns("Telefono") = " "
            .rdoColumns("Celular") = " "
            .rdoColumns("Fax") = " "
            .rdoColumns("Email") = " "
            .rdoColumns("Comentarios") = UCase(" ")
            .rdoColumns("Usuario_Creo") = Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        End With
    Next
    Conexion_Base.CommitTrans
    
    MsgBox "Clientes dados de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Alta_Loyaut_Proveedores_Click
'DESCRIPCION            : Alta DE PROVEEDORES
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 08 - Noviembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''''
Private Sub Btn_Alta_Loyaut_Proveedores_Click()
Dim Rs_Alta_Cat_Proveedores As rdoResultset
Dim Cont_Fila As Integer

On Error GoTo handler
    Conexion_Base.BeginTrans
    
    For Cont_Fila = 1 To Grid_Proveedores.Rows - 1 Step 1
        'Alta de Proveedor
        Set Rs_Alta_Cat_Proveedores = Conectar_Ayudante.Recordset_Agregar("Cat_Proveedores")
        'Llena la tabla de Cat_Proveedores
        With Rs_Alta_Cat_Proveedores
        .AddNew
            .rdoColumns("Proveedor_ID") = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Proveedores", "Proveedor_ID"), "00000")
            .rdoColumns("Clasificacion_ID") = "00002"
            .rdoColumns("Nombre") = Grid_Proveedores.TextMatrix(Cont_Fila, 0)
            .rdoColumns("Direccion") = Grid_Proveedores.TextMatrix(Cont_Fila, 1)
            .rdoColumns("Colonia") = Grid_Proveedores.TextMatrix(Cont_Fila, 2)
            .rdoColumns("Ciudad") = Grid_Proveedores.TextMatrix(Cont_Fila, 3)
            .rdoColumns("Estado") = " "
            .rdoColumns("Correo_Electronico") = " "
            .rdoColumns("Comentarios") = " "
            .rdoColumns("Estatus") = "ACTIVO"
            .rdoColumns("RFC") = Grid_Proveedores.TextMatrix(Cont_Fila, 4)
            .rdoColumns("Codigo_Postal") = " "
            .rdoColumns("Telefono") = Grid_Proveedores.TextMatrix(Cont_Fila, 5)
            .rdoColumns("Celular") = " "
            .rdoColumns("Prioridad") = Val(0)
            .rdoColumns("Fax") = " "
            .rdoColumns("Dias_Credito") = Val(0)
            .rdoColumns("Forma_Pago") = " "
            .rdoColumns("Usuario_Creo") = Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        End With
        Rs_Alta_Cat_Proveedores.Close
    Next
    
    Conexion_Base.CommitTrans
    MsgBox "Proveedor dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Alta_Toma_Inventario_Click
'DESCRIPCION            : Da de Alta la toma de Inventario
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13 - Diciembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''''
Private Sub Btn_Alta_Toma_Inventario_Click()
Dim Mi_SQL As String
Dim Rs_Ope_Entrada As rdoResultset
Dim Rs_Ope_Entrada_Detalles As rdoResultset
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Pedido As rdoResultset
Dim Rs_Alta_Tmp_Facturas_Proveedores As rdoResultset
Dim Cont_Fila As Integer
Dim No_Control As String
Dim Utilidad_Producto As Double
Dim Rs_Consulta_Producto_ID As rdoResultset
Dim Entrada_ID As String
Dim Cantidad_Inventario As Double
Dim Opcion_Almacen As String
Dim No_Salida As String

On Error GoTo handler
    Conexion_Base.BeginTrans
    'SE DAN DE ALTA LOS DETALLES DE LA ENTRADA
    For Cont_Fila = 1 To Grid_Cat_Productos.Rows - 1 Step 1
    
        'ACTUALIZA EXISTENCIA EN EL CATALOGO DE PRODUCTOS
        Mi_SQL = " SELECT * FROM Cat_Productos "
        Mi_SQL = Mi_SQL & " WHERE Clave ='" & Grid_Cat_Productos.TextMatrix(Cont_Fila, 0) & "' "
        Set Rs_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Cat_Productos
            .Edit
                If Not IsNull(.rdoColumns("Existencia")) Then
                    If Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) = Val(.rdoColumns("Existencia")) Then
                        .rdoColumns("Existencia") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                    Else
                        If Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) > Val(.rdoColumns("Existencia")) Then
                            Cantidad_Inventario = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) - Val(.rdoColumns("Existencia"))
                            .rdoColumns("Existencia") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                            Opcion_Almacen = "ENTRADA"
                        Else
                            If Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) < Val(.rdoColumns("Existencia")) Then
                                Cantidad_Inventario = Val(.rdoColumns("Existencia")) - Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                                .rdoColumns("Existencia") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                                Opcion_Almacen = "SALIDA"
                            End If
                        End If
                    End If
                End If
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
        Rs_Cat_Productos.Close
    
        Select Case Opcion_Almacen
        
            Case "SALIDA"
                
                'SE CONSULTA EN CAT PRODUCTOS
                No_Salida = Format(Conectar_Ayudante.Maximo_Catalogo("Alm_Salidas_Almacen", "No_Salida"), "0000000000")
                Mi_SQL = " SELECT Cat_Productos.*,Cat_Marcas.Nombre as Marca,Cat_Impuestos.Impuesto FROM Cat_Impuestos,Cat_Marcas,Cat_Productos WHERE Cat_Productos.Clave ='" & Grid_Cat_Productos.TextMatrix(Cont_Fila, 0) & "' "
                Mi_SQL = Mi_SQL & " AND Cat_Productos.Marca_ID = Cat_Marcas.Marca_ID"
                Mi_SQL = Mi_SQL & " AND Cat_Productos.Impuesto_ID = Cat_Impuestos.Impuesto_ID"
                Set Rs_Consulta_Producto_ID = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Producto_ID.EOF Then
                    Cantidad_Total = 0
                    'SE DAN DE ALTA LOS DATOS GENERALES
                    Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Agregar("Alm_Salidas_Almacen")
                    With Rs_Ope_Entrada
                        .AddNew
                            .rdoColumns("No_Salida") = No_Salida
                            ''.rdoColumns("Cliente_ID")
                            .rdoColumns("Orden_Compra") = "00000"
                            .rdoColumns("Estatus") = "RECEPCION"
                            .rdoColumns("Comentarios") = " "
                            .rdoColumns("Cantidad_Total") = Cantidad_Inventario
                            .rdoColumns("Fecha_Salida") = Format(Now, "MM/dd/yyyy")
                            .rdoColumns("Referencia") = " "
                            .rdoColumns("Usuario_Creo") = Trim(Nombre_Usuario)
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                    End With
                    Rs_Ope_Entrada.Close
                    
                    'SE DAN DE ALTA LOS DETALLES
                    Set Rs_Ope_Entrada_Detalles = Conectar_Ayudante.Recordset_Agregar("Alm_Salidas_Almacen_Detalles")
                    With Rs_Ope_Entrada_Detalles
                        .AddNew
                            Rs_Ope_Entrada_Detalles.rdoColumns("No_Salida") = No_Salida
                            Rs_Ope_Entrada_Detalles.rdoColumns("Lote") = Grid_Cat_Productos.TextMatrix(Cont_Fila, 17)
                            Rs_Ope_Entrada_Detalles.rdoColumns("Fecha_Caducidad") = Format(Grid_Cat_Productos.TextMatrix(Cont_Fila, 18), "MM/dd/yyyy")
                            Rs_Ope_Entrada_Detalles.rdoColumns("Clave") = Rs_Consulta_Producto_ID!Clave
                            Rs_Ope_Entrada_Detalles.rdoColumns("Descripcion") = Rs_Consulta_Producto_ID!Nombre
                            Rs_Ope_Entrada_Detalles.rdoColumns("Cantidad") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Costo") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Importe") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) * Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Producto_ID") = Rs_Consulta_Producto_ID!Producto_ID
                            Rs_Ope_Entrada_Detalles.rdoColumns("Impuesto") = Val(Rs_Consulta_Producto_ID!Impuesto)
                            Rs_Ope_Entrada_Detalles.rdoColumns("IVA") = (Val(Rs_Consulta_Producto_ID!Impuesto) / 100) * (Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) * Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6)))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Total") = Cantidad_Inventario
                            Rs_Ope_Entrada_Detalles.rdoColumns("Facturado") = "NO"
                            'ACTUALIZA FALTANTE  TABLA ALM_ENTRADAS_ALMACEN_DETALLES
                            Mi_SQL = " SELECT * FROM Alm_Entradas_Detalles "
                            Mi_SQL = Mi_SQL & " WHERE No_Lote ='" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 17)) & "' "
                            Mi_SQL = Mi_SQL & " AND Clave ='" & Trim(Rs_Consulta_Producto_ID!Clave) & "' "
                            Set Rs_Alm_Entradas_Almacen = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            With Rs_Alm_Entradas_Almacen
                                .Edit
                                    If Not IsNull(.rdoColumns("Cantidad")) Then
                                        .rdoColumns("Faltante") = Val(.rdoColumns("Cantidad")) - Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                                    End If
                                .Update
                            End With
                            Rs_Alm_Entradas_Almacen.Close
                        .Update
                    End With
                    Rs_Ope_Entrada_Detalles.Close
                End If
                Rs_Consulta_Producto_ID.Close
            
            Case "ENTRADA"
                Entrada_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Alm_Entradas", "Entrada_ID"), "0000000000")
                'SE CONSULTA EN CAT PRODUCTOS
                Mi_SQL = " SELECT Cat_Productos.*,Cat_Marcas.Nombre as Marca,Cat_Impuestos.Impuesto FROM Cat_Impuestos,Cat_Marcas,Cat_Productos WHERE Cat_Productos.Clave ='" & Grid_Cat_Productos.TextMatrix(Cont_Fila, 0) & "' "
                Mi_SQL = Mi_SQL & " AND Cat_Productos.Marca_ID = Cat_Marcas.Marca_ID"
                Mi_SQL = Mi_SQL & " AND Cat_Productos.Impuesto_ID = Cat_Impuestos.Impuesto_ID"
                Set Rs_Consulta_Producto_ID = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Producto_ID.EOF Then
                    'SE DAN DE ALTA LOS DATOS GENERALES DE LA ENTRADA
                    Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas")
                    With Rs_Ope_Entrada
                        .AddNew
                            .rdoColumns("Entrada_ID") = Entrada_ID
                            .rdoColumns("Proveedor_ID") = Rs_Consulta_Producto_ID!Proveedor_ID
                            .rdoColumns("Fecha_Factura") = Format(Now, "MM/dd/yyyy")
                            .rdoColumns("Fecha_Recepcion_Factura") = Format(Now, "MM/dd/yyyy")
                            .rdoColumns("Tipo_Entrada") = "REMISION"
                            .rdoColumns("Total") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            ''.rdoColumns("Pedido_ID") = "00000"
                            .rdoColumns("Costo_Total") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) * Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            .rdoColumns("Estatus") = "FACTURADA"
                            .rdoColumns("Referencia") = " "
                            .rdoColumns("Observaciones") = " "
                            .rdoColumns("Usuario_Creo") = Trim(Nombre_Usuario)
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                    End With
                    Rs_Ope_Entrada.Close
                
                    Set Rs_Ope_Entrada_Detalles = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas_Detalles")
                    With Rs_Ope_Entrada_Detalles
                        .AddNew
                            Rs_Ope_Entrada_Detalles.rdoColumns("Entrada_ID") = Entrada_ID
                            Rs_Ope_Entrada_Detalles.rdoColumns("Producto_ID") = Rs_Consulta_Producto_ID!Producto_ID
                            Rs_Ope_Entrada_Detalles.rdoColumns("Clave") = Rs_Consulta_Producto_ID!Clave
                            Rs_Ope_Entrada_Detalles.rdoColumns("Descripcion") = Rs_Consulta_Producto_ID!Nombre
                            Rs_Ope_Entrada_Detalles.rdoColumns("Marca") = Rs_Consulta_Producto_ID!Marca
                            Rs_Ope_Entrada_Detalles.rdoColumns("Impuesto") = Val(Rs_Consulta_Producto_ID!Impuesto)
                            Rs_Ope_Entrada_Detalles.rdoColumns("IVA") = (Val(Rs_Consulta_Producto_ID!Impuesto) / 100) * (Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) * Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6)))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Costo") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Importe") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16)) * Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Cantidad") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                            Rs_Ope_Entrada_Detalles.rdoColumns("No_Lote") = Grid_Cat_Productos.TextMatrix(Cont_Fila, 17)
                            Rs_Ope_Entrada_Detalles.rdoColumns("Fecha_Caducidad") = Format(Grid_Cat_Productos.TextMatrix(Cont_Fila, 18), "MM/dd/yyyy")
                            Rs_Ope_Entrada_Detalles.rdoColumns("Faltante") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 16))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "FACTURADA"
                        .Update
                    End With
                    Rs_Ope_Entrada_Detalles.Close
                
                End If
                Rs_Consulta_Producto_ID.Close
                                    
        End Select
        
    Next Cont_Fila
    MsgBox "Inventario Actualizado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Conexion_Base.CommitTrans
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    MsgBox Err.Description
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Cargar_Layout_Click
'DESCRIPCION            : Carga el archivo de excel que contiene los productos
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01 - Noviembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''
Private Sub Btn_Cargar_Layout_Click()
Dim Clave As String
Dim Nombre As String
Dim Descripcion As String
Dim Presentacion As String
Dim SAL As String
Dim Laboratorio As String
Dim Costo_Maesba As String
Dim Precio_Max As String
Dim Precio_Publico As String
Dim Aplica_IVA As String
Dim Categoria As String
Dim Marca As String
Dim Proveedor As String
Dim Utilidad As String
Dim NIVEL As String
Dim Especialidad As String
Dim Pos_Final As Integer

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
    Grid_Cat_Productos.Rows = 0
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
            Grid_Cat_Productos.Redraw = False
            .Refresh
            Grid_Cat_Productos.Redraw = True
        End With
        Grid_Cat_Productos.Rows = 0
        Grid_Cat_Productos.Cols = 16
        Cont_Partidas = 0
        Clave = ""
        Nombre = ""
        Descripcion = ""
        Presentacion = ""
        SAL = ""
        Laboratorio = ""
        Costo_Maesba = ""
        Precio_Max = ""
        Precio_Publico = ""
        Aplica_IVA = ""
        Categoria = ""
        Marca = ""
        Proveedor = ""
        Utilidad = ""
        NIVEL = ""
        Especialidad = ""
        'Se agrega encabezado
        Grid_Cat_Productos.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Descripcion" & Chr(9) & "Presentacion" & Chr(9) & "SAL" _
        & Chr(9) & "Laboratorio" & Chr(9) & "Costo_Maesba" & Chr(9) & "Precio_Max" & Chr(9) & "Precio_Publico" & Chr(9) & "Aplica_IVA" _
        & Chr(9) & "Categoria" & Chr(9) & "Marca" & Chr(9) & "Proveedor" & Chr(9) & "Utilidad" _
        & Chr(9) & "NIVEL" & Chr(9) & "Especialidad"
        With Dt_Excel.Recordset
            While Not .EOF
                If Not IsNull(Dt_Excel.Recordset(0).Value) Then
                    If Not IsNull(Dt_Excel.Recordset(0).Value) Then Clave = Dt_Excel.Recordset(0).Value
                    If Not IsNull(Dt_Excel.Recordset(1).Value) Then Nombre = Dt_Excel.Recordset(1).Value
                    If Not IsNull(Dt_Excel.Recordset(2).Value) Then Descripcion = Dt_Excel.Recordset(2).Value
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then
                        Presentacion = Dt_Excel.Recordset(3).Value
                        Pos_Final = InStr(1, Presentacion, ".")
                        If Pos_Final > 0 Then
                            Presentacion = Mid(Presentacion, 1, Pos_Final - 1)
                        Else
                            Presentacion = Dt_Excel.Recordset(3).Value
                        End If
                    End If
                    If Not IsNull(Dt_Excel.Recordset(4).Value) Then SAL = Dt_Excel.Recordset(4).Value
                    If Not IsNull(Dt_Excel.Recordset(5).Value) Then Laboratorio = Dt_Excel.Recordset(5).Value
                    If Not IsNull(Dt_Excel.Recordset(6).Value) Then Costo_Maesba = Dt_Excel.Recordset(6).Value
                    If Not IsNull(Dt_Excel.Recordset(7).Value) Then Precio_Max = Dt_Excel.Recordset(7).Value
                    If Not IsNull(Dt_Excel.Recordset(8).Value) Then Precio_Publico = Dt_Excel.Recordset(8).Value
                    If Not IsNull(Dt_Excel.Recordset(9).Value) Then Aplica_IVA = Dt_Excel.Recordset(9).Value
                    If Not IsNull(Dt_Excel.Recordset(10).Value) Then Categoria = Dt_Excel.Recordset(10).Value
                    If Not IsNull(Dt_Excel.Recordset(11).Value) Then Marca = Dt_Excel.Recordset(11).Value
                    If Not IsNull(Dt_Excel.Recordset(12).Value) Then Proveedor = Dt_Excel.Recordset(12).Value
                    If Not IsNull(Dt_Excel.Recordset(13).Value) Then Utilidad = Dt_Excel.Recordset(13).Value
                    If Not IsNull(Dt_Excel.Recordset(14).Value) Then NIVEL = Dt_Excel.Recordset(14).Value
                    If Not IsNull(Dt_Excel.Recordset(15).Value) Then Especialidad = Dt_Excel.Recordset(15).Value
                    Grid_Cat_Productos.AddItem Clave & Chr(9) & Nombre & Chr(9) & Descripcion & Chr(9) & _
                    Presentacion & Chr(9) & SAL & Chr(9) & Laboratorio & Chr(9) & Costo_Maesba & Chr(9) & _
                    Precio_Max & Chr(9) & Precio_Publico & Chr(9) & Aplica_IVA & Chr(9) & Categoria & Chr(9) & Marca & Chr(9) & Proveedor & Chr(9) & Utilidad & Chr(9) & NIVEL & Chr(9) & Especialidad
                    Cont_Partidas = Cont_Partidas + 1
                End If
                .MoveNext
            Wend
            Txt_Producto_ID.text = Val(Cont_Partidas)
            
            'Configura el tamaño las columnas del Grid
            If Grid_Cat_Productos.Rows > 1 Then
                Grid_Cat_Productos.FixedRows = 1
                Grid_Cat_Productos.ColWidth(0) = 1500 '
                Grid_Cat_Productos.ColWidth(1) = 5000 '
                Grid_Cat_Productos.ColWidth(2) = 2000 '
                Grid_Cat_Productos.ColWidth(3) = 1000 '
                Grid_Cat_Productos.ColWidth(4) = 1000 '
                Grid_Cat_Productos.ColWidth(5) = 1000 '
                Grid_Cat_Productos.ColWidth(6) = 1000 '
                Grid_Cat_Productos.ColWidth(7) = 1000 '
                Grid_Cat_Productos.ColWidth(8) = 1000 '
                Grid_Cat_Productos.ColWidth(9) = 1000 '
                Grid_Cat_Productos.ColWidth(10) = 1000 '
                Grid_Cat_Productos.ColWidth(11) = 1000 '
                Grid_Cat_Productos.ColWidth(12) = 1000 '
                Grid_Cat_Productos.ColWidth(13) = 1000 '
                Grid_Cat_Productos.ColWidth(14) = 1000 '
                Grid_Cat_Productos.ColWidth(15) = 1000 '
                'Pone el setfocus en la primera fila del Grid
                With Grid_Cat_Productos
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
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
'NOMBRE DE LA FUNCION   : Btn_Cargar_Presentaciones_Click
'DESCRIPCION            : Carga los productos del archivo de Excel seleccionado
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Agosto-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''
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
Dim Precio_Oficial As String

    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Grid_Cat_Productos.Rows = 0
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
            Grid_Cat_Productos.Redraw = False
            .Refresh
            Grid_Cat_Productos.Redraw = True
        End With
        Grid_Cat_Productos.Rows = 0
        Grid_Cat_Productos.Cols = 7
        Descripcion = ""
        Clave = ""
        NIVEL = ""
        Presentacion = ""
        Especialidad = ""
        
               
        'Se agrega encabezado
        Grid_Cat_Productos.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Nivel" & Chr(9) & "Presentacion" & Chr(9) & "Especialidad"
        With Dt_Excel.Recordset
            While Not .EOF
                    If Not IsNull(Dt_Excel.Recordset(0).Value) Then Clave = Dt_Excel.Recordset(0).Value
                    If Not IsNull(Dt_Excel.Recordset(1).Value) Then Descripcion = Dt_Excel.Recordset(1).Value
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then NIVEL = Dt_Excel.Recordset(3).Value
                    If Not IsNull(Dt_Excel.Recordset(5).Value) Then Presentacion = Dt_Excel.Recordset(5).Value
                    If Not IsNull(Dt_Excel.Recordset(6).Value) Then Especialidad = Dt_Excel.Recordset(6).Value
                    If Not IsNull(Dt_Excel.Recordset(11).Value) Then Precio_Oficial = Dt_Excel.Recordset(11).Value
                    If Not IsNull(Dt_Excel.Recordset(4).Value) Then Categoria = Dt_Excel.Recordset(4).Value
                    Agregar_Presentacion = "SI"
                    For Cont_Fila = 1 To Grid_Cat_Productos.Rows - 1 Step 1
                        If Trim(Descripcion) = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 2)) Then
                            'no agregar
                            Agregar_Presentacion = "NO"
                            Exit For
                        End If
                    Next Cont_Fila
                    If Agregar_Presentacion = "SI" And Trim(Descripcion) <> "" Then
                        Grid_Cat_Productos.AddItem Clave & Chr(9) & Descripcion & Chr(9) & NIVEL & Chr(9) & Presentacion & Chr(9) & Especialidad & Chr(9) & Precio_Oficial & Chr(9) & Categoria
                    End If
                    Descripcion = ""
                .MoveNext
            Wend
            
            'Configura el tamaño las columnas del Grid
            If Grid_Cat_Productos.Rows > 1 Then
                Grid_Cat_Productos.FixedRows = 1
                Grid_Cat_Productos.ColWidth(0) = 1500 'Descripcion
                Grid_Cat_Productos.ColWidth(1) = 5000 'Descripcion
                Grid_Cat_Productos.ColWidth(2) = 2000 'Descripcion
                Grid_Cat_Productos.ColWidth(3) = 1000 'Descripcion
                Grid_Cat_Productos.ColWidth(4) = 1000 'Descripcion
                Grid_Cat_Productos.ColWidth(5) = 1000 'Descripcion
                Grid_Cat_Productos.ColWidth(6) = 1000 'Descripcion
                'Pone el setfocus en la primera fila del Grid
                With Grid_Cat_Productos
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
        End With
        Exit Sub
handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub

Private Sub Btn_Cerrar_Click()
    Pic_Busqueda_Cat_Productos.Visible = False
    Chk_Descripcion_Producto.Value = 0
    Chk_Clave_Producto.Value = 0
End Sub

Private Sub Btn_Consultar_Cat_Productos_Click()
    Call Btn_Consultar_Click
End Sub

Private Sub Btn_Consultar_Click()
Dim Valor As String             'Valor tecleado por el usuario en el InputBox
    Select Case Catalogo
        Case "CLIENTES"
            Valor = InputBox("Teclee el nombre del Cliente a buscar", "Busqueda de clientes")
            Call Consulta_Clientes(Valor)
        Case "PROVEEDORES"
            Valor = InputBox("Teclee el nombre del Proveedor a buscar", "Busqueda de Proveedor")
            Call Consulta_Proveedor(Valor)
        Case "PRODUCTOS"
            Valor = InputBox("Teclee el nombre del Producto a buscar", "Busqueda de Productos")
            Call Consulta_Cat_Productos(Valor, "DESCRIPCION")
    End Select
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta de Clientes
    'DESCRIPCIÓN: Consulta fragmentos del nombre obtenidos de un InputBox dando
    '             como resultado todos nombres que empiesen con ese nombre
    'PARÁMETROS:
    '             1. Nombre: Texto_Busqueda. Es usada para buscar el nombre en
    '                        la Base de Datos
    'CREO:
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Consulta_Clientes(Texto_Busqueda As String)
Dim Rs_Consulta_Cat_Clientes As rdoResultset    'Manejo de registro
    
    Grid_Clientes.Rows = 0
    'Consulta los clientes de acuerdo al parametro
    Mi_SQL = "SELECT Cliente_ID, Nombre, RFC"
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes "
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Texto_Busqueda & "%'" & " ORDER BY Cliente_ID"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'LLena el grid con los datos el resultado de la busqueda anterior
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        Grid_Clientes.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "RFC"
        While Not Rs_Consulta_Cat_Clientes.EOF
            Grid_Clientes.AddItem Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & Chr(9) & _
            Rs_Consulta_Cat_Clientes.rdoColumns("Nombre") & Chr(9) & _
            Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
            Grid_Clientes.FixedRows = 1
            Rs_Consulta_Cat_Clientes.MoveNext
        Wend
        Rs_Consulta_Cat_Clientes.Close
        'Configura el grid
        Grid_Clientes.ColWidth(0) = 1000
        Grid_Clientes.ColWidth(1) = 6300
        Grid_Clientes.ColWidth(2) = 1400
        'Manda llamara la función Grid_Clientes_Click
        Grid_Clientes_Click
        Tab_Clientes.Tab = 0
    End If
End Sub

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Dar_Alta_Layout_Click
'DESCRIPCION            : Da de alta los productos en la tabla cat_Catlogos; contenidas en el grid
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 02-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''
Private Sub Btn_Dar_Alta_Layout_Click()
Dim Rs_Alta_Cat_Productos As rdoResultset
Dim Rs_Consulta As rdoResultset
Dim Cont_Fila As Integer
Dim Producto_ID As String
Dim Mi_SQL As String
Dim Presentacion_ID As String
Dim Categoria_ID As String
Dim Rs_Agregar_Registro As rdoResultset
Dim Tipo_ID As String
Dim Proveedor_ID As String
Dim Marca_ID As String
Dim Laboratorio_ID As String
Dim Sustancia_Activa_ID As String
Dim Consulta_Cat_Productos As rdoResultset
Dim Claves_Repetidas As String



On Error GoTo handler
    Conexion_Base.BeginTrans
    
    Set Rs_Alta_Cat_Productos = Conectar_Ayudante.Recordset_Agregar("Cat_Productos")
        For Cont_Fila = 1 To Grid_Cat_Productos.Rows - 1 Step 1
            Mi_SQL = " SELECT * FROM Cat_Productos WHERE Clave ='" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 0)) & "'"
            Set Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Consulta_Cat_Productos.EOF Then  ''Or Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 0)) = "7501055328369" Then
                Producto_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos", "Producto_ID"), "00000")
                'Cat_Presentaciones
                Mi_SQL = " SELECT * FROM Cat_Presentaciones WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 3)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Presentacion_ID = Rs_Consulta!Presentacion_ID
                Else
                    Presentacion_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Presentaciones", "Presentacion_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Presentaciones")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Presentacion_ID") = Presentacion_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 3))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
                'Cat_Categorias
                Mi_SQL = " SELECT * FROM Cat_Categorias WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 15)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Categoria_ID = Rs_Consulta!Categoria_ID
                Else
                    Categoria_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Categorias", "Categoria_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Categorias")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Categoria_ID") = Categoria_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 15))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
                'Cat_Sustancia_Activa
                Mi_SQL = " SELECT * FROM Cat_Sustancia_Activa WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 4)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Sustancia_Activa_ID = Rs_Consulta!Sustancia_Activa_ID
                Else
                    Sustancia_Activa_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Sustancia_Activa", "Sustancia_Activa_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Sustancia_Activa")
                        With Rs_Agregar_Registro
                            .AddNew
                            Rs_Agregar_Registro.rdoColumns("Sustancia_Activa_ID") = Sustancia_Activa_ID
                            Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 4))
                            Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                            Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                            Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                            Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
                'Cat_Laboratorios
                Mi_SQL = " SELECT * FROM Cat_Laboratorios WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 5)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Laboratorio_ID = Rs_Consulta!Laboratorio_ID
                Else
                    Laboratorio_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Laboratorios", "Laboratorio_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Laboratorios")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Laboratorio_ID") = Laboratorio_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 5))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
                'Cat_Marcas
                Mi_SQL = " SELECT * FROM Cat_Marcas WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 11)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Marca_ID = Rs_Consulta!Marca_ID
                Else
                    Marca_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Marcas", "Marca_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Marcas")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Marca_ID") = Marca_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 11))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
                'Cat_proveedores
                Mi_SQL = " SELECT * FROM Cat_proveedores WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 12)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Proveedor_ID = Rs_Consulta!Proveedor_ID
                Else
                    Proveedor_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_proveedores", "Proveedor_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_proveedores")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Proveedor_ID") = Proveedor_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 12))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
    
    
                'Cat_Productos_Tipo
                Mi_SQL = " SELECT * FROM Cat_Productos_Tipo WHERE Nombre = '" & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 10)) & "'"
                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta.EOF Then
                    Tipo_ID = Rs_Consulta!Tipo_ID
                Else
                    Tipo_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Tipo", "Tipo_ID"), "00000")
                    Set Rs_Agregar_Registro = Conectar_Ayudante.Recordset_Agregar("Cat_Productos_Tipo")
                        With Rs_Agregar_Registro
                            .AddNew
                                Rs_Agregar_Registro.rdoColumns("Tipo_ID") = Tipo_ID
                                Rs_Agregar_Registro.rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 10))
                                Rs_Agregar_Registro.rdoColumns("Estatus") = "ACTIVO"
                                Rs_Agregar_Registro.rdoColumns("Comentarios") = " "
                                Rs_Agregar_Registro.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Agregar_Registro.rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                    Rs_Agregar_Registro.Close
                End If
                Rs_Consulta.Close
                
                With Rs_Alta_Cat_Productos
                    .AddNew
                        .rdoColumns("Producto_ID") = Producto_ID
                        .rdoColumns("Proveedor_ID") = "00001" 'Proveedor_ID
                        .rdoColumns("Marca_ID") = "00001" 'Marca_ID
                        .rdoColumns("Presentacion_ID") = Presentacion_ID
                        If Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 10)) = "MEDICAMENTO" Or Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 10)) = "Medicamento" Then
                            .rdoColumns("Tipo_ID") = "00001"  'Tipo_ID
                            .rdoColumns("Categoria_ID") = Categoria_ID
                        Else
                            .rdoColumns("Tipo_ID") = "00003"  'Tipo_ID
                            .rdoColumns("Categoria_ID") = Categoria_ID
                        End If
                        .rdoColumns("Laboratorio_ID") = Laboratorio_ID
                        .rdoColumns("Sustancia_Activa_ID") = Sustancia_Activa_ID
                        .rdoColumns("Nombre") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 1))
                        .rdoColumns("Descripcion") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 2))
                        .rdoColumns("Clave") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 0))
                        .rdoColumns("Utilidad") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 13))
                        .rdoColumns("Costo") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 6))
                        .rdoColumns("Precio_Venta") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 7))
                        .rdoColumns("Precio_Venta_Oficial") = Val(Grid_Cat_Productos.TextMatrix(Cont_Fila, 8))
                        .rdoColumns("Existencia") = 0
                        .rdoColumns("Minimo") = 0
                        .rdoColumns("Maximo") = 0
                        .rdoColumns("Reorden") = 0
                        .rdoColumns("Nivel") = Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 14))
                        .rdoColumns("Negativos") = "NO"
                        If Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 9)) = "SI" Then
                            .rdoColumns("Aplica_IVA") = "SI"
                        Else
                            .rdoColumns("Aplica_IVA") = "NO"
                        End If
                        If Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 9)) = "SI" Then
                            .rdoColumns("Impuesto_ID") = "00001"
                        Else
                            .rdoColumns("Impuesto_ID") = "00003"
                        End If
                        .rdoColumns("Estatus") = "ACTIVO"
                        .rdoColumns("Comentarios") = " "
                        .rdoColumns("Usuario_Creo") = Nombre_Usuario
                        .rdoColumns("Fecha_Creo") = Now
                    .Update
                End With
            Else
                Claves_Repetidas = Claves_Repetidas & Chr(9) & Trim(Grid_Cat_Productos.TextMatrix(Cont_Fila, 0))
            End If
        Next Cont_Fila
        
    Rs_Alta_Cat_Productos.Close
    Conexion_Base.CommitTrans
    MsgBox "Registros dados de Alta"
    MsgBox Claves_Repetidas
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    MsgBox Er.Description
End Sub


'Botón para eliminar un registro del catálogo seleccionado
Private Sub Btn_Eliminar_Click()
Set Conectar_Ayudante = New Ayudante
    
    Select Case Catalogo
        'Catálogo de Clientes
        Case "CLIENTES":
            'Si el txt_Cliente_ID no esta vacio entonces le pregunta al usuario si esta seguro de eliminar los datos del mismo
            If Trim(Txt_Cliente_ID.text) <> "" Then
                '1. Elimina el Registro
                '2. Quita los datos del usuario contenidos en el Grid
                If MsgBox("¿Esta seguro de eliminar al cliente?", vbYesNo + vbQuestion) = vbYes Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Clientes", "Cliente_ID", Txt_Cliente_ID) = True Then
                        MsgBox "El cliente ha sido eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        'Elimina la informacion de los datos del cliente del grid
                        If Grid_Clientes.Rows = 2 Then
                            Grid_Clientes.Rows = 0
                        Else
                            Grid_Clientes.RemoveItem Grid_Clientes.RowSel
                        End If
                        'Limpia todos los textos y los combos de la forma
                        Call Conectar_Ayudante.Limpiar_Textos(Me)
                        Tab_Clientes.Tab = 0
                        Txt_Cliente_ID.text = ""
                        Cmb_Status_Cliente.ListIndex = -1
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        'Catálogo de Proveedores
        Case "PROVEEDORES":
            If Trim(Txt_Proveedor_ID.text) <> "" Then
                If MsgBox("¿Esta seguro de eliminar al Proveedor?", vbYesNo + vbQuestion) = vbYes Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Proveedores", "Proveedor_ID", Txt_Proveedor_ID) = True Then
                        MsgBox "El Proveedor ha sido eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        If Grid_Proveedores.Rows = 2 Then
                            Grid_Proveedores.Rows = 0
                        Else
                            Grid_Proveedores.RemoveItem Grid_Proveedores.RowSel
                        End If
                        'Limpia todos los textos y los combos de la forma
                        Call Conectar_Ayudante.Limpiar_Textos(Me)
                        Txt_Proveedor_ID.text = ""
                        Cmb_Estatus_Proveedor.ListIndex = -1
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
            
            
        'Catálogo de Productos
        Case "PRODUCTOS":
            If Trim(Txt_Producto_ID.text) <> "" Then
                If MsgBox("¿Esta seguro de eliminar el Producto?", vbYesNo + vbQuestion) = vbYes Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Productos", "Producto_ID", Txt_Producto_ID) = True Then
                        MsgBox "El Producto ha sido eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                        If Grid_Cat_Productos.Rows = 2 Then
                            Grid_Cat_Productos.Rows = 0
                        Else
                            Grid_Cat_Productos.RemoveItem Grid_Cat_Productos.RowSel
                        End If
                        'Limpia todos los textos y los combos de la forma
                        Call Conectar_Ayudante.Limpiar_Textos(Me)
                        Txt_Producto_ID.text = ""
                        Cmb_Estatus_Producto.ListIndex = -1
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
    End Select
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Btn_Generar_Archivo_DBF_Click
'DESCRIPCIÓN         : Genera un archivo dbf para la exportacion de informacion
'PARÁMETROS          :
'CREO                : Julio Cruz
'FECHA_CREO          : 04-Nov-2010
'MODIFICO            :
'FECHA_MODIFICO      :
'CAUSA_MODIFICACIÓN  :
'*******************************************************************************
Private Sub Btn_Generar_Archivo_DBF_Click()
Dim Rs_Consulta_Catalogos As rdoResultset       'Informacion de la sucursal
Dim Rs_Consulta_Cat_Parametros As rdoResultset
Dim Comando_DBF As New ADODB.Command            'Se utiliza para crear el archivo dbf por medio de comando
Dim Conexion_DBF As ADODB.Connection            'Conexion a la base de datos
Dim Rs_Tabla_DBF As ADODB.Recordset             'Recorset de la info de la base de datos
Dim Cadena_Conexion_DBF As String
Dim Cadena_Crear_Tabla As String
Dim Nombre_ArchivoP As String
Dim Ejecucion As String                         'Cadena de ejecucion para empaquetado
Dim Tipo As String                              'Tipo de campo de las tablas
Dim Valor As String                             'Valor del tipo de campo
Dim Columnas As Integer                         'Columnas a importar
Dim Columna As Integer                          'Recorrido de las columnas
Dim I As Integer                                'Contador del ciclo
Dim hProcess As Long                            'Indica que se ejecuta el proceso de MS-DOS

On Error GoTo handler
    Me.MousePointer = 11
    
    '************************Archivo de Informacion detallada**********************
    'SE ESTABLECE EL NOMBRE DEL ARCHIVO
    Nombre_ArchivoP = "CProduct"
    'CADENA DE CONEXION
    Cadena_Conexion_DBF = "Provider=MSDASQL.1;Persist Security Info=False;" & _
    "Extended Properties=Driver={Driver para o Microsoft Visual FoxPro};UID=;" & _
    "SourceDB=" & App.Path & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine"
    
    'SE ESTABLECE UNA CONEXION CON LA BASE DE DATOS
    Set Conexion_DBF = New ADODB.Connection
    With Conexion_DBF
        '.CursorLocation = adUseClient
        .ConnectionString = Cadena_Conexion_DBF
        .Open
     End With
                  
    If Conexion_DBF.State = 0 Then
        MsgBox "No se pudo establecer la conexion", vbInformation
        Exit Sub
    End If
    
    'SE HABILITA LA OPCION DE EJECUCIÓN POR TEXTO
    Set Comando_DBF = New ADODB.Command
    With Comando_DBF
        .ActiveConnection = Conexion_DBF
        .CommandType = adCmdText
    End With
    
    'SE CREA LA TABLA
    Cadena_Crear_Tabla = "CREATE TABLE " & Nombre_ArchivoP & "("
    For I = 1 To 16 Step 1
        Select Case I
            'Tipo C  char, text, varchar, nchar, ntext, nvarchar
            'Tipo T  datetime
            'Tipo D  smalldatetime
            'Tipo N  bigint, decimal, int, numeric, smallInt, money, tinyint, smallmoney, bit
            Case 1
                Tipo = "C"
                Valor = "(20)"
            Case 2
                Tipo = "C"
                Valor = "(200)"
            Case 3
                Tipo = "C"
                Valor = "(200)"
            Case 4
                Tipo = "C"
                Valor = "(100)"
            Case 5
                Tipo = "C"
                Valor = "(100)"
            Case 6
                Tipo = "C"
                Valor = "(100)"
            Case 7 'Costo
                Tipo = "C"
                Valor = "(50)"
            Case 8
                Tipo = "C"
                Valor = "(50)"
            Case 9 'Precio Publico
                Tipo = "C"
                Valor = "(50)"
            Case 10
                Tipo = "C"
                Valor = "(2)"
            Case 11
                Tipo = "C"
                Valor = "(20)"
            Case 12
                Tipo = "C"
                Valor = "(100)"
            Case 13
                Tipo = "C"
                Valor = "(100)"
            Case 14 'Utilidad
                Tipo = "C"
                Valor = "(50)"
            Case 15
                Tipo = "C"
                Valor = "(20)"
            Case 16
                Tipo = "C"
                Valor = "(20)"
        End Select
        Cadena_Crear_Tabla = Cadena_Crear_Tabla & "C" & Format(I, "000") & " " & Tipo & Valor & ","
    Next
    Cadena_Crear_Tabla = Mid(Cadena_Crear_Tabla, 1, Len(Cadena_Crear_Tabla) - 1)
    Cadena_Crear_Tabla = Cadena_Crear_Tabla & " )"
    
    'SE CREA EL ARCHIVO
    With Comando_DBF
        .CommandText = Cadena_Crear_Tabla
        Set Rs_Tabla_DBF = .Execute
    End With
    
    'SE INGRESA LA INFORMACION DENTRO DE LA TABLA CREADA
    'CONSULTA LOS PRODUCTOS
    Columnas = 0
    Mi_SQL = "SELECT "
    Mi_SQL = Mi_SQL & " Cat_Productos.Clave,Cat_Productos.Nombre,Cat_Productos.Descripcion,Cat_Presentaciones.Nombre as Presentacion_ID,Cat_Sustancia_Activa.Nombre as Sustancia_Activa_ID,Cat_Laboratorios.Nombre as Laboratorio_ID,Cat_Productos.Costo,Cat_Productos.Precio_Venta,"
    Mi_SQL = Mi_SQL & " Cat_Productos.Precio_Venta_Oficial,Cat_Productos.Aplica_IVA,Cat_Productos_Tipo.Nombre as Tipo_ID,Cat_Marcas.Nombre as Marca_ID,Cat_Proveedores.Nombre as Proveedor_ID,Cat_Productos.Utilidad,Cat_Productos.Nivel,Cat_Categorias.NOmbre as Categoria_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Productos,Cat_Sustancia_Activa,Cat_Laboratorios,Cat_Productos_Tipo,Cat_Marcas,Cat_Proveedores,Cat_Categorias,Cat_Presentaciones "
    Mi_SQL = Mi_SQL & " WHERE Cat_Productos.Categoria_ID ='00032'"
    Mi_SQL = Mi_SQL & " AND Cat_Sustancia_Activa.Sustancia_Activa_ID = Cat_Productos.Sustancia_Activa_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Productos_Tipo.Tipo_ID = Cat_Productos.Tipo_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Marcas.Marca_ID = Cat_Productos.Marca_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Proveedores.Proveedor_ID = Cat_Productos.Proveedor_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Categorias.Categoria_ID = Cat_Productos.Categoria_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Laboratorios.Laboratorio_ID = Cat_Productos.Laboratorio_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Presentaciones.Presentacion_ID = Cat_Productos.Presentacion_ID"
    Set Rs_Consulta_Catalogos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Catalogos.EOF Then
        While Not Rs_Consulta_Catalogos.EOF
            Mi_SQL = "INSERT INTO " & Nombre_ArchivoP & " VALUES"
            Mi_SQL = Mi_SQL & "("
            For Columna = 0 To 15
                If IsNull(Rs_Consulta_Catalogos.rdoColumns(Columna)) Then
                    Valor = "NULL"
                Else
                    Valor = Rs_Consulta_Catalogos.rdoColumns(Columna)
                End If
                Mi_SQL = Mi_SQL & "'" & Trim(Replace(Valor, "'", "''")) & "',"
                Debug.Print Mi_SQL
            Next
            Mi_SQL = Mid(Mi_SQL, 1, Len(Mi_SQL) - 1)
            Mi_SQL = Mi_SQL & ")"
            With Comando_DBF
                .CommandText = Mi_SQL
            Set Rs_Tabla_DBF = .Execute
            Rs_Consulta_Catalogos.MoveNext
        End With
        Wend
    Else
        Exit Sub
    End If
    
    Conexion_DBF.Close
    'Rs_Tabla_DBF.Close
    Set Comando_DBF = Nothing
    Set Rs_Tabla_DBF = Nothing
    Set Conexion_DBF = Nothing
    
    '/*Archivo de Informacion*/
    If Len(Dir$(App.Path & "\" & Nombre_ArchivoP & ".DBF")) > 0 Then
        FileCopy App.Path & "\" & Nombre_ArchivoP & ".DBF", App.Path & "\" & Nombre_ArchivoP & ".cat"
        'Si lo creo entonces borra el archivo .DBF
        If Len(Dir$(App.Path & "\" & Nombre_ArchivoP & ".cat")) > 0 Then
            Kill App.Path & "\" & Nombre_ArchivoP & ".DBF"
        End If
        
        '****************Empaqueta el archivo con el pak.exe
        Ejecucion = "pak.exe a " & Nombre_ArchivoP & ".pak" & " " & Nombre_ArchivoP & ".cat"
        hProcess = OpenProcess(SYNCHRONIZE, 0, Shell("cmd /c, cd " & App.Path & " & " & Ejecucion))
        'Indica si se termino de procesar la información para poder continuar con las siguientes
        'ejecuciones
        If hProcess Then
            WaitForSingleObject hProcess, INFINITE
            CloseHandle hProcess
        End If
        'Si lo creo entonces borra el archivo .pak
        If Len(Dir$(App.Path & "\" & Nombre_ArchivoP & ".pak")) > 0 Then
            Kill App.Path & "\" & Nombre_ArchivoP & ".cat"
            MsgBox "El archivo de exportacion se genero correctamente", vbInformation + vbOKOnly
        End If
    Else
        MsgBox "El archivo de exportacion no se genero correctamente, verifique", vbInformation + vbOKOnly
        Exit Sub
    End If
    Me.MousePointer = 0
Exit Sub
handler:
    Me.MousePointer = 0
    MsgBox Err.Description
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Layout_Carga_Clientes_Click
'DESCRIPCION            : Carga el archivo de excel que contiene los Clientes
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 07 - Noviembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''''
Private Sub Btn_Layout_Carga_Clientes_Click()
Dim Clave As String
Dim Nombre As String
Dim Descripcion As String
Dim Presentacion As String
Dim SAL As String
Dim Laboratorio As String
Dim Costo_Maesba As String
Dim Precio_Max As String
Dim Precio_Publico As String
Dim Aplica_IVA As String
Dim Categoria As String
Dim Marca As String
Dim Proveedor As String
Dim Utilidad As String
Dim NIVEL As String
Dim Especialidad As String

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
    Grid_Clientes.Rows = 0
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
            .RecordSource = "CLIENTES$"
            Grid_Clientes.Redraw = False
            .Refresh
            Grid_Clientes.Redraw = True
        End With
        Grid_Clientes.Rows = 0
        Grid_Clientes.Cols = 4
        Cont_Partidas = 0

        Nombre = ""
        Ciudad = ""
        Direccion = ""
        Colonia = ""
        

        'Se agrega encabezado
        Grid_Clientes.AddItem "Nombre" & Chr(9) & "Ciudad" & Chr(9) & "Direccion" & Chr(9) & "Colonia"
        With Dt_Excel.Recordset
            While Not .EOF
                    If Not IsNull(Dt_Excel.Recordset(1).Value) Then Nombre = Dt_Excel.Recordset(1).Value
                    If Not IsNull(Dt_Excel.Recordset(2).Value) Then Ciudad = Dt_Excel.Recordset(2).Value
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then Direccion = Dt_Excel.Recordset(3).Value
                    If Not IsNull(Dt_Excel.Recordset(4).Value) Then Colonia = Dt_Excel.Recordset(4).Value
                    Grid_Clientes.AddItem Nombre & " " & Ciudad & Chr(9) & Ciudad & Chr(9) & Direccion & Chr(9) & Colonia
                    Nombre = ""
                    Ciudad = ""
                    Direccion = ""
                    Colonia = ""
                .MoveNext
            Wend
            
            'Configura el tamaño las columnas del Grid
            If Grid_Clientes.Rows > 1 Then
                Grid_Clientes.FixedRows = 1
                Grid_Clientes.ColWidth(0) = 1500 '
                Grid_Clientes.ColWidth(1) = 5000 '
                Grid_Clientes.ColWidth(2) = 2000 '
                Grid_Clientes.ColWidth(3) = 1000 '
                'Pone el setfocus en la primera fila del Grid
                With Grid_Clientes
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
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
'NOMBRE DE LA FUNCION   : Btn_Layout_Proveedores_Click
'DESCRIPCION            : Carga el archivo de excel que contiene los Proveedores
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 07 - Noviembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************''''''
Private Sub Btn_Layout_Proveedores_Click()
Dim Clave As String
Dim Nombre As String
Dim Descripcion As String
Dim Presentacion As String
Dim SAL As String
Dim Laboratorio As String
Dim Costo_Maesba As String
Dim Precio_Max As String
Dim Precio_Publico As String
Dim Aplica_IVA As String
Dim Categoria As String
Dim Marca As String
Dim Proveedor As String
Dim Utilidad As String
Dim NIVEL As String
Dim Especialidad As String

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
    Grid_Proveedores.Rows = 0
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
            .RecordSource = "PROVEEDORES$"
            Grid_Proveedores.Redraw = False
            .Refresh
            Grid_Proveedores.Redraw = True
        End With
        Grid_Proveedores.Rows = 0
        Grid_Proveedores.Cols = 6
        Cont_Partidas = 0

        Nombre = ""
        Direccion = ""
        Colonia = ""
        Ciudad = ""
        RFC = ""
        Telefono = ""
        

        'Se agrega encabezado
        Grid_Proveedores.AddItem "Nombre" & Chr(9) & "Direccion" & Chr(9) & "Colonia" & Chr(9) & "Ciudad" & Chr(9) & "RFC" & Chr(9) & "Telefono"
        With Dt_Excel.Recordset
            While Not .EOF
                    If Not IsNull(Dt_Excel.Recordset(1).Value) Then Nombre = Dt_Excel.Recordset(1).Value
                    If Not IsNull(Dt_Excel.Recordset(2).Value) Then Direccion = Dt_Excel.Recordset(2).Value
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then Colonia = Dt_Excel.Recordset(3).Value
                    If Not IsNull(Dt_Excel.Recordset(4).Value) Then Ciudad = Dt_Excel.Recordset(4).Value
                    If Not IsNull(Dt_Excel.Recordset(5).Value) Then RFC = Dt_Excel.Recordset(5).Value
                    If Not IsNull(Dt_Excel.Recordset(6).Value) Then Telefono = Dt_Excel.Recordset(6).Value
                    Grid_Proveedores.AddItem Nombre & Chr(9) & Direccion & Chr(9) & Colonia & Chr(9) & Ciudad & Chr(9) & RFC & Chr(9) & Telefono
                    Nombre = ""
                    Direccion = ""
                    Colonia = ""
                    Ciudad = ""
                    RFC = ""
                    Telefono = ""
                .MoveNext
            Wend
            
            'Configura el tamaño las columnas del Grid
            If Grid_Proveedores.Rows > 1 Then
                Grid_Proveedores.FixedRows = 1
                Grid_Proveedores.ColWidth(0) = 1500 '
                Grid_Proveedores.ColWidth(1) = 5000 '
                Grid_Proveedores.ColWidth(2) = 2000 '
                Grid_Proveedores.ColWidth(3) = 1000 '
                Grid_Proveedores.ColWidth(4) = 1000 '
                Grid_Proveedores.ColWidth(5) = 1000 '
                'Pone el setfocus en la primera fila del Grid
                With Grid_Proveedores
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
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
'NOMBRE DE LA FUNCION   : Btn_Loyaut_Toma_Inventario_Click
'DESCRIPCION            : Carga el archivo de excel que contiene los productos
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13 - Diciembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************'''''
Private Sub Btn_Loyaut_Toma_Inventario_Click()
Dim Clave As String
Dim Nombre As String
Dim Descripcion As String
Dim Presentacion As String
Dim SAL As String
Dim Laboratorio As String
Dim Costo_Maesba As String
Dim Precio_Max As String
Dim Precio_Publico As String
Dim Aplica_IVA As String
Dim Categoria As String
Dim Marca As String
Dim Proveedor As String
Dim Utilidad As String
Dim NIVEL As String
Dim Especialidad As String
Dim Cantidad As String
Dim LOTE As String
Dim Fecha_Caducidad As String
Dim Cont_Partidas As Integer
Dim path_XLS As String


    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Grid_Cat_Productos.Rows = 0
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
            Grid_Cat_Productos.Redraw = False
            .Refresh
            Grid_Cat_Productos.Redraw = True
        End With
        Grid_Cat_Productos.Rows = 0
        Grid_Cat_Productos.Cols = 19
        
        Cont_Partidas = 0
        Clave = ""
        Nombre = ""
        Descripcion = ""
        Presentacion = ""
        SAL = ""
        Laboratorio = ""
        Costo_Maesba = ""
        Precio_Max = ""
        Precio_Publico = ""
        Aplica_IVA = ""
        Categoria = ""
        Marca = ""
        Proveedor = ""
        Utilidad = ""
        NIVEL = ""
        Especialidad = ""
        Cantidad = ""
        LOTE = ""
        Fecha_Caducidad = ""
        
        'Se agrega encabezado
        Grid_Cat_Productos.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Descripcion" & Chr(9) & "Presentacion" & Chr(9) & "SAL" _
        & Chr(9) & "Laboratorio" & Chr(9) & "Costo_Maesba" & Chr(9) & "Precio_Max" & Chr(9) & "Precio_Publico" & Chr(9) & "Aplica_IVA" _
        & Chr(9) & "Categoria" & Chr(9) & "Marca" & Chr(9) & "Proveedor" & Chr(9) & "Utilidad" _
        & Chr(9) & "NIVEL" & Chr(9) & "Especialidad" & Chr(9) & "Cantidad" & Chr(9) & "LOTE" & Chr(9) & "Fecha_Caducidad"
        With Dt_Excel.Recordset
            While Not .EOF
            
                    If Not IsNull(Dt_Excel.Recordset(0).Value) Then Clave = Dt_Excel.Recordset(0).Value
                    If Not IsNull(Dt_Excel.Recordset(1).Value) Then Nombre = Dt_Excel.Recordset(1).Value
                    If Not IsNull(Dt_Excel.Recordset(2).Value) Then Descripcion = Dt_Excel.Recordset(2).Value
                    If Not IsNull(Dt_Excel.Recordset(3).Value) Then Presentacion = Dt_Excel.Recordset(3).Value
                    If Not IsNull(Dt_Excel.Recordset(4).Value) Then SAL = Dt_Excel.Recordset(4).Value
                    If Not IsNull(Dt_Excel.Recordset(5).Value) Then Laboratorio = Dt_Excel.Recordset(5).Value
                    If Not IsNull(Dt_Excel.Recordset(6).Value) Then Costo_Maesba = Dt_Excel.Recordset(6).Value
                    If Not IsNull(Dt_Excel.Recordset(7).Value) Then Precio_Max = Dt_Excel.Recordset(7).Value
                    If Not IsNull(Dt_Excel.Recordset(8).Value) Then Precio_Publico = Dt_Excel.Recordset(8).Value
                    If Not IsNull(Dt_Excel.Recordset(9).Value) Then Aplica_IVA = Dt_Excel.Recordset(9).Value
                    If Not IsNull(Dt_Excel.Recordset(10).Value) Then Categoria = Dt_Excel.Recordset(10).Value
                    If Not IsNull(Dt_Excel.Recordset(11).Value) Then Marca = Dt_Excel.Recordset(11).Value
                    If Not IsNull(Dt_Excel.Recordset(12).Value) Then Proveedor = Dt_Excel.Recordset(12).Value
                    If Not IsNull(Dt_Excel.Recordset(13).Value) Then Utilidad = Dt_Excel.Recordset(13).Value
                    If Not IsNull(Dt_Excel.Recordset(14).Value) Then NIVEL = Dt_Excel.Recordset(14).Value
                    If Not IsNull(Dt_Excel.Recordset(15).Value) Then Especialidad = Dt_Excel.Recordset(15).Value
                    If Not IsNull(Dt_Excel.Recordset(16).Value) Then Cantidad = Dt_Excel.Recordset(16).Value
                    If Not IsNull(Dt_Excel.Recordset(17).Value) Then LOTE = Dt_Excel.Recordset(17).Value
                    If Not IsNull(Dt_Excel.Recordset(18).Value) Then Fecha_Caducidad = Dt_Excel.Recordset(18).Value
                    
                    Grid_Cat_Productos.AddItem Clave & Chr(9) & Nombre & Chr(9) & Descripcion & Chr(9) & _
                    Presentacion & Chr(9) & SAL & Chr(9) & Laboratorio & Chr(9) & Costo_Maesba & Chr(9) & _
                    Precio_Max & Chr(9) & Precio_Publico & Chr(9) & Aplica_IVA & Chr(9) & Categoria & Chr(9) & Marca & Chr(9) & Proveedor & Chr(9) & Utilidad & Chr(9) & NIVEL & Chr(9) & Especialidad _
                    & Chr(9) & Cantidad & Chr(9) & LOTE & Chr(9) & Fecha_Caducidad
                    Cont_Partidas = Cont_Partidas + 1
                    Cont_Partidas = 0
                    
                    Clave = ""
                    Nombre = ""
                    Descripcion = ""
                    Presentacion = ""
                    SAL = ""
                    Laboratorio = ""
                    Costo_Maesba = ""
                    Precio_Max = ""
                    Precio_Publico = ""
                    Aplica_IVA = ""
                    Categoria = ""
                    Marca = ""
                    Proveedor = ""
                    Utilidad = ""
                    NIVEL = ""
                    Especialidad = ""
                    Cantidad = ""
                    LOTE = ""
                    Fecha_Caducidad = ""
                    
                .MoveNext
            Wend
            Txt_Producto_ID.text = Val(Cont_Partidas)
            
            'Configura el tamaño las columnas del Grid
            If Grid_Cat_Productos.Rows > 1 Then
                Grid_Cat_Productos.FixedRows = 1
                Grid_Cat_Productos.ColWidth(0) = 1500 '
                Grid_Cat_Productos.ColWidth(1) = 5000 '
                Grid_Cat_Productos.ColWidth(2) = 2000 '
                Grid_Cat_Productos.ColWidth(3) = 1000 '
                Grid_Cat_Productos.ColWidth(4) = 1000 '
                Grid_Cat_Productos.ColWidth(5) = 1000 '
                Grid_Cat_Productos.ColWidth(6) = 1000 '
                Grid_Cat_Productos.ColWidth(7) = 1000 '
                Grid_Cat_Productos.ColWidth(8) = 1000 '
                Grid_Cat_Productos.ColWidth(9) = 1000 '
                Grid_Cat_Productos.ColWidth(10) = 1000 '
                Grid_Cat_Productos.ColWidth(11) = 1000 '
                Grid_Cat_Productos.ColWidth(12) = 1000 '
                Grid_Cat_Productos.ColWidth(13) = 1000 '
                Grid_Cat_Productos.ColWidth(14) = 1000 '
                Grid_Cat_Productos.ColWidth(15) = 1000 '
                Grid_Cat_Productos.ColWidth(16) = 1000 '
                Grid_Cat_Productos.ColWidth(17) = 1000 '
                Grid_Cat_Productos.ColWidth(18) = 1000 '
                'Pone el setfocus en la primera fila del Grid
                With Grid_Cat_Productos
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
        End With
        Exit Sub
handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub

'Botón para modificar un registro del catálogo seleccionado
Private Sub Btn_Modificar_Click()
    
    Select Case Catalogo
        'Catálogo de clientes
        Case "CLIENTES":
            '1. Habiltar los Frames para poder modificar los datos del usuario
            '2. Modifica los datos del cliente
            If Btn_Modificar.Caption = "Modificar" Then
                If Trim(Txt_Cliente_ID.text) <> "" Then
                    Fra_Datos_Generales_clientes.Enabled = True
                    Fra_Datos_Factura.Enabled = True
                    Fra_Remisiones.Enabled = True
                    Fra_Clientes.Enabled = False
                    Tab_Clientes.Tab = 0
                    Btn_Nuevo.Enabled = False
                    Btn_Consultar.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                    Btn_Salir.Caption = "Cancelar"
                    Txt_Nombre_Cliente.SetFocus
                Else
                    MsgBox "Debe seleccionar un cliente" & Chr(13) & Chr(13) & _
                           "para poder modificar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
            Else
                'Actualizar los datos del cliente
                If Trim(Txt_Nombre_Cliente.text) <> "" Then
                    If Txt_Cuenta_Pago.text <> "" Then
                        If Trim(Len(Txt_Cuenta_Pago.text)) < 4 Then
                            MsgBox "Debe indicar la menos los últimos 4 dígitos de l npumero de cuenta de pago", vbExclamation
                            Txt_Cuenta_Pago.SetFocus
                            Exit Sub
                        End If
                    End If
                    If Cmb_Tipo_Persona.ListIndex = -1 Then
                        MsgBox "Debe indicar el tipo de persona", vbExclamation
                        Exit Sub
                    End If
                    Modifica_Cliente
                Else
                    MsgBox "Faltan datos para actualizar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    Txt_Nombre_Cliente.SetFocus
                End If
            End If
            
        'Catálogo de Proveedores
        Case "PROVEEDORES":
            '1. Habiltar los Frames para poder modificar los datos del Proveedor
            '2. Modifica los datos del Proveedor
            If Btn_Modificar.Caption = "Modificar" Then
                If Trim(Txt_Proveedor_ID.text) <> "" Then
                    Fra_Generales_Proveedores.Enabled = True
                    Fra_Proveedores.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Consultar.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                    Btn_Salir.Caption = "Cancelar"
                    Txt_Nombre_Proveedor.SetFocus
                Else
                    MsgBox "Debe seleccionar un Proveedor" & Chr(13) & Chr(13) & _
                           "para poder modificar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
            Else
                'Actualizar los datos del Proveedor
                If Trim(Txt_Nombre_Proveedor.text) <> "" And Trim(Txt_PRFC_Proveedores.text) <> "" And _
                Trim(Txt_Direccion_Proveedor.text) <> "" And Trim(Txt_colonia_Proveedor.text) <> "" _
                And Cmb_Estatus_Proveedor.ListIndex > -1 Then
                    Modifica_Proveedor
                Else
                    MsgBox "Faltan datos para actualizar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
            End If
            
        'Catálogo de Productos
        Case "PRODUCTOS":
            '1. Habiltar los Frames para poder modificar los datos del Producto
            '2. Modifica los datos del producto
            If Btn_Modificar.Caption = "Modificar" Then
                If Trim(Txt_Producto_ID.text) <> "" Then
                    Fra_Generales_Productos.Enabled = True
                    Fra_Comentario.Enabled = True
                    Fra_Almacen_Cat_Productos.Enabled = True
                    Fra_Comentario.Enabled = True
                    Fra_Costos_Cat_Productos.Enabled = True
                    Fra_Almacen_Cat_Productos.Enabled = True
                    Fra_Costos_Cat_Productos.Enabled = True
                    Fra_Detalles_Productos.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Consultar.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                    Btn_Salir.Caption = "Cancelar"
                Else
                    MsgBox "Debe seleccionar un Producto" & Chr(13) & Chr(13) & _
                           "para poder modificar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
            Else
            'Actualizar los datos del Producto
            If Cmb_Estatus_Producto.ListIndex > -1 And _
            Trim(Txt_Nombre_Cat_Productos.text) <> "" And Cmb_Cat_Producto_Tipo.ListIndex > -1 And Cmb_Presentaciones_Cat_Productos.ListIndex > -1 And Cmb_Cat_Productos_Categorias.ListIndex > -1 Then
                    Modifica_Cat_productos
                Else
                    MsgBox "Faltan datos para actualizar el registro", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                End If
            End If
    End Select
End Sub

'Botón para agregar un registro al catálogo seleccionado
Private Sub Btn_Nuevo_Click()
Set Conectar_Ayudante = New Ayudante
    
    Select Case Catalogo
        'Catálogos de clientes
        Case "CLIENTES":
        '1. Habilita los Frames para poder introducir los datos
        If Btn_Nuevo.Caption = "Nuevo" Then
            Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Clientes)
            Txt_Cliente_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Clientes", "Cliente_ID"), "00000")
            Fra_Datos_Generales_clientes.Enabled = True
            Fra_Datos_Factura.Enabled = True
            Fra_Remisiones.Enabled = True
            Fra_Clientes.Enabled = False
            Tab_Clientes.Tab = 0
            Txt_Nombre_Cliente.SetFocus
            Cmb_Status_Cliente.ListIndex = 0
            Cmb_Status_Cliente.Enabled = False
            Btn_Nuevo.Caption = "Dar de Alta"
            Btn_Modificar.Enabled = False
            Btn_Eliminar.Enabled = False
            Btn_Consultar.Enabled = False
            Btn_Salir.Caption = "Cancelar"
            Cmb_Credoto_Flexible.ListIndex = 0
            Cmb_Remision_Con_Precios.ListIndex = 1
        Else
            'Alta de Clientes
            If Trim(Txt_Nombre_Cliente.text) <> "" Then
                 If Txt_Cuenta_Pago.text <> "" Then
                    If Trim(Len(Txt_Cuenta_Pago.text)) < 4 Then
                        MsgBox "Debe indicar la menos los últimos 4 dígitos de l npumero de cuenta de pago", vbExclamation
                        Txt_Cuenta_Pago.SetFocus
                        Exit Sub
                    End If
                End If
                If Cmb_Tipo_Persona.ListIndex = -1 Then
                    MsgBox "Debe indicar el tipo de persona al que pertenece", vbExclamation
                    Exit Sub
                End If
                Alta_Clientes
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Txt_Nombre_Cliente.SetFocus
            End If
        End If
        
        
        'Catálogos de Proveedores
        Case "PROVEEDORES":
        '1. Habilita los Frames para poder introducir los datos
        If Btn_Nuevo.Caption = "Nuevo" Then
            Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Clientes)
            Txt_Proveedor_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Proveedores", "Proveedor_ID"), "00000")
            Fra_Generales_Proveedores.Enabled = True
            Fra_Proveedores.Enabled = False
            Txt_Nombre_Proveedor.SetFocus
            Cmb_Estatus_Proveedor.ListIndex = 0
            Cmb_Estatus_Proveedor.Enabled = False
            Btn_Nuevo.Caption = "Dar de Alta"
            Btn_Modificar.Enabled = False
            Btn_Eliminar.Enabled = False
            Btn_Consultar.Enabled = False
            Btn_Salir.Caption = "Cancelar"
        Else
            'Alta de Proveedores
            If Trim(Txt_Nombre_Proveedor.text) <> "" And Trim(Txt_PRFC_Proveedores.text) <> "" And _
            Trim(Txt_Direccion_Proveedor.text) <> "" And Trim(Txt_colonia_Proveedor.text) <> "" And Trim(Cmb_Clasificacion_Proveedor.text) <> "" Then
                Alta_Proveedor
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        End If
        
        
        'Catálogos de Productos
        Case "PRODUCTOS":
        '1. Habilita los Frames para poder introducir los datos
        If Btn_Nuevo.Caption = "Nuevo" Then
            Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Clientes)
            Txt_Producto_ID.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos", "Producto_ID"), "00000")
            Fra_Generales_Productos.Enabled = True
            Fra_Comentario.Enabled = True
            Fra_Almacen_Cat_Productos.Enabled = True
            Fra_Comentario.Enabled = True
            Fra_Costos_Cat_Productos.Enabled = True
            Fra_Detalles_Productos.Enabled = False
            Txt_Nombre_Cat_Productos.SetFocus
            Cmb_Estatus_Producto.ListIndex = 0
            Cmb_Estatus_Producto.Enabled = False
            Btn_Nuevo.Caption = "Dar de Alta"
            Btn_Modificar.Enabled = False
            Btn_Eliminar.Enabled = False
            Btn_Consultar.Enabled = False
            Btn_Salir.Caption = "Cancelar"
        Else
            'Alta de Productos
            If Cmb_Estatus_Producto.ListIndex > -1 And _
            Trim(Txt_Nombre_Cat_Productos.text) <> "" And Cmb_Cat_Producto_Tipo.ListIndex > -1 And Cmb_Presentaciones_Cat_Productos.ListIndex > -1 And Cmb_Cat_Productos_Categorias.ListIndex > -1 Then
                Alta_Cat_Productos
            Else
                MsgBox "Faltan datos para dar de alta", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        End If
    End Select

End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Alta_Clientes
'DESCRIPCIÓN: Hace al alta de clientes en la tabla Cat_Clientes de la base
'             de datos Proyectos
'PARÁMETROS:
'CREO:
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Alta_Clientes()
Dim Rs_Alta_Cat_Clientes As rdoResultset 'Del manejo de registro

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de cliente
    Set Rs_Alta_Cat_Clientes = Conectar_Ayudante.Recordset_Agregar("Cat_Clientes")
    'Llena la tabla de Cat_Clientes con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Clientes
    .AddNew
        .rdoColumns("Cliente_ID") = Trim(Txt_Cliente_ID.text)
        If Cmb_Clasificacion_Clientes.ListIndex > -1 Then
            .rdoColumns("Clasificacion_ID") = Format(Cmb_Clasificacion_Clientes.ItemData(Cmb_Clasificacion_Clientes.ListIndex), "00000")
        End If
        .rdoColumns("Status") = Mid(Cmb_Status_Cliente.text, 1, 1)
        .rdoColumns("Nombre") = UCase(Txt_Nombre_Cliente.text)
        .rdoColumns("RFC") = UCase(Txt_RFC.text)
        .rdoColumns("Dias_Credito") = Val(Txt_Dias_Credito.text)
        .rdoColumns("Direccion") = UCase(Txt_Direccion_Cliente.text)
        .rdoColumns("No_Ext") = Trim(UCase(Txt_Cliente_No_Ext.text))
        .rdoColumns("No_Int") = Trim(UCase(Txt_Cliente_No_Int.text))
        .rdoColumns("Colonia") = UCase(Txt_Colonia_Cliente.text)
        .rdoColumns("CP") = Txt_Codigo_Postal_Cliente.text
        .rdoColumns("Ciudad") = UCase(Txt_Ciudad_Cliente.text)
        .rdoColumns("Estado") = UCase(Txt_Estado_Cliente.text)
        .rdoColumns("Pais") = Trim(UCase(Txt_Cliente_Pais.text))
        .rdoColumns("Telefono") = Txt_Telefono_Cliente.text
        .rdoColumns("Celular") = Txt_Celular_Cliente.text
        .rdoColumns("Fax") = Txt_Fax_Cliente.text
        .rdoColumns("Email") = Txt_E_Mail_Cliente.text
        .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Cliente.text)
        .rdoColumns("Credito_Flexible") = Cmb_Credoto_Flexible.text
        .rdoColumns("Usuario_Creo") = Nombre_Usuario
        .rdoColumns("Fecha_Creo") = Now
        .rdoColumns("Remision_Con_Presio") = Cmb_Remision_Con_Precios.text
        .rdoColumns("Almacen") = Txt_Almacen.text
        .rdoColumns("Tipo_Persona") = Cmb_Tipo_Persona.text
        'Datos  remision
        .rdoColumns("Direccion_Remision") = UCase(Txt_Direccion_Remision.text)
        .rdoColumns("Colonia_Remision") = UCase(Txt_Colonia_Remision.text)
        .rdoColumns("CP_Remision") = Txt_CP_Remision.text
        .rdoColumns("Ciudad_Remision") = UCase(Txt_Ciudad_Remision.text)
        .rdoColumns("Estado_Remision") = UCase(Txt_Estado_Remision.text)
        .rdoColumns("Metodo_Pago") = Trim(Cmb_Metodo_Pago.text)
        .rdoColumns("Cuenta_Pago") = Trim(Txt_Cuenta_Pago.text)
    .Update
    End With
    Conexion_Base.CommitTrans
    Fra_Datos_Generales_clientes.Enabled = False
    Fra_Datos_Factura.Enabled = False
    Fra_Remisiones.Enabled = False
    Fra_Clientes.Enabled = True
    Tab_Clientes.Tab = 0
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Consultar.Enabled = True
    Cmb_Status_Cliente.Enabled = True
    Btn_Salir.Caption = "Salir"
    'Llena el grid
    If Grid_Clientes.Rows = 0 Then
        Grid_Clientes.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "RFC"
        Grid_Clientes.ColWidth(0) = 1000 'Cliente_ID
        Grid_Clientes.ColWidth(1) = 6300 'Nombre
        Grid_Clientes.ColWidth(2) = 1400 'RFC
    End If
    'Agrega en el Grid los datos contenidos de las cajas de texto
    Grid_Clientes.AddItem UCase(Txt_Cliente_ID.text) & Chr(9) & UCase(Txt_Nombre_Cliente.text) & Chr(9) & UCase(Txt_RFC.text)
    Grid_Clientes.FixedRows = 1
    Grid_Clientes.FixedCols = 1
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clientes", Frm_Cat_Clientes)
    MsgBox "Cliente dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
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
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Modificar.Caption = "Modificar"
        Select Case Catalogo
            Case "CLIENTES":
                Fra_Datos_Generales_clientes.Enabled = False
                Fra_Datos_Factura.Enabled = False
                Fra_Remisiones.Enabled = False
                Fra_Clientes.Enabled = True
                Cmb_Status_Cliente.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clientes", Frm_Cat_Clientes)
            Case "PROVEEDORES":
                Fra_Generales_Proveedores.Enabled = False
                Fra_Proveedores.Enabled = True
                Cmb_Estatus_Proveedor.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Proveedores", Frm_Cat_Clientes)
            Case "PRODUCTOS":
                Fra_Generales_Productos.Enabled = False
                Fra_Comentario.Enabled = False
                Fra_Almacen_Cat_Productos.Enabled = False
                Fra_Comentario.Enabled = False
                Fra_Costos_Cat_Productos.Enabled = False
                Fra_Detalles_Productos.Enabled = True
                Cmb_Estatus_Producto.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos", Frm_Cat_Clientes)
        End Select
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Clientes
    'DESCRIPCIÓN: Modifica los datos existentes del cliente en la tabla
    'PARÁMETROS:
    'CREO:
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Modifica_Cliente()
Dim Rs_Modifica_Cat_Clientes As rdoResultset  'Manejo de registro

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Selecciona los campos de la tabla de Cat_Cliente para ser modificados
    Mi_SQL = "SELECT * FROM Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Trim(Txt_Cliente_ID.text) & "'"
    Set Rs_Modifica_Cat_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Sustituye los datos que hay en la tabla Cat_Cliente por los datos existentes en las cajas de texto
    If Not Rs_Modifica_Cat_Clientes.EOF Then
        With Rs_Modifica_Cat_Clientes
            .Edit
                If Cmb_Clasificacion_Clientes.ListIndex > -1 Then
                    .rdoColumns("Clasificacion_ID") = Format(Cmb_Clasificacion_Clientes.ItemData(Cmb_Clasificacion_Clientes.ListIndex), "00000")
                End If
                .rdoColumns("Status") = Mid(Cmb_Status_Cliente.text, 1, 1)
                .rdoColumns("Nombre") = UCase(Txt_Nombre_Cliente.text)
                .rdoColumns("RFC") = UCase(Txt_RFC.text)
                .rdoColumns("Dias_Credito") = Val(Txt_Dias_Credito.text)
                .rdoColumns("Direccion") = UCase(Txt_Direccion_Cliente.text)
                .rdoColumns("No_Ext") = Trim(UCase(Txt_Cliente_No_Ext.text))
                .rdoColumns("No_Int") = Trim(UCase(Txt_Cliente_No_Int.text))
                .rdoColumns("Colonia") = UCase(Txt_Colonia_Cliente.text)
                .rdoColumns("CP") = Txt_Codigo_Postal_Cliente.text
                .rdoColumns("Ciudad") = UCase(Txt_Ciudad_Cliente.text)
                .rdoColumns("Estado") = UCase(Txt_Estado_Cliente.text)
                .rdoColumns("Pais") = Trim(UCase(Txt_Cliente_Pais.text))
                .rdoColumns("Telefono") = Txt_Telefono_Cliente.text
                .rdoColumns("Celular") = Txt_Celular_Cliente.text
                .rdoColumns("Fax") = Txt_Fax_Cliente.text
                .rdoColumns("Email") = Txt_E_Mail_Cliente.text
                .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Cliente.text)
                .rdoColumns("Credito_Flexible") = Cmb_Credoto_Flexible.text
                .rdoColumns("Usuario_Modifico") = Usuario
                .rdoColumns("Fecha_Modifico") = Now
                .rdoColumns("Remision_Con_Presio") = Cmb_Remision_Con_Precios.text
                .rdoColumns("Almacen") = Txt_Almacen.text
                .rdoColumns("Tipo_Persona") = Cmb_Tipo_Persona.text
                'Datos  remision
                .rdoColumns("Direccion_Remision") = UCase(Txt_Direccion_Remision.text)
                .rdoColumns("Colonia_Remision") = UCase(Txt_Colonia_Remision.text)
                .rdoColumns("CP_Remision") = Txt_CP_Remision.text
                .rdoColumns("Ciudad_Remision") = UCase(Txt_Ciudad_Remision.text)
                .rdoColumns("Estado_Remision") = UCase(Txt_Estado_Remision.text)
                .rdoColumns("Metodo_Pago") = Trim(Cmb_Metodo_Pago.text)
                .rdoColumns("Cuenta_Pago") = Trim(Txt_Cuenta_Pago.text)
            .Update
        End With
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Eliminar.Enabled = True
        Btn_Consultar.Enabled = True
        Fra_Datos_Generales_clientes.Enabled = False
        Fra_Datos_Factura.Enabled = False
        Fra_Remisiones.Enabled = False
        Fra_Clientes.Enabled = True
        Tab_Clientes.Tab = 0
        Grid_Clientes.TextMatrix(Grid_Clientes.RowSel, 1) = UCase(Txt_Nombre_Cliente.text)
        Grid_Clientes.TextMatrix(Grid_Clientes.RowSel, 2) = Txt_RFC.text
        MsgBox "Cliente Modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        Conexion_Base.CommitTrans
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clientes", Frm_Cat_Clientes)
    Else
        MsgBox "Cliente Inexistente", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub



Private Sub Combo1_Change()

End Sub

Private Sub Chk_Clave_Producto_Click()
    If Chk_Clave_Producto.Value = 1 Then Chk_Descripcion_Producto.Value = 0
End Sub

Private Sub Chk_Descripcion_Producto_Click()
    If Chk_Descripcion_Producto.Value = 1 Then Chk_Clave_Producto.Value = 0
End Sub

Private Sub Cmb_Cat_Producto_Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_ID,Nombre", "Cat_Productos_Tipo", Frm_Cat_Clientes.Cmb_Cat_Producto_Tipo, 1, " Estatus='ACTIVO' AND Nombre")
    End If
End Sub

Private Sub Cmb_Cat_Productos_Categorias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Categoria_ID,Nombre", "Cat_Categorias", Cmb_Cat_Productos_Categorias, 1, " Estatus='ACTIVO' AND Nombre")
    End If
End Sub

Private Sub Cmb_Metodo_Pago_Click()
    Txt_Cuenta_Pago.text = ""
    Txt_Cuenta_Pago.Locked = True
    Txt_Cuenta_Pago.TabStop = False
    Txt_Cuenta_Pago.Appearance = 0
    If Cmb_Metodo_Pago.ListIndex > -1 Then
        If Cmb_Metodo_Pago.ListIndex > 1 Then
            Txt_Cuenta_Pago.Locked = False
            Txt_Cuenta_Pago.TabStop = True
            Txt_Cuenta_Pago.Appearance = 1
        End If
    End If
End Sub

Private Sub Cmb_Presentaciones_Cat_Productos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Presentacion_ID,Nombre", "Cat_Presentaciones", Frm_Cat_Clientes.Cmb_Presentaciones_Cat_Productos, 1, " Estatus='ACTIVO' AND Nombre")
    End If
End Sub

Private Sub Cmb_Seleccionar_Impuesto_Cat_Productos_Click()
    Txt_IVA.text = Format(Val(Txt_Costo_Cat_Productos.text) * (Val(Cmb_Seleccionar_Impuesto_Cat_Productos.text) / 100), "##.00")
    Txt_Costo_Con_IVA = Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)
    Call Txt_Utilidad_Cat_Productos_Change
End Sub

Private Sub Form_Initialize()
    Set Mi_Ayudante = New Ayudante
    Set Mi_Ayudante.Forma = Me
End Sub

Private Sub Form_Load()
    'Medias de la Forma para que no puedan ser modificadas
    Me.Width = 10050 '9795
    Me.Height = 7950 '7200
    Frm_Cat_Clientes.Top = 0
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Resize()
    Mi_Ayudante.Redimensionar_Controles
End Sub

Private Sub Grid_Cat_Productos_Click()
Dim Rs_Consulta_Cat_Productos As rdoResultset
Dim Estatus As String
Dim Rs_Consulta_Impuesto As rdoResultset

    If Grid_Cat_Productos.Rows > 1 Then
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    'Selecciona los campos de la tabla de Cat_Proveedores
        Mi_SQL = "SELECT * FROM Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Producto_ID ='" & Grid_Cat_Productos.TextMatrix(Grid_Cat_Productos.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acurdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Productos.EOF Then
            With Rs_Consulta_Cat_Productos
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Comentarios_Cat_Productos.text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("Producto_ID")) Then Txt_Producto_ID.text = .rdoColumns("Producto_ID")
                If Not IsNull(.rdoColumns("Presentacion_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Presentacion_ID")), Cmb_Presentaciones_Cat_Productos)
                If Not IsNull(.rdoColumns("Tipo_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Tipo_ID")), Cmb_Cat_Producto_Tipo)
                If Not IsNull(.rdoColumns("Estatus")) Then Cmb_Estatus_Producto.text = .rdoColumns("Estatus")
                If Not IsNull(.rdoColumns("Clave")) Then Txt_Clave_Cat_Productos.text = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Costo")) Then Txt_Costo_Cat_Productos.text = .rdoColumns("Costo")
                If Not IsNull(.rdoColumns("Existencia")) Then Txt_Existencia.text = .rdoColumns("Existencia")
                If Not IsNull(.rdoColumns("Precio_Venta")) Then Txt_Precio_Venta.text = .rdoColumns("Precio_Venta")
                If Not IsNull(.rdoColumns("Categoria_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Categoria_ID")), Cmb_Cat_Productos_Categorias)
                If Not IsNull(.rdoColumns("Nombre")) Then Txt_Nombre_Cat_Productos.text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Aplica_IVA")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Aplica_IVA")), Cmb_Aplica_IVA)
                If Not IsNull(.rdoColumns("Aplica_Caja")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Aplica_Caja")), Cmb_Cajas)
                If Not IsNull(.rdoColumns("Cantidad_Cajas")) Then Txt_Cantidad_Cajas.text = .rdoColumns("Cantidad_Cajas")
                Call Txt_Impuesto_Cat_Productos_Change
            End With
            Rs_Consulta_Cat_Productos.Close
        End If
    End If
End Sub
Private Sub Grid_Cat_Productos_EnterCell()
    Call Grid_Cat_Productos_Click
End Sub
'Llena los campos con los datos del registro del grid
Private Sub Grid_Clientes_Click()
Dim Rs_Consulta_Cat_Clientes As rdoResultset  'Manejo de registro
Dim Estatus As String                          'Obtiene el estatus del cliente consultado

    If Grid_Clientes.Rows > 1 Then
    'Selecciona los campos de la tabla de Cat_Cliente
        Mi_SQL = "SELECT * FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID ='" & Grid_Clientes.TextMatrix(Grid_Clientes.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        
        'Llena las cajas de texto de acurdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Clientes.EOF Then
            With Rs_Consulta_Cat_Clientes
                Txt_Cliente_ID.text = .rdoColumns("Cliente_ID")
                If .rdoColumns("Status") = "A" Then
                    Estatus = "ACTIVO"
                Else
                    Estatus = CANCELADO
                End If
                If Trim(.rdoColumns("Status")) = "C" Then
                    Cmb_Status_Cliente.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo("CANCELADO", Cmb_Status_Cliente)
                Else
                    If Trim(.rdoColumns("Status")) = "A" Then Cmb_Status_Cliente.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo("ACTIVO", Cmb_Status_Cliente)
                End If
                If Not IsNull(.rdoColumns("Clasificacion_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Clasificacion_ID")), Cmb_Clasificacion_Clientes)
                If Not IsNull(.rdoColumns("Nombre")) Then Txt_Nombre_Cliente.text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("RFC")) Then Txt_RFC.text = .rdoColumns("RFC")
                If Not IsNull(.rdoColumns("Dias_Credito")) Then Txt_Dias_Credito = .rdoColumns("Dias_Credito")
                If Not IsNull(.rdoColumns("Direccion")) Then Txt_Direccion_Cliente.text = .rdoColumns("Direccion")
                If Not IsNull(.rdoColumns("No_Ext")) Then
                    Txt_Cliente_No_Ext.text = .rdoColumns("No_Ext")
                Else
                    Txt_Cliente_No_Ext.text = ""
                End If
                If Not IsNull(.rdoColumns("No_Int")) Then
                    Txt_Cliente_No_Int.text = .rdoColumns("No_Int")
                Else
                    Txt_Cliente_No_Int.text = ""
                End If
                If Not IsNull(.rdoColumns("Colonia")) Then Txt_Colonia_Cliente.text = .rdoColumns("Colonia")
                If Not IsNull(.rdoColumns("CP")) Then Txt_Codigo_Postal_Cliente.text = .rdoColumns("CP")
                If Not IsNull(.rdoColumns("Ciudad")) Then Txt_Ciudad_Cliente.text = .rdoColumns("Ciudad")
                If Not IsNull(.rdoColumns("Estado")) Then Txt_Estado_Cliente.text = .rdoColumns("Estado")
                If Not IsNull(.rdoColumns("Pais")) Then
                    Txt_Cliente_Pais.text = .rdoColumns("Pais")
                Else
                    Txt_Cliente_Pais.text = ""
                End If
                If Not IsNull(.rdoColumns("Telefono")) Then Txt_Telefono_Cliente.text = .rdoColumns("Telefono")
                If Not IsNull(.rdoColumns("Celular")) Then Txt_Celular_Cliente.text = .rdoColumns("Celular")
                If Not IsNull(.rdoColumns("Fax")) Then Txt_Fax_Cliente.text = .rdoColumns("Fax")
                If Not IsNull(.rdoColumns("Email")) Then Txt_E_Mail_Cliente.text = .rdoColumns("Email")
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Comentarios_Cliente.text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("Tipo_Persona")) Then
                    Cmb_Tipo_Persona.text = .rdoColumns("Tipo_Persona")
                Else
                    Cmb_Tipo_Persona.ListIndex = -1
                End If
                Cmb_Credoto_Flexible.ListIndex = -1
                If Not IsNull(.rdoColumns(.rdoColumns("Credito_Flexible"))) Then
                    If Trim(.rdoColumns("Credito_Flexible")) = "SI" Then
                        Cmb_Credoto_Flexible.ListIndex = 0
                    Else
                        Cmb_Credoto_Flexible.ListIndex = 1
                    End If
                End If
                If Not IsNull(.rdoColumns("Remision_Con_Presio")) Then
                    If Trim(.rdoColumns("Remision_Con_Presio")) = "SI" Then
                        Cmb_Remision_Con_Precios.ListIndex = 0
                    Else
                        Cmb_Remision_Con_Precios.ListIndex = 1
                    End If
                Else
                    Cmb_Remision_Con_Precios.ListIndex = -1
                End If
                Txt_Almacen.text = ""
                Txt_Direccion_Remision.text = ""
                Txt_Colonia_Remision.text = ""
                Txt_CP_Remision.text = ""
                Txt_Ciudad_Remision.text = ""
                Txt_Estado_Remision.text = ""
                If Not IsNull(.rdoColumns("Almacen")) Then Txt_Almacen.text = .rdoColumns("Almacen")
                If Not IsNull(.rdoColumns("Direccion_Remision")) Then Txt_Direccion_Remision.text = .rdoColumns("Direccion_Remision")
                If Not IsNull(.rdoColumns("Colonia_Remision")) Then Txt_Colonia_Remision.text = .rdoColumns("Colonia_Remision")
                If Not IsNull(.rdoColumns("CP_Remision")) Then Txt_CP_Remision.text = .rdoColumns("CP_Remision")
                If Not IsNull(.rdoColumns("Ciudad_Remision")) Then Txt_Ciudad_Remision.text = .rdoColumns("Ciudad_Remision")
                If Not IsNull(.rdoColumns("Estado_Remision")) Then Txt_Estado_Remision.text = .rdoColumns("Estado_Remision")
                If Not IsNull(.rdoColumns("Metodo_Pago")) Then
                    Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Metodo_Pago"), Cmb_Metodo_Pago)
                Else
                    Cmb_Metodo_Pago.ListIndex = -1
                End If
                If Not IsNull(.rdoColumns("Cuenta_Pago")) Then Txt_Cuenta_Pago.text = .rdoColumns("Cuenta_Pago")
            End With
        End If
    End If
End Sub

Private Sub Grid_Clientes_EnterCell()
    Call Grid_Clientes_Click
End Sub

Private Sub Grid_Proveedores_Click()
Dim Rs_Consulta_Cat_Proveedores As rdoResultset
Dim Estatus As String

    If Grid_Proveedores.Rows > 1 Then
    'Selecciona los campos de la tabla de Cat_Proveedores
        Mi_SQL = "SELECT * FROM Cat_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID ='" & Grid_Proveedores.TextMatrix(Grid_Proveedores.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acurdo a los datos obtenidos de la consulta
        Txt_Dias_Credio_Proveedor.text = ""
        If Not Rs_Consulta_Cat_Proveedores.EOF Then
            With Rs_Consulta_Cat_Proveedores
                If Not IsNull(.rdoColumns("Proveedor_ID")) Then Txt_Proveedor_ID.text = .rdoColumns("Proveedor_ID")
                If Not IsNull(.rdoColumns("Estatus")) Then Cmb_Estatus_Proveedor.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Proveedor)
                If Not IsNull(.rdoColumns("Clasificacion_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Clasificacion_ID")), Cmb_Clasificacion_Proveedor)
                If Not IsNull(.rdoColumns("Nombre")) Then Txt_Nombre_Proveedor.text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Direccion")) Then Txt_Direccion_Proveedor.text = .rdoColumns("Direccion")
                If Not IsNull(.rdoColumns("Colonia")) Then Txt_colonia_Proveedor.text = .rdoColumns("Colonia")
                If Not IsNull(.rdoColumns("Ciudad")) Then Txt_Ciudad.text = .rdoColumns("Ciudad")
                If Not IsNull(.rdoColumns("Estado")) Then Txt_Estado_Proveedor.text = .rdoColumns("Estado")
                If Not IsNull(.rdoColumns("Correo_Electronico")) Then Txt_Correo_Proveedores.text = .rdoColumns("Correo_Electronico")
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_comentarios_proveedores.text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("RFC")) Then Txt_PRFC_Proveedores.text = .rdoColumns("RFC")
                If Not IsNull(.rdoColumns("Codigo_Postal")) Then Txt_CP_Proveedores.text = .rdoColumns("Codigo_Postal")
                If Not IsNull(.rdoColumns("Telefono")) Then Txt_Telefono_Proveedores.text = .rdoColumns("Telefono")
                If Not IsNull(.rdoColumns("Celular")) Then Txt_Celular_Proveedores.text = .rdoColumns("Celular")
                If Not IsNull(.rdoColumns("Fax")) Then Txt_Fax_Proveedores.text = .rdoColumns("Fax")
                If Not IsNull(.rdoColumns("Dias_Credito")) Then Txt_Dias_Credio_Proveedor.text = .rdoColumns("Dias_Credito")
                Cmd_Tipo_Pago.ListIndex = -1
                If Not IsNull(.rdoColumns("Tipo_Pago")) Then Cmd_Tipo_Pago.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Tipo_Pago"), Cmd_Tipo_Pago)
            End With
            Rs_Consulta_Cat_Proveedores.Close
        End If
    End If
End Sub

Private Sub Grid_Proveedores_EnterCell()
    Call Grid_Proveedores_Click
End Sub






Private Sub Text1_Change()

End Sub

Private Sub Txt_CAntidad_Caljas_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_CAntidad_Caljas.text, False)
End Sub



Private Sub Txt_Costo_Cat_Productos_Change()
''   If Val(Txt_Costo_Cat_Productos.Text) = 0 Then
''        Txt_Costo_Cat_Productos.Text = 0
''    End If
    Txt_IVA.text = Format(Val(Txt_Costo_Cat_Productos.text) * (Val(PG_Retencion_IVA)), "#.00")
    Txt_Costo_Con_IVA = Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)
    Txt_Precio_Venta = Format((Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)), "#,###,###.00")
    ''Call Txt_Utilidad_Cat_Productos_Change
End Sub

Private Sub Txt_Costo_Cat_Productos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Costo_Cat_Productos.text, True)
End Sub


Private Sub Txt_Dias_Credito_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Dias_Credito.text, False)
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Alta_Proveedor
'DESCRIPCIÓN            : Da de alta los Proveedores
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13 agosto 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Alta_Proveedor()
Dim Rs_Alta_Cat_Proveedores As rdoResultset

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de Proveedor
    Set Rs_Alta_Cat_Proveedores = Conectar_Ayudante.Recordset_Agregar("Cat_Proveedores")
    'Llena la tabla de Cat_Proveedores
    With Rs_Alta_Cat_Proveedores
    .AddNew
        .rdoColumns("Proveedor_ID") = Trim(Txt_Proveedor_ID.text)
        .rdoColumns("Clasificacion_ID") = Format(Cmb_Clasificacion_Proveedor.ItemData(Cmb_Clasificacion_Proveedor.ListIndex), "00000")
        .rdoColumns("Nombre") = UCase(Txt_Nombre_Proveedor.text)
        .rdoColumns("Direccion") = UCase(Txt_Direccion_Proveedor.text)
        .rdoColumns("Colonia") = UCase(Txt_colonia_Proveedor.text)
        .rdoColumns("Ciudad") = UCase(Txt_Ciudad.text)
        .rdoColumns("Estado") = UCase(Txt_Estado_Proveedor.text)
        .rdoColumns("Correo_Electronico") = Txt_Correo_Proveedores.text
        .rdoColumns("Comentarios") = UCase(Txt_comentarios_proveedores.text)
        .rdoColumns("Estatus") = UCase(Cmb_Estatus_Proveedor.text)
        .rdoColumns("RFC") = Txt_PRFC_Proveedores.text
        .rdoColumns("Codigo_Postal") = Txt_CP_Proveedores.text
        .rdoColumns("Telefono") = Txt_Telefono_Proveedores.text
        .rdoColumns("Celular") = Txt_Celular_Proveedores.text
        .rdoColumns("Fax") = Txt_Fax_Proveedores.text
        .rdoColumns("Dias_Credito") = Val(Txt_Dias_Credio_Proveedor.text)
        .rdoColumns("Usuario_Creo") = Nombre_Usuario
        .rdoColumns("Fecha_Creo") = Now
        .rdoColumns("Tipo_Pago") = Trim(Cmd_Tipo_Pago.text)
    .Update
    End With
    Rs_Alta_Cat_Proveedores.Close
    Conexion_Base.CommitTrans
    Fra_Generales_Proveedores.Enabled = False
    Fra_Proveedores.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Consultar.Enabled = True
    Cmb_Estatus_Proveedor.Enabled = True
    Btn_Salir.Caption = "Salir"
    'Llena el grid
    If Grid_Proveedores.Rows = 0 Then
        Grid_Proveedores.AddItem "Proveedor ID" & Chr(9) & "Nombre" & Chr(9) & "RFC"
        Grid_Proveedores.ColWidth(0) = 1000 'Proveedor_ID
        Grid_Proveedores.ColWidth(1) = 6300 'Nombre
        Grid_Proveedores.ColWidth(2) = 1400 'RFC
    End If
    'Agrega en el Grid los datos contenidos de las cajas de texto
    Grid_Proveedores.AddItem UCase(Txt_Proveedor_ID.text) & Chr(9) & UCase(Txt_Nombre_Proveedor.text) & Chr(9) & UCase(Txt_PRFC_Proveedores.text)
    Grid_Proveedores.FixedRows = 1
    Grid_Proveedores.FixedCols = 1
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Proveedores", Frm_Cat_Clientes)
    MsgBox "Proveedor dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Modifica_Proveedor
'DESCRIPCIÓN            : Modifica el Proveedor seleccionado
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13 Agosto de "2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Modifica_Proveedor()
Dim Rs_Modifica_Cat_Proveedores As rdoResultset

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Selecciona los campos de la tabla de Cat_Proveedores
    Mi_SQL = "SELECT * FROM Cat_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Trim(Txt_Proveedor_ID.text) & "'"
    Set Rs_Modifica_Cat_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Cat_Proveedores.EOF Then
        With Rs_Modifica_Cat_Proveedores
            .Edit
                .rdoColumns("Clasificacion_ID") = Format(Cmb_Clasificacion_Proveedor.ItemData(Cmb_Clasificacion_Proveedor.ListIndex), "00000")
                .rdoColumns("Nombre") = UCase(Txt_Nombre_Proveedor.text)
                .rdoColumns("Direccion") = UCase(Txt_Direccion_Proveedor.text)
                .rdoColumns("Colonia") = UCase(Txt_colonia_Proveedor.text)
                .rdoColumns("Ciudad") = UCase(Txt_Ciudad.text)
                .rdoColumns("Estado") = UCase(Txt_Estado_Proveedor.text)
                .rdoColumns("Correo_Electronico") = Txt_Correo_Proveedores.text
                .rdoColumns("Comentarios") = UCase(Txt_comentarios_proveedores.text)
                .rdoColumns("Estatus") = UCase(Cmb_Estatus_Proveedor.text)
                .rdoColumns("RFC") = Txt_PRFC_Proveedores.text
                .rdoColumns("Codigo_Postal") = Txt_CP_Proveedores.text
                .rdoColumns("Telefono") = Txt_Telefono_Proveedores.text
                .rdoColumns("Celular") = Txt_Celular_Proveedores.text
                .rdoColumns("Fax") = Txt_Fax_Proveedores.text
                .rdoColumns("Dias_Credito") = Txt_Dias_Credio_Proveedor.text
                .rdoColumns("Usuario_Modifico") = Usuario
                .rdoColumns("Fecha_Modifico") = Now
                .rdoColumns("Tipo_Pago") = Trim(Cmd_Tipo_Pago.text)
            .Update
        End With
        Rs_Modifica_Cat_Proveedores.Close
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Eliminar.Enabled = True
        Btn_Consultar.Enabled = True
        Fra_Generales_Proveedores.Enabled = False
        Fra_Proveedores.Enabled = True
        Grid_Proveedores.TextMatrix(Grid_Proveedores.RowSel, 1) = UCase(Txt_Nombre_Proveedor.text)
        Grid_Proveedores.TextMatrix(Grid_Proveedores.RowSel, 2) = Txt_PRFC_Proveedores.text
        MsgBox "Proveedor Modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        Conexion_Base.CommitTrans
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Proveedores", Frm_Cat_Clientes)
    Else
        MsgBox "Proveedor Inexistente", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Consulta_Proveedor
'DESCRIPCIÓN            : Consulta los proveedores segun el parametro
'PARÁMETROS             : Texto_Busqueda; parametro por el cual se va hacer la consulta del proveedor
'CREO                   : Julio Cruz
'FECHA_CREO             : 13 Agosto de 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Consulta_Proveedor(Texto_Busqueda As String)
Dim Rs_Consulta_Cat_Proveedores As rdoResultset
    
    Grid_Proveedores.Rows = 0
    'Consulta los clientes de acuerdo al parametro
    Mi_SQL = "SELECT Proveedor_ID, Nombre, RFC"
    Mi_SQL = Mi_SQL & " FROM Cat_Proveedores "
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Texto_Busqueda & "%'" & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'LLena el grid con los datos el resultado de la busqueda anterior
    If Not Rs_Consulta_Cat_Proveedores.EOF Then
        Grid_Proveedores.AddItem "Proveedor ID" & Chr(9) & "Nombre" & Chr(9) & "RFC"
        While Not Rs_Consulta_Cat_Proveedores.EOF
            Grid_Proveedores.AddItem Rs_Consulta_Cat_Proveedores.rdoColumns("Proveedor_ID") & Chr(9) & _
            Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre") & Chr(9) & _
            Rs_Consulta_Cat_Proveedores.rdoColumns("RFC")
            Grid_Proveedores.FixedRows = 1
            Grid_Proveedores.FixedCols = 1
            Rs_Consulta_Cat_Proveedores.MoveNext
        Wend
        Rs_Consulta_Cat_Proveedores.Close
        'Configura el grid
        Grid_Proveedores.ColWidth(0) = 1000
        Grid_Proveedores.ColWidth(1) = 6300
        Grid_Proveedores.ColWidth(2) = 1400
        'Manda llamara la función Grid_Proveedores_Click
        Grid_Proveedores_Click
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Alta_Cat_Productos
'DESCRIPCIÓN            : Hace al alta de Productos en la tabla Cat_Productos de la BD
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 27/Agosto/2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Alta_Cat_Productos()
Dim Rs_Alta_Cat_Productos As rdoResultset 'Del manejo de registro

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Alta de Productos
    Set Rs_Alta_Cat_Productos = Conectar_Ayudante.Recordset_Agregar("Cat_Productos")
    'Llena la tabla de Cat_Productos con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Productos
        .AddNew
            .rdoColumns("Producto_ID") = Trim(Txt_Producto_ID.text)
            If Cmb_Presentaciones_Cat_Productos.ListIndex > -1 Then .rdoColumns("Presentacion_ID") = Format(Cmb_Presentaciones_Cat_Productos.ItemData(Cmb_Presentaciones_Cat_Productos.ListIndex), "00000")
            If Cmb_Cat_Producto_Tipo.ListIndex > -1 Then .rdoColumns("Tipo_ID") = Format(Cmb_Cat_Producto_Tipo.ItemData(Cmb_Cat_Producto_Tipo.ListIndex), "00000")
            .rdoColumns("Estatus") = Cmb_Estatus_Producto.text
            .rdoColumns("Clave") = Txt_Clave_Cat_Productos.text
            .rdoColumns("Costo") = Val(Txt_Costo_Cat_Productos.text)
            .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Cat_Productos.text)
            .rdoColumns("Precio_Venta") = Val(Txt_Precio_Venta.text)
            If Cmb_Cat_Productos_Categorias.ListIndex > -1 Then .rdoColumns("Categoria_ID") = Format(Cmb_Cat_Productos_Categorias.ItemData(Cmb_Cat_Productos_Categorias.ListIndex), "00000")
            .rdoColumns("Nombre") = Trim(Txt_Nombre_Cat_Productos.text)
            .rdoColumns("Aplica_IVA") = Trim(Cmb_Aplica_IVA.text)
            .rdoColumns("Impuesto") = PG_Retencion_IVA
            .rdoColumns("Aplica_Caja") = Trim(Cmb_Cajas)
            .rdoColumns("Cantidad_Cajas") = Val(Txt_Cantidad_Cajas.text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Productos.Close
    Conexion_Base.CommitTrans
    Fra_Generales_Productos.Enabled = False
    Fra_Comentario.Enabled = False
    Fra_Almacen_Cat_Productos.Enabled = False
    Fra_Comentario.Enabled = False
    Fra_Costos_Cat_Productos.Enabled = False
    Fra_Detalles_Productos.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Consultar.Enabled = True
    Cmb_Estatus_Producto.Enabled = True
    Btn_Salir.Caption = "Salir"
    'Llena el grid
    If Grid_Cat_Productos.Rows = 0 Then
        Grid_Cat_Productos.AddItem "Producto ID" & Chr(9) & "Clave" & Chr(9) & "Nombre"
        Grid_Cat_Productos.ColWidth(0) = 1000
        Grid_Cat_Productos.ColAlignment(0) = 3
        Grid_Cat_Productos.ColWidth(1) = 1400
        Grid_Cat_Productos.ColAlignment(1) = 1
        Grid_Cat_Productos.ColWidth(2) = 6300
        Grid_Cat_Productos.ColAlignment(2) = 1
    End If
    'Agrega en el Grid los datos contenidos de las cajas de texto
    Grid_Cat_Productos.AddItem UCase(Txt_Producto_ID.text) & Chr(9) & UCase(Txt_Clave_Cat_Productos.text) & Chr(9) & UCase(Txt_Nombre_Cat_Productos.text)
    Grid_Cat_Productos.FixedRows = 1
    Grid_Cat_Productos.FixedCols = 1
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos", Frm_Cat_Clientes)
    MsgBox "Producto dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Cat_productos
'DESCRIPCIÓN: Consulta fragmentos del la Clave obtenidos de un InputBox dando
'             como resultado todas las Claves que empiesen igual
'PARÁMETROS:
'             1. Nombre: Texto_Busqueda. Es usada para buscar el nombre en
'                        la Base de Datos
'CREO               : Julio Cruz
'FECHA_CREO         : 27/Agosto/2010
'MODIFICO           :
'FECHA_MODIFICO     :
'CAUSA_MODIFICACIÓN :
'*******************************************************************************
Public Sub Consulta_Cat_Productos(Texto_Busqueda As String, Tipo_Busqueda As String)
Dim Rs_Consulta_Cat_Productos As rdoResultset    'Manejo de registro
    
    Grid_Cat_Productos.Rows = 0
    If Tipo_Busqueda = "CLAVE" Then
        'Consulta los Productos de acuerdo ala clave
        Mi_SQL = "SELECT Producto_ID, Descripcion,Clave,Nombre"
        Mi_SQL = Mi_SQL & " FROM Cat_Productos "
        Mi_SQL = Mi_SQL & " WHERE Clave LIKE '%" & Texto_Busqueda & "%'" & " ORDER BY Producto_ID"
    Else
        If Tipo_Busqueda = "DESCRIPCION" Then
            'Consulta los Productos de acuerdo ala Descripcion
            Mi_SQL = "SELECT Producto_ID, Descripcion,Clave,Nombre"
            Mi_SQL = Mi_SQL & " FROM Cat_Productos "
            Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Texto_Busqueda & "%'" & " ORDER BY Producto_ID"
        Else
            'Consulta los Productos de acuerdo al ID del producto
            Mi_SQL = "SELECT Producto_ID, Descripcion,Clave,Nombre"
            Mi_SQL = Mi_SQL & " FROM Cat_Productos "
            Mi_SQL = Mi_SQL & " WHERE Producto_ID= '" & Trim(Texto_Busqueda) & "' ORDER BY Producto_ID"
        End If
    End If
    Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'LLena el grid con los datos el resultado de la busqueda anterior
    If Not Rs_Consulta_Cat_Productos.EOF Then
        Grid_Cat_Productos.AddItem "Producto  ID" & Chr(9) & "Clave" & Chr(9) & "Nombre"
        While Not Rs_Consulta_Cat_Productos.EOF
            Grid_Cat_Productos.AddItem Rs_Consulta_Cat_Productos.rdoColumns("Producto_ID") & Chr(9) & _
            Rs_Consulta_Cat_Productos.rdoColumns("Clave") & Chr(9) & _
            Rs_Consulta_Cat_Productos.rdoColumns("Nombre")
            Grid_Cat_Productos.FixedRows = 1
            Rs_Consulta_Cat_Productos.MoveNext
        Wend
        Rs_Consulta_Cat_Productos.Close
        'Configura el grid
        Grid_Cat_Productos.ColWidth(0) = 1000
        Grid_Cat_Productos.ColAlignment(0) = 3
        Grid_Cat_Productos.ColWidth(1) = 1400
         Grid_Cat_Productos.ColAlignment(1) = 1
        Grid_Cat_Productos.ColWidth(2) = 6300
        Grid_Cat_Productos.ColAlignment(2) = 1
        'Manda llamara la función Grid_Cat_Productos_Click
        Grid_Cat_Productos_Click
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Modifica_Cat_productos
'DESCRIPCIÓN            : Modifica el Producto seleccionado
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 27 Agosto de 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Modifica_Cat_productos()
Dim Rs_Modifica_Cat_Productos As rdoResultset
Dim Rs_Consultar As rdoResultset
Dim Mi_SQL As String
Dim Clave_Temporal As String

On Error GoTo handler
    Conexion_Base.BeginTrans
    
    'Se consulta la clave para determinar si se modifico
    Mi_SQL = " SELECT * FROM Cat_Productos "
    Mi_SQL = Mi_SQL & " WHERE Producto_ID ='" & Format(Txt_Producto_ID.text, "00000") & "'"
    Set Rs_Consultar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consultar.EOF Then
        Clave_Temporal = Rs_Consultar!Clave
    End If
    Rs_Consultar.Close
    'Selecciona los campos de la tabla de Cat_Proveedores
    Mi_SQL = "SELECT * FROM Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Trim(Txt_Producto_ID.text) & "'"
    Set Rs_Modifica_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Cat_Productos.EOF Then
        With Rs_Modifica_Cat_Productos
            .Edit
                .rdoColumns("Producto_ID") = Trim(Txt_Producto_ID.text)
                If Cmb_Presentaciones_Cat_Productos.ListIndex > -1 Then .rdoColumns("Presentacion_ID") = Format(Cmb_Presentaciones_Cat_Productos.ItemData(Cmb_Presentaciones_Cat_Productos.ListIndex), "00000")
                If Cmb_Cat_Producto_Tipo.ListIndex > -1 Then .rdoColumns("Tipo_ID") = Format(Cmb_Cat_Producto_Tipo.ItemData(Cmb_Cat_Producto_Tipo.ListIndex), "00000")
                .rdoColumns("Estatus") = Cmb_Estatus_Producto.text
                .rdoColumns("Clave") = Txt_Clave_Cat_Productos.text
                .rdoColumns("Costo") = Val(Txt_Costo_Cat_Productos.text)
                .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Cat_Productos.text)
                .rdoColumns("Precio_Venta") = Val(Txt_Precio_Venta.text)
                If Cmb_Cat_Productos_Categorias.ListIndex > -1 Then .rdoColumns("Categoria_ID") = Format(Cmb_Cat_Productos_Categorias.ItemData(Cmb_Cat_Productos_Categorias.ListIndex), "00000")
                .rdoColumns("Nombre") = Trim(Txt_Nombre_Cat_Productos.text)
                .rdoColumns("Aplica_IVA") = Trim(Cmb_Aplica_IVA.text)
                .rdoColumns("Impuesto") = PG_Retencion_IVA
                .rdoColumns("Aplica_Caja") = Trim(Cmb_Cajas)
                .rdoColumns("Cantidad_Cajas") = Val(Txt_Cantidad_Cajas.text)
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
        Rs_Modifica_Cat_Productos.Close
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Eliminar.Enabled = True
        Btn_Consultar.Enabled = True
        Fra_Generales_Productos.Enabled = False
        Fra_Comentario.Enabled = False
        Fra_Almacen_Cat_Productos.Enabled = False
        Fra_Comentario.Enabled = False
        Fra_Costos_Cat_Productos.Enabled = False
        Fra_Detalles_Productos.Enabled = True
        Grid_Cat_Productos.TextMatrix(Grid_Cat_Productos.RowSel, 1) = UCase(Txt_Clave_Cat_Productos.text)
        Grid_Cat_Productos.TextMatrix(Grid_Cat_Productos.RowSel, 2) = Txt_Nombre_Cat_Productos.text
        MsgBox "Producto Modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        Conexion_Base.CommitTrans
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos", Frm_Cat_Clientes)
    Else
        MsgBox "Producto Inexistente", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub



Private Sub Txt_Impuesto_Cat_Productos_Change()
    Txt_IVA.text = Format(Val(Txt_Costo_Cat_Productos.text) * (Val(PG_Retencion_IVA)), "##.00")
    Txt_Costo_Con_IVA = Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)
    Call Txt_Utilidad_Cat_Productos_Change
End Sub

Private Sub Txt_Precio_Venta_Change()
Dim Utilidad As Long
    If Cambio = False Then
''        If (Val(Txt_Costo_Cat_Productos.Text) + Val(Txt_IVA.Text)) <> 0 Then
''            Utilidad = (Val(Txt_Precio_Venta.Text) - (Val(Txt_Costo_Cat_Productos.Text) + Val(Txt_IVA.Text))) * 100 / (Val(Txt_Costo_Cat_Productos.Text) + Val(Txt_IVA.Text))
''            Cambio = True
''            Cambio = False
''        End If
    End If
End Sub

Private Sub Txt_Prioridad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Prioridad.text, False)
End Sub


Private Sub Txt_Utilidad_Cat_Productos_Change()
Dim Precio_Venta As String
    If Cambio = False Then
        If (Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)) <> 0 Then
            Precio_Venta = Format((Val(Txt_Costo_Cat_Productos.text) + Val(Txt_IVA.text)), "#,###,###.00")
            Cambio = True
            Txt_Precio_Venta.text = Precio_Venta
            Cambio = False
        End If
    End If
End Sub


