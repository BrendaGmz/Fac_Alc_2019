VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Notas_Credito 
   Caption         =   "NOTAS DE CRÉDITO"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11565
   Begin VB.PictureBox Pic_Notas_Credito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8385
      ScaleWidth      =   11505
      TabIndex        =   37
      Top             =   0
      Width           =   11535
      Begin VB.Frame Fra_Datos_NC 
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
         Height          =   2925
         Left            =   7995
         TabIndex        =   49
         Top             =   480
         Width           =   3420
         Begin VB.ComboBox Cmb_Forma_Pago 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":0000
            Left            =   1245
            List            =   "Frm_Adm_Notas_Credito.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   2160
            Width           =   2055
         End
         Begin VB.ComboBox Cmb_Metodo_Pago 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":0004
            Left            =   1245
            List            =   "Frm_Adm_Notas_Credito.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   2505
            Width           =   2055
         End
         Begin VB.TextBox Txt_UUID_Relacion 
            Height          =   285
            Left            =   2880
            TabIndex        =   72
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox Cmb_Relacionados 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":0008
            Left            =   1230
            List            =   "Frm_Adm_Notas_Credito.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1800
            Width           =   2055
         End
         Begin VB.ComboBox Cmb_FacRef 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":000C
            Left            =   2020
            List            =   "Frm_Adm_Notas_Credito.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1365
            Width           =   1335
         End
         Begin VB.ComboBox Cmb_Serie 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":0010
            Left            =   1300
            List            =   "Frm_Adm_Notas_Credito.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1365
            Width           =   735
         End
         Begin VB.TextBox Txt_Factura_Referencia 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3000
            TabIndex        =   16
            Top             =   720
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox Txt_No_Nota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   615
            Width           =   1065
         End
         Begin VB.TextBox Txt_Serie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   615
            Width           =   615
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_NC 
            Height          =   300
            Left            =   1230
            TabIndex        =   15
            Top             =   990
            Width           =   1750
            _ExtentX        =   3096
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   6619139
            CurrentDate     =   38712
         End
         Begin VB.Label Lbl_Forma_Pago 
            BackColor       =   &H80000005&
            Caption         =   "Forma Pago"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   2235
            Width           =   1095
         End
         Begin VB.Label Lbl_Metodo_Pago 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Método Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   2595
            Width           =   960
         End
         Begin VB.Label Lbl_Tipo_Relacion 
            BackColor       =   &H80000005&
            Caption         =   "Tipo Relación"
            Height          =   255
            Left            =   75
            TabIndex        =   71
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Fact. Referencia"
            Height          =   195
            Left            =   75
            TabIndex        =   63
            Top             =   1485
            Width           =   1185
         End
         Begin VB.Label Lbl_No_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "No. Documento"
            Height          =   195
            Left            =   75
            TabIndex        =   51
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label Lbl_Fecha_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   75
            TabIndex        =   50
            Top             =   1080
            Width           =   450
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
            TabIndex        =   12
            Top             =   225
            Width           =   2895
         End
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
         Left            =   9825
         Picture         =   "Frm_Adm_Notas_Credito.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "A"
         Top             =   7680
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
         Left            =   360
         Picture         =   "Frm_Adm_Notas_Credito.frx":3713
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "A"
         Top             =   7680
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
         Left            =   2250
         Picture         =   "Frm_Adm_Notas_Credito.frx":6C4A
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "A"
         Top             =   7680
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
         Left            =   4140
         Picture         =   "Frm_Adm_Notas_Credito.frx":A110
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "A"
         Top             =   7680
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
         Left            =   7935
         Picture         =   "Frm_Adm_Notas_Credito.frx":D7A7
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "C"
         Top             =   7680
         UseMaskColor    =   -1  'True
         Width           =   1350
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
         Left            =   6045
         Picture         =   "Frm_Adm_Notas_Credito.frx":10D33
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "A"
         Top             =   7680
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Detalle_NC 
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
         Height          =   3915
         Left            =   240
         TabIndex        =   52
         Top             =   3360
         Width           =   11175
         Begin VB.ComboBox Cmb_Descripcion_Sat 
            Height          =   315
            Left            =   2505
            TabIndex        =   68
            Top             =   840
            Width           =   5175
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
            Picture         =   "Frm_Adm_Notas_Credito.frx":112BD
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2670
            Width           =   1260
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
            Height          =   795
            Left            =   120
            TabIndex        =   57
            Top             =   3120
            Width           =   7455
            Begin VB.TextBox Txt_Comentarios 
               Height          =   570
               Left            =   30
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   27
               Top             =   180
               Width           =   7395
            End
         End
         Begin VB.TextBox Text_Impuesto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2325
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.ComboBox Cmb_Descripcion 
            Height          =   315
            Left            =   2505
            TabIndex        =   20
            Top             =   480
            Width           =   5175
         End
         Begin VB.TextBox Txt_Importe 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   8745
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   480
            Width           =   1185
         End
         Begin VB.TextBox Txt_Precio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            TabIndex        =   21
            Top             =   480
            Width           =   1050
         End
         Begin VB.TextBox Txt_Cantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   105
            TabIndex        =   17
            Top             =   480
            Width           =   930
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
            Left            =   9930
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Adm_Notas_Credito.frx":1456F
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   390
            Width           =   750
         End
         Begin VB.TextBox Txt_Aplica_IVA 
            Height          =   285
            Left            =   3090
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   255
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
            Left            =   7650
            TabIndex        =   53
            Top             =   2685
            Width           =   3390
            Begin VB.TextBox Txt_Total 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   795
               Width           =   2160
            End
            Begin VB.TextBox Txt_Subtotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   165
               Width           =   2160
            End
            Begin VB.TextBox Txt_IVA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   480
               Width           =   2160
            End
            Begin VB.Label Lbl_Total 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Total"
               Height          =   195
               Left            =   75
               TabIndex        =   56
               Top             =   840
               Width           =   360
            End
            Begin VB.Label Lbl_IVA 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "I.V.A."
               Height          =   195
               Left            =   75
               TabIndex        =   55
               Top             =   525
               Width           =   390
            End
            Begin VB.Label Lbl_Subtotal 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Subtotal"
               Height          =   195
               Left            =   75
               TabIndex        =   54
               Top             =   210
               Width           =   585
            End
         End
         Begin VB.ComboBox Cmb_Unidad 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":17825
            Left            =   1035
            List            =   "Frm_Adm_Notas_Credito.frx":17835
            TabIndex        =   18
            Top             =   480
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Detalle_NC 
            Height          =   1470
            Left            =   105
            TabIndex        =   25
            Top             =   1170
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   2593
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Lbl_Descripcion_Sat 
            BackColor       =   &H80000005&
            Caption         =   "Descripción SAT"
            Height          =   255
            Left            =   1200
            TabIndex        =   69
            Top             =   890
            Width           =   1335
         End
         Begin VB.Label Lbl_Importe 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Importe"
            Height          =   195
            Left            =   8970
            TabIndex        =   62
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Lbl_Precio 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Precio"
            Height          =   195
            Left            =   8025
            TabIndex        =   61
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Lbl_Descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Descripción"
            Height          =   195
            Left            =   4455
            TabIndex        =   60
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Lbl_Cantidad 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   225
            TabIndex        =   59
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Unidad"
            Height          =   195
            Left            =   1320
            TabIndex        =   58
            Top             =   240
            Width           =   510
         End
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
         Height          =   2835
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   7695
         Begin VB.ComboBox Cmb_Uso_CFDI 
            Height          =   315
            ItemData        =   "Frm_Adm_Notas_Credito.frx":17859
            Left            =   945
            List            =   "Frm_Adm_Notas_Credito.frx":1785B
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1920
            Width           =   6600
         End
         Begin VB.TextBox Txt_Pais 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1248
            Width           =   1815
         End
         Begin VB.TextBox Txt_Estado 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1248
            Width           =   1725
         End
         Begin VB.TextBox Txt_Cliente_ID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   735
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox Txt_Ciudad_Cliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   907
            Width           =   1815
         End
         Begin VB.TextBox Txt_RFC_Cliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   566
            Width           =   1725
         End
         Begin VB.ComboBox Cmb_Nombre_Cliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   945
            TabIndex        =   1
            Top             =   195
            Width           =   6600
         End
         Begin VB.TextBox Txt_Direccion_Cliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   566
            Width           =   2490
         End
         Begin VB.TextBox Txt_No_Exterior 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6015
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   566
            Width           =   820
         End
         Begin VB.TextBox Txt_No_Interior 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6855
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   566
            Width           =   700
         End
         Begin VB.TextBox Txt_Colonia_Cliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   907
            Width           =   1725
         End
         Begin VB.TextBox Txt_Codigo_Postal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6015
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   907
            Width           =   1550
         End
         Begin VB.TextBox Txt_Email 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1590
            Width           =   6600
         End
         Begin VB.Label Lbl_Uso_CFDI 
            BackColor       =   &H80000005&
            Caption         =   "Uso CFDI"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Nombre"
            Height          =   195
            Left            =   105
            TabIndex        =   48
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Lbl_Direccion_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Dirección"
            Height          =   195
            Left            =   2745
            TabIndex        =   47
            Top             =   611
            Width           =   675
         End
         Begin VB.Label Lbl_Colonia_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Colonia"
            Height          =   195
            Left            =   105
            TabIndex        =   46
            Top             =   952
            Width           =   525
         End
         Begin VB.Label Lbl_Ciudad_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   2745
            TabIndex        =   45
            Top             =   952
            Width           =   495
         End
         Begin VB.Label Lbl_RFC_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "R.F.C."
            Height          =   195
            Left            =   105
            TabIndex        =   44
            Top             =   611
            Width           =   450
         End
         Begin VB.Label Lbl_CP 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "CP"
            Height          =   195
            Left            =   5580
            TabIndex        =   43
            Top             =   952
            Width           =   210
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Estado"
            Height          =   195
            Left            =   105
            TabIndex        =   42
            Top             =   1293
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Pais"
            Height          =   195
            Left            =   2745
            TabIndex        =   41
            Top             =   1293
            Width           =   300
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Email"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1635
            Width           =   375
         End
      End
      Begin VB.Label lbl_estatus_cancel 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   7320
         Visible         =   0   'False
         Width           =   10815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOTAS DE CRÉDITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         TabIndex        =   38
         Top             =   120
         Width           =   3795
      End
   End
End
Attribute VB_Name = "Frm_Adm_Notas_Credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Buscar_Click()
    Busca_Nota_Credito
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Cancela_nota_Credito
'DESCRIPCIÓN                : Realiza el proceso de cancelación de una nota de crédito
'PARÁMETROS                 :
'CREO                       : Sergio Godínez Banda
'FECHA_CREO                 : 29-Agosto-2012
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Public Sub Cancela_Nota_Credito()
Dim Rs_Cancela_Factura_Clientes As rdoResultset            'Manejo del registro de Adm_Factura_Clientes
Dim Rs_Cancela_Factura_Clientes_Detalles As rdoResultset
Dim Resultado As Integer
Dim Mi_SQL As String
Dim Cont_Fila As Integer
Dim Codigo_UUID As String                   'Almacena el codigo fiscal sat
Dim Motivo As String
Dim Factura() As String
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
        Mi_SQL = "SELECT No_Nota_Credito, Cancelada,Mensaje_Cancelado, Usuario_Cancelo, Fecha_Cancelo, Motivo_Cancelo, Timbre_UUID"
        Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito"
        Mi_SQL = Mi_SQL & " WHERE No_Nota_Credito = '" & Format(Txt_No_Nota.text, "0000000000") & "'"
        Set Rs_Cancela_Factura_Clientes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            'Llena la tabla de Adm_Clientes_Facturas con los datos contenidos en las cajas de textos
            With Rs_Cancela_Factura_Clientes
                .Edit
                    If Not IsNull(.rdoColumns("Timbre_UUID")) Then
                        Codigo_UUID = .rdoColumns("Timbre_UUID")
                    End If
'                    Mensaje = CFD_Cancela_Xml(Codigo_UUID, False)
                    If Lbl_Facturacion.Caption = "CANCELACIÓN EN PROCESO" Then
                        Mensaje = CFD_Cancela_Xml(Codigo_UUID, True)
                    Else
                        Mensaje = CFD_Cancela_Xml(Codigo_UUID, False)
                    End If
                    
                    If Mensaje Like "*Cancelación exitosa*" Or Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                        Cancel = True
                       .rdoColumns("Cancelada") = "S"
                       Factura = Split(Txt_Factura_Referencia.text, " ")
                       If UBound(Factura) > 0 Then
                            Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas"
                            Mi_SQL = Mi_SQL & " WHERE No_Factura_Electronica = '" & Format(Val(Factura(1)), "0000000000") & "'"
                            Mi_SQL = Mi_SQL & " AND Serie='" & Factura(0) & "'"
                            Set Rs_Modifica_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                                With Rs_Modifica_Factura
                                    If Not .EOF Then
                                        .Edit
                                            .rdoColumns("Saldo") = Val(.rdoColumns("Saldo")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                                            .rdoColumns("Pagada") = "N"
                                        .Update
                                    End If
                                End With
                            Rs_Modifica_Factura.Close
                        End If
                    Else
                       Cancel = False
                       .rdoColumns("Cancelada") = "PC"
                    End If
                    .rdoColumns("Mensaje_Cancelado") = Mensaje
                    .rdoColumns("Fecha_Cancelo") = Now()
                    .rdoColumns("Motivo_Cancelo") = Trim(Motivo)
                    .rdoColumns("Usuario_Cancelo") = Nombre_Usuario
                    lbl_estatus_cancel.Caption = Mensaje
                    lbl_estatus_cancel.Visible = True
                .Update
            End With
        Rs_Cancela_Factura_Clientes.Close
        'Genera la estructura del xml y cancelacion del timbrado
        
    Conexion_Base.CommitTrans
    MDIFrm_Apl_Principal.MousePointer = 0
    If Cancel Then
        MsgBox "Documento Cancelado Existosamente", vbInformation
        Lbl_Facturacion.Caption = "CANCELADA"
    Else
        MsgBox "Documento en Proceso de Cancelación", vbInformation
        Lbl_Facturacion.Caption = "Cancelación en Proceso"
    End If
    Btn_Cancelar.Enabled = False
    Btn_Imprimir.Enabled = False
    Btn_Enviar_Email.Enabled = False
    
    Fra_Datos_Cliente.Enabled = False
'    Fra_Datos_Factura.Enabled = False
'    Fra_Detalle_Factura.Enabled = False
    Fra_Comentarios.Enabled = False
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    'Obtiene el error
    If Err.Number = 7777 Or Err.Number = -255 Then
        MsgBox Err.Description
    Else
        For Each Rdo_Error In rdoErrors
            MsgBox Rdo_Error.Description
        Next
    End If
End Sub

Private Sub Btn_Cancelar_Click()
    If Txt_No_Nota.text <> "" Then
        If MsgBox("¿Esta segura de cancelar la nota de crédito?", vbYesNo + vbQuestion, "CANCELACIÓN DE DOCUMENTOS") = vbYes Then
            Cancela_Nota_Credito
        End If
    Else
        MsgBox "Consulte una nota de crédito para poder cancelar", vbExclamation
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
    Mi_SQL = "SELECT No_Factura_Electronica FROM Adm_Clientes_Facturas WHERE Cliente_ID = '" & Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "#00000") & "' and Timbre_UUID is not null and Timbre_UUID<>'' and Cancelada='N'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta.EOF
         Cmb_FacRef.AddItem (Rs_Consulta.rdoColumns("No_Factura_Electronica"))
         Rs_Consulta.MoveNext
    Wend
     Rs_Consulta.Close
End Sub

Private Sub Btn_Nuevo_Click()
Set Conectar_Ayudante = New Ayudante  'Manejador del ayudante
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        Cmb_Nombre_Cliente.text = ""
        Cmb_Descripcion.Clear
        Call Conectar_Ayudante.Limpiar_Textos(Frm_Adm_Notas_Credito)
        'Consulta la serie del rango activo actual
        Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
        Mi_SQL = Mi_SQL & " AND Tipo = 'NOTA_CREDITO'"
        Set Rs_Consulta_Serie = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Serie.EOF Then
                Txt_Serie.text = Trim(Rs_Consulta_Serie.rdoColumns("Serie"))
            End If
        Rs_Consulta_Serie.Close
        'Valida si aun existen folios de facturas disponibles para utilizar
        Call Aviso_Termino_Folios("NOTA_CREDITO")
        'si la bandera esta habilitada muestra mensaje y cancela la operación
        If Folios_Terminados = True Then
            Txt_No_Nota.text = ""
            Txt_Serie.text = ""
            MsgBox "No se encontraron folios de nota de crédito disponibles, favor de verificar", vbCritical
            Exit Sub
        Else
            Txt_No_Nota.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Notas_Credito WHERE Serie = '" & Trim(Txt_Serie.text) & "'", "No_Nota_Credito")
        End If
        'Habilita los controles para poder capturar la información y deshabilita otros
        Fra_Datos_Cliente.Enabled = True
        Fra_Datos_NC.Enabled = True
        Fra_Detalle_NC.Enabled = True
        Fra_Comentarios.Enabled = True
        Cmb_Nombre_Cliente.Enabled = True
        Cmb_Nombre_Cliente.text = ""
        Cmb_Nombre_Cliente_KeyPress 13
        Cmb_Nombre_Cliente.SetFocus
        Cmb_Descripcion_KeyPress 13
        Btn_Imprimir.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Nuevo.Visible = True
        Btn_Buscar.Enabled = False
        Btn_Cancelar.Enabled = False
        Btn_Enviar_Email.Enabled = False
        Btn_Salir.Caption = "Cancelar"
        Lbl_Facturacion.Caption = "Estatus"
        Btn_Nuevo.Caption = "Dar de Alta"       'Cambia el texto del botón
        Dtp_Fecha_NC.Value = Now
        Grid_Detalle_NC.Rows = 0
        Btn_Agregar.Enabled = True
        lbl_estatus_cancel.Visible = False
    Else
        'Validacion de que esten todos los datos requeridos para dar de alta la factura
        If Grid_Detalle_NC.Rows > 1 Then
            If Txt_No_Nota.text <> "" Then
                If Cmb_Nombre_Cliente.ListIndex > -1 Then
                    If Cmb_Serie.ListIndex > -1 And Cmb_FacRef.ListIndex > -1 Then
                        Txt_Factura_Referencia.text = Cmb_Serie.text & " " & Cmb_FacRef.text
                    Else
                        MsgBox "Falta la factura referenciada, favor de verificar", vbExclamation
                        Cmb_Serie.SetFocus
                        Exit Sub
                    End If
                    If Cmb_Uso_CFDI.ListIndex < 0 Then
                        MsgBox "Falta el uso del CFDI, favor de verificar", vbExclamation
                        Cmb_Uso_CFDI.SetFocus
                    End If
                    If Cmb_Relacionados.ListIndex = -1 Then
                        MsgBox "Falta el tipo de relación, favor de verificar", vbExclamation
                        Cmb_Relacionados.SetFocus
                        Exit Sub
                    Else
                        Busca_UUID
                    End If
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
                    'Alta en la base de datos de la nota de crédito
                    Alta_Nota_Credito
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

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Grid_Detalle_NC.Rows = 0
        Cmb_Nombre_Cliente.text = ""
        Cmb_Descripcion.Clear
        Btn_Nuevo.Enabled = True
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Buscar.Enabled = True
        Btn_Salir.Caption = "Salir"
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Fra_Datos_Cliente.Enabled = False
        Fra_Datos_NC.Enabled = False
        Fra_Detalle_NC.Enabled = False
        Fra_Comentarios.Enabled = False
        lbl_estatus_cancel.Visible = False
    End If
End Sub

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
Public Sub Busca_UUID()
    Dim Rs_Consulta As rdoResultset
    Txt_UUID_Relacion.text = ""
    Mi_SQL = "SELECT Timbre_UUID, Tipo_Pago, Forma_Pago FROM Adm_Clientes_Facturas WHERE Serie='" & Cmb_Serie.text & "' and No_Factura_Electronica='" & Cmb_FacRef.text & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
         Txt_UUID_Relacion.text = Rs_Consulta.rdoColumns("Timbre_UUID")
    End If
    Rs_Consulta.Close
End Sub

Private Sub Cmb_Relacionados_Click()
    CFD_Relacionados.Existe = True
End Sub


Private Sub Form_Load()
    Me.Top = 300
    Me.Height = 9000
    Me.Width = 11805
    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Nombre_Cliente, 1, "Nombre")
    'Call Conectar_Ayudante.Llena_Combo_Item("Presentacion_ID,Nombre", "Cat_Presentaciones", Cmb_Unidad, 1, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Clave_Unidad + '-' + Nombre", "Cat_Unidades_Medida", Cmb_Unidad, 1, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Codigo_Uso_Comprobante + ' ' + Descripcion as Descripcion", "Cat_Uso_Comprobantes", Cmb_Uso_CFDI, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Codigo_Tipo_Relacion + '-' + Descripcion", "Cat_Tipos_Relacion", Cmb_Relacionados, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Metodo_ID + ' ' + Descripcion ", "Cat_Metodo_Pago", Cmb_Metodo_Pago, 1, "Descripcion")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Clave + ' ' + Descripcion", "Cat_Formas_Pago", Cmb_Forma_Pago, 1, "Descripcion")
    Call Cmb_Nombre_Cliente_KeyPress(13)
    Call Cmb_Descripcion_KeyPress(13)
    Limpia_Variables
    Consulta_Cancelados
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
    Txt_Codigo_Postal.text = ""
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
    
    'Realiza la consulta para enviarla al recordset
    Mi_SQL = "SELECT Credito_Flexible,Cliente_ID, Cat_Clientes.Nombre as Nombre_C, Dias_Credito, RFC, Direccion, No_Ext,"
    Mi_SQL = Mi_SQL & " No_Int, Colonia, Ciudad, Estado, Pais, Telefono, CP, Metodo_Pago, Cuenta_Pago, Email"
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
            If Not IsNull(.rdoColumns("Cliente_ID")) Then Txt_Cliente_ID.text = .rdoColumns("Cliente_ID")
            If Not IsNull(.rdoColumns("CP")) Then Txt_Codigo_Postal.text = .rdoColumns("CP")
        Else
            Exit Sub
        End If
    End With
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

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Alta_Nota_Credito
'DESCRIPCIÓN            : Hace el alta de la nota de crédito en la base de datos
'PARÁMETROS             :
'CREO                   : Sergio Godínez Banda
'FECHA_CREO             : 29-Agosto-2012
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Alta_Nota_Credito()
Dim Rs_Alta_Factura_Clientes As rdoResultset                            'Manejo del registro de Adm_Factura_Clientes
Dim Rs_Alta_Descripcion_Facturas As rdoResultset                        'Manejo del registro de Adm_Descripcion_Facturas
Dim Rs_Alta_Remision_Clientes As rdoResultset                           'Manejo del registro de Adm_Remision_Clientes
Dim Rs_Alta_Descripcion_Remision As rdoResultset                        'Manejo del registro de Adm_Descripcion_Remision
Dim Rs_Modifica_Alm_Salidas_Almacen_Detalles As rdoResultset            'Manejo del registro de Adm_Descripcion_Remision
Dim Rs_Modifica_Alm_Salidas_Almacen As rdoResultset
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
Dim Fecha_Generacion As Date
Dim Copias As Integer
Dim Impresiones As Integer
Dim Unidades() As String
Dim Rs_Consulta_Regimen As rdoResultset
Dim Factura_Referencia() As String
On Error GoTo handler
    Conexion_Base.BeginTrans
        Me.MousePointer = 11
        Txt_No_Nota.text = Conectar_Ayudante.Maximo_Catalogo("Adm_Notas_Credito WHERE Serie = '" & Trim(Txt_Serie.text) & "'", "No_Nota_Credito")
        'Asigna los valores a las variable para la generación de la factura electrónica
        CFD_Generales.Version = "3.3"
        CFD_Generales.Serie = Trim(Txt_Serie.text)
        CFD_Generales.Folio = Val(Txt_No_Nota.text)
        CFD_Generales.Factura_ID = Format(Txt_No_Nota.text, "0000000000")
        'Asigna la fecha del xml
        Fecha_Xml = Format(Dtp_Fecha_NC.Value, "dd/MM/yyyy") & " " & Format(Now, "HH:mm:ss")
        CFD_Generales.Fecha = Format(Dtp_Fecha_NC.Value, "yyyy-MM-dd") & "T" & Format(Dtp_Fecha_NC.Value, "HH:mm:ss")
        CFD_Generales.Forma_Pago = CFD_Elimina_Espacios(Cmb_Forma_Pago.text)
        CFD_Generales.Condiciones_Pago = ""
        CFD_Generales.SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
        CFD_Generales.Total = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
        Impuesto = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
        CFD_Generales.Descuento = 0
        CFD_Generales.Tipo_Moneda = "MXN"
        CFD_Generales.Tipo_Documento = "NORMAL"
        CFD_Generales.Tipo_Comprobante = "E"
        CFD_Generales.Fecha_Vencimiento = ""
        CFD_Generales.Metodo_Pago = CFD_Elimina_Espacios(Cmb_Metodo_Pago.text)
        CFD_Generales.Cuenta_Pago = "No Identificado"
        CFD_Generales.Uso_CFDI = Cmb_Uso_CFDI.text
        CFD_Receptor.Uso_CFDI = Cmb_Uso_CFDI.text
        'Documentos relacionados
        CFD_Relacionados.Relacionados = Cmb_Relacionados.text
        CFD_Relacionados.UUID_Relacionados = Txt_UUID_Relacion.text
        ReDim CFD_Relacionados_Conceptos(0)
        CFD_Relacionados.Existe = True

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
        Mi_SQL = "SELECT Mensaje_Factura FROM Cat_Parametros_Factura_Electronica"
        Set Rs_Consulta_Regimen = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        CFD_Emisor.Regimen_Fiscal = Rs_Consulta_Regimen.rdoColumns("Mensaje_Factura")
        Rs_Consulta_Regimen.Close
        
        Contador = 0
        'Valida que el numero de partidas
        For Cont_Detalles_Factura = 1 To Grid_Detalle_NC.Rows - 1
            Contador = Contador + 1
        Next
        'Asigna el conteo de partidas al arreglo
        ReDim CFD_Conceptos(Contador)
        'Recorre las partidas del grid
        Contador = 0
        For Cont_Detalles_Factura = 1 To Grid_Detalle_NC.Rows - 1
            Contador = Contador + 1
'            If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2) <> "" Then
'                CFD_Conceptos(Contador).No_Identificacion = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 2))
'            End If
            Unidades = Split(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 10), "-")
            CFD_Conceptos(Contador).Unidad = Unidades(1)
            CFD_Conceptos(Contador).Unidad_Medida = Unidades(0)
            CFD_Conceptos(Contador).Cod_prod = Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 12)
            CFD_Conceptos(Contador).Cantidad = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 0), ","))
            CFD_Conceptos(Contador).Descripcion = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 1))
            CFD_Conceptos(Contador).Valor_Unitario = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 2), ",")), "#0.00")
            CFD_Conceptos(Contador).Importe = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 3), ",")), "#0.00")
            'CFD_Conceptos(Contador).Unidad = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 10))
            If Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 9) = "SI" Then
                 CFD_Conceptos(Contador).IVA_Producto = True
            Else
                CFD_Conceptos(Contador).IVA_Producto = False
            End If
        Next
        
        'Asigna los datos de los IMPUESTO al CFD para generar el xml segun el tipo de factura
        ReDim CFD_Impuestos(0)
        ReDim CFD_Impuestos_Retenidos(0)
        ReDim CFD_Impuestos_Locales(0)
        If Impuesto > 0 Then
            ReDim CFD_Impuestos(1)
            CFD_Impuestos(1).Impuesto = "002"
            CFD_Impuestos(1).Tasa = PG_Retencion_IVA
            CFD_Impuestos(1).Importe = Impuesto
            IVA_EXENTO = False
        Else
'            ReDim CFD_Impuestos(1)
'            CFD_Impuestos(1).Impuesto = "002"
'            CFD_Impuestos(1).Tasa = "0"
'            CFD_Impuestos(1).Importe = 0
            IVA_EXENTO = True
        End If
        'Crea el sello digital con toda la informacion
        CFD_Generales.No_Certificado = CFD_Consulta_Serie_Certificado(Ruta_Certificado)
        CFD_Generales.Certificado = CFD_Consulta_Certificado(Ruta_Certificado)
        Str_Cadena_Original = CFD_Cadena_Original("")
        Str_Cadena_UTF = CFD_Valida_Caracteres_UTF(Str_Cadena_Original)
        Str_Cadena_MD5 = CFD_Genera_MD5(Str_Cadena_UTF)
        Str_Cadena_Sello = CFD_Genera_Sello(Str_Cadena_UTF, Ruta_Llave_Privada)
        CFD_Generales.Cadena_Original = Str_Cadena_UTF
        
        CFD_Generales.Sello = Str_Cadena_Sello
        CFD_Generales.Importe_Letra = Conectar_Ayudante.Convierte_Cantidad_Letras(Format(CStr(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))), "#0.00"))
        'Registra laminformacion en la BD
        Set Rs_Alta_Factura_Clientes = Conectar_Ayudante.Recordset_Agregar("Adm_Notas_Credito")
            With Rs_Alta_Factura_Clientes
                .AddNew
                    .rdoColumns("No_Nota_Credito") = Format(Txt_No_Nota.text, "0000000000")
                    .rdoColumns("Serie") = Trim(Txt_Serie.text)
                    'Convierte la fecha de timbrado
                    Grupo_Fecha = Split(CFD_Generales.Fecha, "T")
                    Fecha_Generacion = Grupo_Fecha(0) & " " & Grupo_Fecha(1)
                    .rdoColumns("Fecha_Creo_XML") = Format(Fecha_Generacion, "MM/dd/yyyy HH:mm:ss")
                    .rdoColumns("Cliente_ID") = Format(Cmb_Nombre_Cliente.ItemData(Cmb_Nombre_Cliente.ListIndex), "00000")
                    .rdoColumns("Fecha") = Format(Dtp_Fecha_NC.Value, "MM/dd/yyyy")
                    .rdoColumns("Subtotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
                    .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
                    .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                    .rdoColumns("Cancelada") = "N"
                    .rdoColumns("Comentarios") = Txt_Comentarios.text
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now()
                    .rdoColumns("Metodo_Pago") = CFD_Generales.Metodo_Pago
                    .rdoColumns("No_Cuenta_Pago") = CFD_Generales.Cuenta_Pago
                    .rdoColumns("Factura_Referencia") = Trim(Txt_Factura_Referencia.text)
                    .rdoColumns("Relacionado") = CFD_Relacionados.Relacionados
                    .rdoColumns("UUID_Relacionado") = CFD_Relacionados.UUID_Relacionados
                    .rdoColumns("Uso_CFDI") = CFD_Generales.Uso_CFDI
                    '.rdoColumns("Forma_Pago") = CFD_Generales.Forma_Pago
                .Update
            End With
        Rs_Alta_Factura_Clientes.Close
        'Registra los detalles
        Set Rs_Alta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Agregar("Adm_Notas_Credito_Detalles")
            For Cont_Detalles_Factura = 1 To Grid_Detalle_NC.Rows - 1
                'Llena la tabla de Adm_Clientes_Facturas_Detalles con los datos contenidos en el grid
                With Rs_Alta_Descripcion_Facturas
                    .AddNew
                        .rdoColumns("No_Nota_Credito") = Format(Txt_No_Nota.text, "0000000000")
                        If Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 4) <> "" Then
                            .rdoColumns("Producto_ID") = Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 4)
                        Else
                            .rdoColumns("Producto_ID") = Null
                        End If
                        .rdoColumns("Descripcion") = Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 1)
                        .rdoColumns("Cantidad") = Val(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 0))
                        .rdoColumns("Precio") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 2), ","))
                        .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 3), ","))
                        .rdoColumns("Unidad") = Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 10)
                        .rdoColumns("Clave_SAT") = Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 12)
                    .Update
                End With
            Next Cont_Detalles_Factura
        Rs_Alta_Descripcion_Facturas.Close
        
        Factura_Referencia = Split(Txt_Factura_Referencia.text, " ")
        'Actualiza la factura con el timbrado
        Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas"
        Mi_SQL = Mi_SQL & " WHERE No_Factura_Electronica = '" & Format(Val(Factura_Referencia(1)), "0000000000") & "'"
        Mi_SQL = Mi_SQL & " AND Serie='" & Trim(Factura_Referencia(0)) & "'"
        Set Rs_Modifica_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            With Rs_Modifica_Factura
                If Not .EOF Then
                    .Edit
                        .rdoColumns("Saldo") = .rdoColumns("Saldo") - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
                        If .rdoColumns("Saldo") = 0 Then
                            .rdoColumns("Pagada") = "S"
                        End If
                    .Update
                End If
            End With
        Rs_Modifica_Factura.Close
        
        
        'Crea el xml con los datos de la factura
        Call CFD_Crea_Xml("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text), "NOTA", "")
        
        'Convierte la fecha de timbrado
        Grupo_Fecha_Timbrado = Split(Timbrado_FechaTimbrado, "T")
        Fecha_Timbrado = Grupo_Fecha_Timbrado(0) & " " & Grupo_Fecha_Timbrado(1)
        
        'Actualiza la factura con el timbrado
        Mi_SQL = "SELECT * FROM Adm_Notas_Credito"
        Mi_SQL = Mi_SQL & " WHERE No_Nota_Credito = '" & Format(Val(Txt_No_Nota.text), "0000000000") & "'"
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
                        .rdoColumns("Ruta_Codigo_BD") = Ruta_NC & "\CFDI_" & Trim(Txt_Serie.text) & "_" & Trim(CFD_Generales.Folio) & ".bmp"
                        CFD_Generales.Imagen_BMP = Ruta_NC & "\CFDI_" & Trim(Txt_Serie.text) & "_" & Trim(CFD_Generales.Folio) & ".bmp"
                    .Update
                End If
            End With
        Rs_Modifica_Factura.Close
    Conexion_Base.CommitTrans
    
    Me.MousePointer = 0
    MsgBox "La nota de crédito ha sido dada de alta", vbInformation
    Me.MousePointer = 11
    Call Valida_Termino_Folios_Activos("NOTA_CREDITO", Trim(Txt_Serie.text), Trim(Txt_No_Nota.text))
    Call CFD_Crea_PDF("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text), "NC", "NORMAL", Year(Fecha_Xml))
    Muestra_PDF
    If MsgBox("¿Desea realizar el envío de la nota de crédito por correo", vbYesNo + vbQuestion, "ENVÍO DE CORREO") = vbYes Then
        Btn_Enviar_Email_Click
    End If
    Fra_Datos_Cliente.Enabled = False
    Fra_Datos_NC.Enabled = False
    Fra_Detalle_NC.Enabled = False
    Fra_Comentarios.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Imprimir.Enabled = True
    Btn_Cancelar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Enviar_Email.Enabled = True
    Me.MousePointer = 0
    Exit Sub
handler:
    Me.MousePointer = 0
    Conexion_Base.RollbackTrans
    MsgBox Err.Description
'    For Each Er In rdoErrors
'        MsgBox Er.Description
'    Next Er
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
Private Sub Muestra_PDF()
Dim Nombre_Archivo As String   'Almacena el nombre del archivo

    If Txt_No_Nota.text <> "" Then
        'Asigna el nombre del archivo
        Nombre_Archivo = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text) & ".pdf"
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_NC & "\" & Nombre_Archivo, "ARCHIVO") = True Then
            'Envia para abrir el archivo
            ShellExecute ByVal 0&, "open", Ruta_NC & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
        Else 'Regenera el pdf
            Regenerar_PDF_NC
            'Envia para abrir el archivo
            ShellExecute ByVal 0&, "open", Ruta_NC & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
        End If
    Else
        MsgBox "Seleccione la nota de crédito", vbExclamation
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Regenerar_PDF_NC
'DESCRIPCIÓN: Regenera el PDF de la nota de crédito seleccionada
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 29-Agosto-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'******************************************************************************'
Private Sub Regenerar_PDF_NC()
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
    Mi_SQL = "SELECT No_Nota_Credito, Subtotal, IVA, Total, Fecha_Creo_XML, No_Certificado, Timbre_Version, Timbre_UUID, Timbre_Fecha_Timbrado,"
    Mi_SQL = Mi_SQL & " Timbre_SelloCFD, Timbre_noCertificadoSAT, Timbre_selloSAT, Ruta_Codigo_BD"
    Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito"
    Mi_SQL = Mi_SQL & " WHERE No_Nota_Credito = '" & Format(Val(Txt_No_Nota.text), "0000000000") & "'"
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
    CFD_Generales.Folio = Val(Txt_No_Nota.text)
    CFD_Generales.Factura_ID = Format(Txt_No_Nota.text, "0000000000")
    'Asigna la fecha del xml
    CFD_Generales.Fecha = Format(Fecha_Xml, "yyyy-MM-dd") & "T" & Format(Fecha_Xml, "HH:mm:ss")
    CFD_Generales.Forma_Pago = CFD_Elimina_Espacios("pago en una sola exhibicion")
    CFD_Generales.Condiciones_Pago = ""
    CFD_Generales.SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.text, ","))
    CFD_Generales.Total = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","))
    Impuesto = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.text, ","))
    CFD_Generales.Descuento = 0
    CFD_Generales.Tipo_Moneda = "PESOS"
    CFD_Generales.Tipo_Documento = "NORMAL"
    CFD_Generales.Tipo_Comprobante = "egreso"
    
    CFD_Generales.Fecha_Vencimiento = ""
    CFD_Generales.Metodo_Pago = "No_Identificado"
    CFD_Generales.Cuenta_Pago = "No_Identificado"
    ReDim CFD_Relacionados_Conceptos(0)
            
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
    For Cont_Detalles_Factura = 1 To Grid_Detalle_NC.Rows - 1
        Contador = Contador + 1
    Next
    'Asigna el conteo de partidas al arreglo
    ReDim CFD_Conceptos(Contador)
    'Recorre las partidas del grid
    Contador = 0
    For Cont_Detalles_Factura = 1 To Grid_Detalle_NC.Rows - 1
        Contador = Contador + 1
'            If Grid_Detalle_Factura.TextMatrix(Cont_Detalles_Factura, 2) <> "" Then
'                CFD_Conceptos(Contador).No_Identificacion = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 2))
'            End If
        CFD_Conceptos(Contador).Cantidad = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 0), ","))
        CFD_Conceptos(Contador).Descripcion = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 1))
        CFD_Conceptos(Contador).Valor_Unitario = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 2), ",")), "#0.00")
        CFD_Conceptos(Contador).Importe = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 3), ",")), "#0.00")
        CFD_Conceptos(Contador).Unidad = CFD_Elimina_Espacios(Grid_Detalle_NC.TextMatrix(Cont_Detalles_Factura, 10))
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
    Call CFD_Crea_PDF("CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text), "NC", "NORMAL", Year(Fecha_Xml))
    
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

Private Sub Formatea_Columnas_Grid()
    If Grid_Detalle_NC.Rows > 1 Then
        Grid_Detalle_NC.FixedRows = 1
        Grid_Detalle_NC.ColWidth(0) = 800      'Cantidad
        Grid_Detalle_NC.ColWidth(1) = 5500     'Descripcion
        Grid_Detalle_NC.ColAlignment(1) = 1
        Grid_Detalle_NC.ColWidth(2) = 1000     'Precio
        Grid_Detalle_NC.ColWidth(3) = 1100     'Importe
        Grid_Detalle_NC.ColWidth(4) = 0        'Producto_ID
        Grid_Detalle_NC.ColWidth(5) = 0        'No_Salida
        Grid_Detalle_NC.ColWidth(6) = 0        'Impuesto
        Grid_Detalle_NC.ColWidth(7) = 0        '
        Grid_Detalle_NC.ColWidth(8) = 0        'IVA
        Grid_Detalle_NC.ColWidth(9) = 0        'aplica_IVA
        Grid_Detalle_NC.ColWidth(10) = 3000    'Unidad
        Grid_Detalle_NC.ColWidth(11) = 0 '600    'Incluir
        Grid_Detalle_NC.ColAlignment(11) = 3
        Grid_Detalle_NC.ColWidth(12) = 3000   'Clave SAT
    Else
        Grid_Detalle_NC.Rows = 0
    End If
End Sub

'Botón para agregar un registro en el grid de detalles de productos
Private Sub Btn_Agregar_Click()
Dim Cont_Detalles As Integer        'Usada para contar detalle del grid
Dim Suma As Double                  'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double              'Suma I.V.A.
    
    'Valida que los campos tengan valores
    If Val(Txt_Cantidad.text) >= 0 And (Cmb_Descripcion.text <> "" And Val(Txt_Precio.text) >= 0 And Val(Txt_Importe.text) >= 0) Then
        If Cmb_Unidad.ListIndex = -1 Then
            MsgBox "Indique la unidad de medida", vbExclamation
            Cmb_Unidad.SetFocus
            Exit Sub
        End If
        If Cmb_Descripcion_Sat.ListIndex = -1 Then
            MsgBox "Indique la descripción del SAT.", vbExclamation
            Cmb_Descripcion_Sat.SetFocus
            Exit Sub
        End If
        If Grid_Detalle_NC.Rows = 0 Then
            'Coloca el número de columnas
            Grid_Detalle_NC.Cols = 13
            'Pone el encabezado en las columnas
            Grid_Detalle_NC.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) _
                & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) & "" & Chr(9) & "IVA" & Chr(9) _
                & "Aplica_IVA" & Chr(9) & "Unidad" & Chr(9) & "Incluir" & Chr(9) & "Clave SAT"
        End If
        'Agrega el dato en el grid
        If Trim(Txt_Aplica_IVA.text) = "SI" Then
            If Cmb_Descripcion.ListIndex > -1 Then
                Grid_Detalle_NC.AddItem Txt_Cantidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) _
                    & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                    & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                    & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & "" & Chr(9) _
                    & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) _
                    & Val(PG_Retencion_IVA) * ((Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")))) & Chr(9) _
                    & "SI" & Chr(9) & Cmb_Unidad.text & Chr(9) & "SI" & Chr(9) & Cmb_Descripcion_Sat.text
            Else
                Grid_Detalle_NC.AddItem Txt_Cantidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) _
                    & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                    & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                    & "" & Chr(9) & "" & Chr(9) _
                    & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) _
                    & Val(PG_Retencion_IVA) * ((Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")))) & Chr(9) _
                    & "SI" & Chr(9) & Cmb_Unidad.text & Chr(9) & "SI" & Chr(9) & Cmb_Descripcion_Sat.text
            End If
        Else
            If Cmb_Descripcion.ListIndex > -1 Then
                Grid_Detalle_NC.AddItem Txt_Cantidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) _
                    & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                    & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                    & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & "" & Chr(9) _
                    & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) & "" & Chr(9) & "NO" & Chr(9) _
                    & Cmb_Unidad.text & Chr(9) & "SI" & Chr(9) & Cmb_Descripcion_Sat.text
            Else
                Grid_Detalle_NC.AddItem Txt_Cantidad.text & Chr(9) & UCase(Trim(Cmb_Descripcion.text)) & Chr(9) _
                    & Format((Txt_Precio.text), "#,##0.00") & Chr(9) _
                    & Format(Val(Txt_Cantidad.text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio.text, ",")), "#,##0.00") & Chr(9) _
                    & "" & Chr(9) & "" & Chr(9) & Val(Text_Impuesto.text) & Chr(9) & "" & Chr(9) & "" & Chr(9) _
                    & "NO" & Chr(9) & Cmb_Unidad.text & Chr(9) & "SI" & Chr(9) & Cmb_Descripcion_Sat.text
            End If
        End If
        
        Formatea_Columnas_Grid
        'Cacula los totales
        Suma = 0
        'Hace el recorrido de los datos del grid para hacer la suma
        For Cont_Detalles = 1 To Grid_Detalle_NC.Rows - 1
            Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles, 3), ","))
            Suma_IVA = Suma_IVA + Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles, 8), ",")), "#0.00")
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
        Txt_Cantidad.SetFocus
        Btn_Buscar.Enabled = False
        Cmb_Descripcion_Sat.ListIndex = -1
    Else
        MsgBox "Faltan datos para agregar", vbExclamation
    End If
End Sub

'Botón para eliminar un registro del grid de detalles de producto
Private Sub Btn_Eliminar_Click()
Dim Cont_Detalles As Integer            'Usada para contar detalles del grid
Dim Suma As Double                      'Usada para sumar el importe y manejo del I.V.A.
Dim Resp As Integer
Dim Suma_IVA As Double

    If Grid_Detalle_NC.Rows > 1 Then
  
        Resp = MsgBox("¿Esta seguro de borrar la partida?", vbYesNo + vbExclamation)
        If Resp = 6 Then
            'Si la respuesta es afirmativa elimina el registro seleccionado
            If Grid_Detalle_NC.Rows = 2 Then
                Grid_Detalle_NC.FixedRows = 0
                'Quita el item del grid
                Grid_Detalle_NC.RemoveItem (Grid_Detalle_NC.RowSel + 1)
                Btn_Buscar.Enabled = True
            Else
                If Grid_Detalle_NC.Rows > 2 Then
                    Grid_Detalle_NC.RemoveItem (Grid_Detalle_NC.RowSel)
                End If
            End If
            Suma = 0
            Productos = 0
            
            'Hace el recorrido de los datos del grid para hacer la suma
            For Cont_Detalles = 1 To Grid_Detalle_NC.Rows - 1
                Suma = Suma + CDbl(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles, 3), ","))
                Suma_IVA = Suma_IVA + Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_NC.TextMatrix(Cont_Detalles, 8), ",")), "#0.00")
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

'Botón para imprimir, llama a la función Imprimir_Factura
Private Sub Btn_Imprimir_Click()
    Muestra_PDF
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
        
'Función para que cuando cambie el texto de la cantidad calcule automáticamente el importe
Private Sub Txt_Cantidad_Change()
    Txt_Importe.text = (Val(Txt_Cantidad.text) * Val(Txt_Precio.text))
End Sub

Private Sub Txt_Cantidad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cantidad.text, True)
End Sub

'Función para que calcule la cantidad automáticamente al ingresar el precio manualmente
Private Sub Txt_Precio_Change()
    Txt_Importe.text = (Val(Txt_Cantidad.text) * Val(Txt_Precio.text))
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Busca_Nota_Credito
'DESCRIPCIÓN: Busca la nota de crédito de acuerdo al número proporcionado
'PARÁMETROS:
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  29-Agosto-2012
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Busca_Nota_Credito()
Dim Mi_SQL As String
Dim Rs_Consulta_Factura_Clientes As rdoResultset        'Manejo del registro para buscar facturas
Dim Rs_Consulta_Descripcion_Facturas As rdoResultset    'Manejo del registro para el detalle de la factura
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo del registro del catálogo de clientes
Dim Rs_Consulta_Movimientos_Facturas As rdoResultset    'Manejo del registro de movimientos
Dim No_Factura As String                                'Variable para capturar el número de la factura a buscar
Dim Suma As Double                  'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double              'Suma I.V.A.
Dim Unidad As String
    
    No_Factura = InputBox("Teclee el número de Documento a consultar", "Consulta de Documentos")
    If No_Factura <> "" Then
        'Prepara el recordset para consultar el número de factura de la tabla Adm_Clientes_Facturas
        Mi_SQL = "SELECT * FROM Adm_Notas_Credito"
        Mi_SQL = Mi_SQL & " WHERE No_Nota_Credito = '" & Format(No_Factura, "0000000000") & "'"
        Set Rs_Consulta_Factura_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena los controles con los datos de la consulta
        If Not Rs_Consulta_Factura_Clientes.EOF Then
            Txt_No_Nota.text = Val(Rs_Consulta_Factura_Clientes!No_Nota_Credito)
            Txt_Serie.text = Rs_Consulta_Factura_Clientes!Serie
            Dtp_Fecha_NC.Value = Rs_Consulta_Factura_Clientes!Fecha
            If Not IsNull(Rs_Consulta_Factura_Clientes!Factura_Referencia) Then Txt_Factura_Referencia.text = Rs_Consulta_Factura_Clientes!Factura_Referencia
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
                    If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("CP")) Then Txt_Codigo_Postal.text = Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                End If
                Txt_Cliente_ID.text = Rs_Consulta_Factura_Clientes!Cliente_ID
                If Not IsNull(Rs_Consulta_Factura_Clientes!Comentarios) Then Txt_Comentarios.text = Rs_Consulta_Factura_Clientes!Comentarios
            Rs_Consulta_Cat_Clientes.Close
            
            If Trim(Rs_Consulta_Factura_Clientes!cancelada) = "N" Then
                Btn_Cancelar.Enabled = True
                Lbl_Facturacion.Caption = "ACTIVA"
                Btn_Enviar_Email.Enabled = True
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
            Btn_Imprimir.Enabled = True
                        
            'Prepara el recordset para consultar el número de factura de la tabla Adm_Descripcion_Facturas
            Mi_SQL = "SELECT * FROM Adm_Notas_credito_Detalles WHERE No_Nota_Credito ='" & Rs_Consulta_Factura_Clientes!No_Nota_Credito & "'"
            Set Rs_Consulta_Descripcion_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Grid_Detalle_NC.Rows = 0
                Grid_Detalle_NC.Cols = 13
                'Pone el encabezado en las columnas
                Grid_Detalle_NC.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) _
                    & "Importe" & Chr(9) & "Producto ID" & Chr(9) & "No Salida" & Chr(9) & "Impuesto" & Chr(9) _
                    & "No_Salida" & Chr(9) & "IVA" & Chr(9) & "Aplica_IVA" & Chr(9) & "Unidad" & Chr(9) & "Incluye"
                'Llenado del grid de la factura consultada
                While Not Rs_Consulta_Descripcion_Facturas.EOF
                    If Not IsNull(Rs_Consulta_Descripcion_Facturas!Unidad) Then
                        Unidad = Trim(Rs_Consulta_Descripcion_Facturas!Unidad)
                    Else
                        Unidad = ""
                    End If
                    Grid_Detalle_NC.AddItem Rs_Consulta_Descripcion_Facturas!Cantidad & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Descripcion & _
                        Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Precio, "###,##0.00") & _
                        Chr(9) & Format(Rs_Consulta_Descripcion_Facturas!Importe, "###,##0.00") & _
                        Chr(9) & Rs_Consulta_Descripcion_Facturas!Producto_ID & Chr(9) & "" & Chr(9) & "" & _
                        Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Unidad & Chr(9) & ""
                    Grid_Detalle_NC.FixedRows = 1
                    Rs_Consulta_Descripcion_Facturas.MoveNext
                Wend
            Formatea_Columnas_Grid
            
            
            'Coloca la información consultada en los controles correspondientes
            Txt_Subtotal.text = Format(Rs_Consulta_Factura_Clientes!SubTotal, "#,##0.00")
            Txt_IVA.text = Format(Rs_Consulta_Factura_Clientes!Iva, "#,##0.00")
            Txt_Total.text = Format(Rs_Consulta_Factura_Clientes!Total, "#,##0.00")
            If Not IsNull(Rs_Consulta_Factura_Clientes!Mensaje_Cancelado) Then
                lbl_estatus_cancel.Visible = True
                lbl_estatus_cancel.Caption = Rs_Consulta_Factura_Clientes!Mensaje_Cancelado
            Else
                lbl_estatus_cancel.Visible = False
                lbl_estatus_cancel.Caption = ""
            End If
            Rs_Consulta_Factura_Clientes.Close
            Rs_Consulta_Descripcion_Facturas.Close
            Fra_Datos_Cliente.Enabled = False
            Fra_Datos_NC.Enabled = False
            Fra_Detalle_NC.Enabled = True
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

    If Txt_No_Nota.text <> "" Then
        'Por default se deshabilita la bandera
        Enviar = False
        Archivos_Adjuntos = ""
        'Valida que existan los archivos a adjuntar
        'Asgna a las variables los nombres de los archivos
        Archivo_PDF = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text) & ".pdf"
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_NC & "\" & Archivo_PDF, "ARCHIVO") = True Then
            Archivo_XML = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_No_Nota.text) & ".xml"
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_NC & "\" & Archivo_XML, "ARCHIVO") = True Then
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
            Archivos_Adjuntos = Ruta_NC & "\" & Archivo_PDF & "|" & Ruta_NC & "\" & Archivo_XML
            
            Mensaje = "<HTML>" & vbNewLine & _
                "<BODY>" & vbNewLine & _
                    "<P>Estimado cliente: <BR></P>" & vbNewLine & _
                    "<P>Por medio de este correo, se le hace entrega de un Comprobante Fiscal Digital por Internet(CFDI).<BR></P>" & vbNewLine & _
                    "<P>Adjunto a este correo encontrará un archivo PDF y un archivo XML correspondientes a su nota de crédito.</P>" & vbNewLine & _
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


