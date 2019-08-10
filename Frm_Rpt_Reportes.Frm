VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_Rpt_Reportes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTES"
   ClientHeight    =   7740
   ClientLeft      =   2895
   ClientTop       =   1875
   ClientWidth     =   17325
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Frm_Rpt_Reportes.frx":0000
   ScaleHeight     =   7740
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox Edm_Reporte 
      Height          =   6915
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   17235
      _ExtentX        =   30401
      _ExtentY        =   12197
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"Frm_Rpt_Reportes.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Pic_Reportes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4005
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   30
      Width           =   9270
      Begin VB.Frame Fra_Reporte_General_Facturas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte General de Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   100
         Top             =   45
         Visible         =   0   'False
         Width           =   4830
         Begin VB.CommandButton Btn_Reporte_General_Facturas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1845
            Picture         =   "Frm_Rpt_Reportes.frx":04C2
            Style           =   1  'Graphical
            TabIndex        =   105
            Tag             =   "A"
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Reporte_General_Facturas 
            Height          =   315
            Left            =   1290
            TabIndex        =   101
            Top             =   1095
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Reporte_General_Facturas 
            Height          =   315
            Left            =   1305
            TabIndex        =   102
            Top             =   600
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   210
            TabIndex        =   104
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   210
            TabIndex        =   103
            Top             =   1155
            Width           =   825
         End
      End
      Begin VB.Frame Fra_Consulta_Cobranza 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consulta de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   4845
         Begin VB.CommandButton Btn_Reporte_Cobranza 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1530
            Picture         =   "Frm_Rpt_Reportes.frx":3A52
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "A"
            Top             =   1710
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Cobranza 
            Height          =   315
            Left            =   1050
            TabIndex        =   22
            Top             =   1065
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39519
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Cobranza 
            Height          =   315
            Left            =   1050
            TabIndex        =   21
            Top             =   720
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39519
         End
         Begin VB.Label Lbl_Fecha_Inicio_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   24
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Lbl_Fecha_Fin_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   1125
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Reporte_Pagos_Cheques 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte pagos con cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   45
         TabIndex        =   51
         Top             =   45
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CommandButton Btn_Reporte_Pagos_Cheques 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1755
            Picture         =   "Frm_Rpt_Reportes.frx":6FE2
            Style           =   1  'Graphical
            TabIndex        =   52
            Tag             =   "A"
            Top             =   1890
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Pagos_Fechas 
            Height          =   315
            Left            =   1245
            TabIndex        =   53
            Top             =   1155
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Pagos_Cheques 
            Height          =   315
            Left            =   1260
            TabIndex        =   54
            Top             =   810
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   56
            Top             =   1215
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   870
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Reporte_Inventarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Inventarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CommandButton Btn_Inventario_General 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1575
            Picture         =   "Frm_Rpt_Reportes.frx":A572
            Style           =   1  'Graphical
            TabIndex        =   60
            Tag             =   "A"
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Tipo_Producto 
            Height          =   315
            Left            =   1320
            TabIndex        =   58
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Lbl_Tipo_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Producto"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   1020
            Width           =   1005
         End
      End
      Begin VB.Frame Fra_Anticipos_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anticipos Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   29
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CommandButton Btn_Reporte_Anticipos_Clientes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1620
            Picture         =   "Frm_Rpt_Reportes.frx":DB02
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "A"
            Top             =   2025
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Estatus_Anticipo 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":11092
            Left            =   840
            List            =   "Frm_Rpt_Reportes.frx":1109F
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1035
            Width           =   3900
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Anticipo_Cliente 
            Height          =   315
            Left            =   2835
            TabIndex        =   33
            Top             =   1380
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852995
            CurrentDate     =   39554
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Anticipo_Cliente 
            Height          =   315
            Left            =   840
            TabIndex        =   32
            Top             =   1380
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852995
            CurrentDate     =   39554
         End
         Begin VB.ComboBox Cmb_Cliente_Anticipo 
            Height          =   315
            Left            =   75
            TabIndex        =   30
            Top             =   690
            Width           =   4665
         End
         Begin VB.Label Lbl_Estatus_Anticipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   45
            TabIndex        =   36
            Top             =   1095
            Width           =   525
         End
         Begin VB.Label Lbl_Fecha_Inicio_Anticipo_Cilente 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   45
            TabIndex        =   35
            Top             =   1440
            Width           =   525
         End
         Begin VB.Label Lbl_Cliente_Anticipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   90
            TabIndex        =   34
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.Frame Fra_Reporte_Ventas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   45
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   5835
         Begin VB.CommandButton Btn_Reporte_Historico_Ventas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   2070
            Picture         =   "Frm_Rpt_Reportes.frx":110BF
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "A"
            Top             =   2070
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Producto_Venta 
            Height          =   315
            Left            =   1155
            TabIndex        =   14
            Top             =   660
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Historico_Ventas 
            Height          =   315
            Left            =   1155
            TabIndex        =   15
            Top             =   1020
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
            MaxDate         =   73415
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Historico_Ventas 
            Height          =   315
            Left            =   1155
            TabIndex        =   16
            Top             =   1470
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   135
            TabIndex        =   42
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   1530
            Width           =   825
         End
      End
      Begin VB.Frame Fra_Facturas_Proveedor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas por Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton Btn_Reporte_Productos_Facturados_Cliente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1575
            Picture         =   "Frm_Rpt_Reportes.frx":1464F
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "A"
            Top             =   2115
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Facturas_Proveedores 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   1245
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin VB.ComboBox Cmb_Facturas_Proveedores 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":17BDF
            Left            =   75
            List            =   "Frm_Rpt_Reportes.frx":17BE1
            TabIndex        =   4
            Top             =   750
            Width           =   4590
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Facturas_Proveedores 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   1665
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Fecha_Final_Proveedor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   1740
            Width           =   705
         End
         Begin VB.Label Lbl_Fecha_Inicio_Proveedor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   1305
            Width           =   870
         End
         Begin VB.Label Lbl_Proveedor_Reporte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   105
            TabIndex        =   3
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.Frame Fra_Reporte_Notas_Credito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte General de Notas de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   121
         Top             =   30
         Visible         =   0   'False
         Width           =   4845
         Begin VB.ComboBox Cmb_Cliente_Nota_Credito 
            Height          =   315
            Left            =   1080
            TabIndex        =   122
            Top             =   240
            Width           =   3630
         End
         Begin VB.CommandButton Btn_Reporte_Notas_Credito 
            Caption         =   "Reporte"
            Height          =   450
            Left            =   1665
            MousePointer    =   99  'Custom
            TabIndex        =   129
            Tag             =   "C"
            Top             =   2400
            Width           =   1500
         End
         Begin VB.CheckBox Chk_Nota_Credito_Activa 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Activa"
            Height          =   150
            Left            =   1140
            TabIndex        =   125
            Top             =   1725
            Width           =   1215
         End
         Begin VB.CheckBox Chk_Nota_Credito_Cancelada 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cancelada"
            Height          =   195
            Left            =   2670
            TabIndex        =   126
            Top             =   1725
            Width           =   1215
         End
         Begin VB.OptionButton Opt_Tipo_Reporte_Nota_Credito 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrupado por Cliente"
            Height          =   195
            Index           =   1
            Left            =   2670
            TabIndex        =   128
            Top             =   2085
            Width           =   1830
         End
         Begin VB.OptionButton Opt_Tipo_Reporte_Nota_Credito 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Listado"
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   127
            Top             =   2085
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Nota_Credito 
            Height          =   315
            Left            =   1080
            TabIndex        =   123
            Top             =   750
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39141
            MaxDate         =   73415
            MinDate         =   32874
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Nota_Credito 
            Height          =   315
            Left            =   1080
            TabIndex        =   124
            Top             =   1275
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39141
            MaxDate         =   73415
            MinDate         =   32874
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   90
            TabIndex        =   132
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   90
            TabIndex        =   131
            Top             =   810
            Width           =   900
         End
         Begin VB.Label Lbl_Fecha_Final 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Final"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   130
            Top             =   1335
            Width           =   825
         End
      End
      Begin VB.Frame Fra_Saldos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Saldo por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   4845
         Begin VB.CheckBox Chk_Imprimir_Saldos_ceros 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Imprimir con saldos cero"
            Height          =   285
            Left            =   1305
            TabIndex        =   120
            Top             =   1710
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.CommandButton Btn_Reporte_Saldos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1755
            Picture         =   "Frm_Rpt_Reportes.frx":17BE3
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "A"
            Top             =   2025
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Cliente_Saldo 
            Height          =   315
            Left            =   180
            TabIndex        =   26
            Top             =   960
            Width           =   4530
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Saldo_Fin 
            Height          =   315
            Left            =   2445
            TabIndex        =   27
            Top             =   1305
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   112852995
            CurrentDate     =   39885
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Saldo_Inicio 
            Height          =   315
            Left            =   180
            TabIndex        =   50
            Top             =   1305
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   112852995
            CurrentDate     =   39885
         End
         Begin VB.Label Lbl_Cliente_Saldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   750
            Width           =   480
         End
      End
      Begin VB.Frame Fra_Reporte_Producto_Por_Documento 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Producto Por Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   45
         TabIndex        =   106
         Top             =   45
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox Cmb_Reporte_ventas_Cliente 
            Height          =   315
            Left            =   180
            TabIndex        =   118
            Top             =   555
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Categorias 
            Height          =   315
            Left            =   1350
            TabIndex        =   116
            Top             =   1845
            Width           =   3285
         End
         Begin VB.ComboBox Cmb_Tipo_Producto_Producto_Por_Documento 
            Height          =   315
            Left            =   1350
            TabIndex        =   114
            Top             =   1485
            Width           =   3285
         End
         Begin VB.CommandButton Btn_Reporte_Producto_Por_Documento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   1665
            Picture         =   "Frm_Rpt_Reportes.frx":1B173
            Style           =   1  'Graphical
            TabIndex        =   113
            Tag             =   "A"
            Top             =   2880
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Producto_Reporte_Producto_Por_Documento 
            Height          =   315
            Left            =   180
            TabIndex        =   107
            Top             =   1140
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Reporte_Producto_Por_Documento 
            Height          =   315
            Left            =   1350
            TabIndex        =   108
            Top             =   2595
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Reporte_Producto_Por_Documento 
            Height          =   315
            Left            =   1350
            TabIndex        =   109
            Top             =   2205
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   180
            TabIndex        =   119
            Top             =   315
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Categoria"
            Height          =   195
            Left            =   180
            TabIndex        =   117
            Top             =   1905
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Producto"
            Height          =   195
            Left            =   195
            TabIndex        =   115
            Top             =   1515
            Width           =   1005
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   195
            TabIndex        =   112
            Top             =   2655
            Width           =   705
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   195
            TabIndex        =   111
            Top             =   2265
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   240
            Left            =   180
            TabIndex        =   110
            Top             =   900
            Width           =   645
         End
      End
      Begin VB.Frame Fra_Historico_Facturas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Historico de Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   4830
         Begin VB.ComboBox Cmb_Cliente_Factura 
            Height          =   315
            Left            =   135
            TabIndex        =   47
            Top             =   630
            Width           =   4590
         End
         Begin VB.CommandButton Btn_Reporte_Historico_Facturas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1620
            Picture         =   "Frm_Rpt_Reportes.frx":1E703
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "A"
            Top             =   2070
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Historico_Factura 
            Height          =   315
            Left            =   1207
            TabIndex        =   11
            Top             =   1365
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Historico_Factura 
            Height          =   315
            Left            =   1215
            TabIndex        =   10
            Top             =   1005
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852992
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Reporte_hitorico_Facturas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   165
            TabIndex        =   48
            Top             =   420
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1425
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1065
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Reporte_Salidas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Salidas de almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   45
         TabIndex        =   69
         Top             =   45
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton Btn_Reporte_Pedido_Factura 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1530
            Picture         =   "Frm_Rpt_Reportes.frx":21C93
            Style           =   1  'Graphical
            TabIndex        =   70
            Tag             =   "A"
            Top             =   1845
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Pedido_Factura 
            Height          =   315
            Left            =   1125
            TabIndex        =   71
            Top             =   1125
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Pedido_Factura 
            Height          =   315
            Left            =   1125
            TabIndex        =   72
            Top             =   540
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39536
         End
         Begin VB.Label Lbl_Fecha_Inicio_Pedido_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   74
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Lbl_Fecha_Fin_Pedido_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   73
            Top             =   1185
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Kardex 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kardex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   61
         Top             =   45
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Producto 
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   615
            Width           =   5610
         End
         Begin VB.CommandButton Btn_Kardex_Reporte 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1935
            Picture         =   "Frm_Rpt_Reportes.frx":25223
            Style           =   1  'Graphical
            TabIndex        =   62
            Tag             =   "A"
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Kardex_Fecha_Final 
            Height          =   315
            Left            =   1440
            TabIndex        =   64
            Top             =   1680
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39448
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker Dtp_Kardex_Fecha_Inicial 
            Height          =   315
            Left            =   1440
            TabIndex        =   65
            Top             =   1200
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39448
            MinDate         =   2
         End
         Begin VB.Label Lbl_Kardex_Fecha_Final 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   285
            TabIndex        =   68
            Top             =   1740
            Width           =   825
         End
         Begin VB.Label Lbl_Kardex_Fecha_Inicial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   285
            TabIndex        =   67
            Top             =   1275
            Width           =   900
         End
         Begin VB.Label Lbl_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   150
            TabIndex        =   66
            Top             =   330
            Width           =   645
         End
      End
      Begin VB.Frame Fra_Pedidos_Facturas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Documnetos por pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   45
         TabIndex        =   94
         Top             =   0
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton Btn_Reporte_Pedidos_Documento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1665
            Picture         =   "Frm_Rpt_Reportes.frx":287B3
            Style           =   1  'Graphical
            TabIndex        =   95
            Tag             =   "A"
            Top             =   1755
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Fin_Reporte_documentos_Pedidos 
            Height          =   315
            Left            =   1080
            TabIndex        =   96
            Top             =   1170
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicial_Reporte_documentos_Pedidos 
            Height          =   315
            Left            =   1080
            TabIndex        =   97
            Top             =   675
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39536
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   99
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   98
            Top             =   720
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Reporte_Entradas_Salidas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entradas y Salidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   5835
         Begin VB.ComboBox Cmb_Proveedor_Entrada 
            Height          =   315
            Left            =   840
            TabIndex        =   89
            Top             =   750
            Width           =   3615
         End
         Begin VB.CommandButton Btn_Reporte_Entradas_Salidas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1710
            Picture         =   "Frm_Rpt_Reportes.frx":2BD43
            Style           =   1  'Graphical
            TabIndex        =   88
            Tag             =   "A"
            Top             =   1980
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Entrada_Salida_Inicio 
            Height          =   330
            Left            =   840
            TabIndex        =   90
            Top             =   1140
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852995
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Entrada_Salida_Fin 
            Height          =   330
            Left            =   2670
            TabIndex        =   91
            Top             =   1140
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   112852995
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Fecha_Inicio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   75
            TabIndex        =   93
            Top             =   1215
            Width           =   525
         End
         Begin VB.Label Lbl_Proveedor_Entrada 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   75
            TabIndex        =   92
            Top             =   810
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame Fra_Reporte_Pedidos_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedidos por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   45
         TabIndex        =   75
         Top             =   45
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox Cmb_Reporte_Pedido_Producto 
            Height          =   315
            Left            =   195
            TabIndex        =   79
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox Txt_Reporte_No_Pedido 
            Height          =   315
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   78
            Top             =   585
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton Btn_Reporte_Pedidos_Cliente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
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
            Left            =   1770
            Picture         =   "Frm_Rpt_Reportes.frx":2F2D3
            Style           =   1  'Graphical
            TabIndex        =   77
            Tag             =   "A"
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Estatus_Banco 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":32863
            Left            =   1365
            List            =   "Frm_Rpt_Reportes.frx":32870
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1710
            Width           =   3240
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Pedidos_Productos 
            Height          =   315
            Left            =   1350
            TabIndex        =   80
            Top             =   1335
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Pedido_Producto 
            Height          =   315
            Left            =   1350
            TabIndex        =   81
            Top             =   945
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852992
            CurrentDate     =   39513
         End
         Begin VB.Label Lbl_Fecha_Fin_Pedidos_Clientes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   150
            TabIndex        =   86
            Top             =   1395
            Width           =   705
         End
         Begin VB.Label Lbl_Fecha_Inicio_Pedidos_Clientes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   150
            TabIndex        =   85
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label Lbl_Reporte_Pedido_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   195
            TabIndex        =   84
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Lbl_Reporte_Pedido_No_Pedido 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Pedido"
            Height          =   195
            Left            =   210
            TabIndex        =   83
            Top             =   690
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Llb_Estatus_Banco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
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
            TabIndex        =   82
            Top             =   1755
            Width           =   555
         End
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
      Left            =   15600
      Picture         =   "Frm_Rpt_Reportes.frx":3288A
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "A"
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Exportar_Excel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar"
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
      Left            =   7882
      Picture         =   "Frm_Rpt_Reportes.frx":35F89
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "A"
      Top             =   7020
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Imprimir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
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
      Left            =   135
      Picture         =   "Frm_Rpt_Reportes.frx":387C3
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "A"
      Top             =   7020
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Reporte 
      Height          =   270
      Left            =   3630
      TabIndex        =   1
      Top             =   5475
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   476
      _Version        =   393216
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      Appearance      =   0
   End
End
Attribute VB_Name = "Frm_Rpt_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngFormWidth As Long        'Ancho de la forma
Private lngFormHeight As Long       'Largo de la forma
Dim Ancho_Frame As Double
Dim Largo_Frame As Double
Dim Despliega_Lista As Long
Private Lineas_Encabezado As Integer                '# de lineas del encabezado

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Saldos_Cliente
'DESCRIPCIÓN            : Muestra el saldo de lo clientes
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Saldos_Cliente()
Dim Cliente As String                                       'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Remisiones_Clientes As rdoResultset     'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_Factura As Double                                 'Almacena el total de la facura
Dim Total_Saldo As Double                                   'Almacena el total del saldo
Dim Total_Abono As Double                                   'Almacena el total de los abonos
Dim Fecha_Vencimiento As Date                               'Almacena la fecha de vencimiento de la factura
Dim Vencimiento As Integer                                  'Almacena los dias de vnecimiento de la factura
Dim Total_Vencimiento As Integer                            'Almacena los espacios para colocar la información
Dim Espacio As String
Dim Total_0_30 As Double
Dim Total_30_60 As Double
Dim Total_60_90 As Double
Dim Total_90 As Double
Dim Total_Cliente As Double
Dim Abono_Cliente As Double
Dim Saldo_Cliente As Double
Dim Cliente_ID As String
Dim Tipo_Pago As String
Dim Tem_Cliente As String

On Error GoTo handler

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE SALDO POR CLIENTES", Len("***************************************************************************************************************"))
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    

    Mi_SQL = " SELECT * FROM  Cat_clientes "
    If Cmb_Cliente_Saldo.ListIndex > -1 Then
         Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
    End If
    Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Adm_Remisiones_Clientes.EOF
        Mi_SQL = " SELECT Distinct Cat_Clientes.Cliente_ID, Cat_Clientes.* "
        Mi_SQL = Mi_SQL & " From Cat_Clientes, Adm_Clientes_Facturas "
        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID='" & Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID = Adm_Clientes_Facturas.Cliente_ID AND  Adm_Clientes_Facturas.Pagada = 'N' "
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cancelada = 'N' "
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Total > 0 "
        Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " UNION "
        Mi_SQL = Mi_SQL & "  SELECT Distinct Cat_Clientes.Cliente_ID,Cat_Clientes.* "
        Mi_SQL = Mi_SQL & " From Adm_Clientes_Remisiones, Cat_Clientes "
        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID='" & Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID = Adm_Clientes_Remisiones.Cliente_ID AND  Adm_Clientes_Remisiones.Pagada = 'N' "
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cancelada = 'N' "
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Total > 0 "
        Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre "
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Clientes.EOF Then
            With Rs_Consulta_Cat_Clientes
                If Cliente_ID <> Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") Then
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Ciudad) Then Ciudad = Rs_Consulta_Cat_Clientes!Ciudad
                    If Cliente_ID <> "" Then
                        Print #1, "_________________________________________________________________________________________________________________________"
                        Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
                                  Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
                                  Spc(3); Conectar_Ayudante.Alinea_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
                        Print #1,
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                        Grid_Reporte.AddItem ""
                    End If
                    Print #1,
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
                    ''Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion"); " "; .rdoColumns("Colonia")
                    ''Print #1, Spc(2); "Ciudad:    "; Ciudad; Spc(2); "CP: "; .rdoColumns("CP"); Spc(2); "Telefono: "; .rdoColumns("Telefono")
                    ''Print #1, Spc(2); "Dias Credito:    "; .rdoColumns("Dias_Credito")
                    Print #1, Spc(2); "(*) Documentos de contado"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                    Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion") & Chr(9) & .rdoColumns("Colonia")
                    Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP") & Chr(9) & "Telefono" & Chr(9) & .rdoColumns("Telefono")
                    Grid_Reporte.AddItem "Dias Credito:" & Chr(9) & .rdoColumns("Dias_Credito")
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " Documento   Fecha         Total        Abono          Saldo           1-30         31-60         61-90          +90     "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Documento" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "0-30" & Chr(9) & "30-60" & Chr(9) & "60-90" & Chr(9) & "'+90"
                    Cliente_ID = Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")
                    Total_Cliente = 0
                    Saldo_Cliente = 0
                    Abono_Cliente = 0
                End If
            End With
        End If
        'SE CONSULTAN LAS FACTURAS
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Facturas.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Pagada='N'"
        Mi_SQL = Mi_SQL & " AND Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Total > 0 "
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cliente_ID='" & Format(Rs_Consulta_Adm_Remisiones_Clientes!Cliente_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Facturas.Cliente_ID"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
            With Rs_Consulta_Adm_Facturas_Clientes
                If Trim(.rdoColumns("Tipo_Pago")) = "CONTADO" Then
                    Tipo_Pago = "*"
                Else
                    Tipo_Pago = " "
                End If
                Fecha_Vencimiento = Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha_Pago"), "MM/dd/yyyy")
                Vencimiento = DateDiff("d", Format(Fecha_Vencimiento, "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                If Vencimiento <= 30 And Vencimiento > 0 Then
                    Total_0_30 = Total_0_30 + .rdoColumns("Saldo")
                    Total_Vencimiento = 2
                    Espacio = ""
                Else
                    If Vencimiento >= 31 And Vencimiento <= 60 Then
                        Total_30_60 = Total_30_60 + .rdoColumns("Saldo")
                        Total_Vencimiento = 16
                        Espacio = Chr(9)
                    Else
                        If Vencimiento >= 61 And Vencimiento <= 90 Then
                            Total_60_90 = Total_60_90 + .rdoColumns("Saldo")
                            Total_Vencimiento = 30
                            Espacio = Chr(9) & Chr(9)
                        Else
                            If Vencimiento >= 91 Then
                                Total_90 = Total_90 + .rdoColumns("Saldo")
                                Total_Vencimiento = 44
                                Espacio = Chr(9) & Chr(9) & Chr(9)
                            End If
                        End If
                    End If
                End If
                'VALIDA SI VA A IMPRIMIR FACTURAS VENCIDAS
                If Vencimiento > 0 Then
                    Print #1, Format(.rdoColumns("No_Factura_Electronica"), "00000000"); Spc(1); Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                    Spc(Total_Vencimiento); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Espacio & _
                    Format(.rdoColumns("Saldo"), "#,###,##0.00")
                Else
                    Print #1, Format(.rdoColumns("No_Factura_Electronica"), "00000000"); Spc(1); Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                    Spc(Total_Vencimiento)
                    Grid_Reporte.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Saldo"), "#,###,##0.00")
                End If
                Total_Factura = Total_Factura + .rdoColumns("Total")
                Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                Total_Abono = Total_Abono + .rdoColumns("Abono")
                Total_Cliente = Total_Cliente + .rdoColumns("Total")
                Saldo_Cliente = Saldo_Cliente + .rdoColumns("Saldo")
                Abono_Cliente = Abono_Cliente + .rdoColumns("Abono")
            End With
            Rs_Consulta_Adm_Facturas_Clientes.MoveNext
        Wend
        Rs_Consulta_Adm_Facturas_Clientes.Close
        If Not Trim(Tem_Cliente) = Trim(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")) Then
            'SE ASIGNA EL CLIENTE DEL QUE SE ESTA HACIENDO LA CONSULTA
            Tem_Cliente = Trim(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID"))
            'SE AGREGAN LAS REMISIONNES
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones,Cat_Clientes "
            Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Remisiones.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Pagada='N'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Total > 0 "
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cliente_ID='" & Format(Rs_Consulta_Adm_Remisiones_Clientes!Cliente_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Remisiones.Cliente_ID"
            Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                With Rs_Consulta_Adm_Facturas_Clientes
                    If Trim(.rdoColumns("Tipo_Pago")) = "CONTADO" Then
                        Tipo_Pago = "*"
                    Else
                        Tipo_Pago = " "
                    End If
                    Fecha_Vencimiento = Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha_Pago"), "MM/dd/yyyy")
                    Vencimiento = DateDiff("d", Format(Fecha_Vencimiento, "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                    If Vencimiento <= 30 And Vencimiento > 0 Then
                        Total_0_30 = Total_0_30 + .rdoColumns("Saldo")
                        Total_Vencimiento = 2
                        Espacio = ""
                    Else
                        If Vencimiento >= 31 And Vencimiento <= 60 Then
                            Total_30_60 = Total_30_60 + .rdoColumns("Saldo")
                            Total_Vencimiento = 16
                            Espacio = Chr(9)
                        Else
                            If Vencimiento >= 61 And Vencimiento <= 90 Then
                                Total_60_90 = Total_60_90 + .rdoColumns("Saldo")
                                Total_Vencimiento = 30
                                Espacio = Chr(9) & Chr(9)
                            Else
                                If Vencimiento >= 91 Then
                                    Total_90 = Total_90 + .rdoColumns("Saldo")
                                    Total_Vencimiento = 44
                                    Espacio = Chr(9) & Chr(9) & Chr(9)
                                End If
                            End If
                        End If
                    End If
                    'VALIDA SI HAY REMISIONES VENCIDAS
                    If Vencimiento > 0 Then
                        Print #1, .rdoColumns("No_Remision"); Spc(1); Empresa; Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                        Spc(Total_Vencimiento); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12)
                        Grid_Reporte.AddItem .rdoColumns("No_Remision") & _
                        Empresa & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Espacio & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00")
                    Else
                        Print #1, .rdoColumns("No_Remision"); Spc(1); Empresa; Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                        Spc(Total_Vencimiento)
                        Grid_Reporte.AddItem .rdoColumns("No_Remision") & _
                        Empresa & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00")
                    End If
                    Total_Factura = Total_Factura + .rdoColumns("Total")
                    Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                    Total_Abono = Total_Abono + .rdoColumns("Abono")
                    Total_Cliente = Total_Cliente + .rdoColumns("Total")
                    Saldo_Cliente = Saldo_Cliente + .rdoColumns("Saldo")
                    Abono_Cliente = Abono_Cliente + .rdoColumns("Abono")
                End With
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
            Rs_Consulta_Adm_Facturas_Clientes.Close
        End If
        Rs_Consulta_Adm_Remisiones_Clientes.MoveNext
    Wend
    Rs_Consulta_Adm_Remisiones_Clientes.Close
    Print #1, "_________________________________________________________________________________________________________________________"
    Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
    Spc(3); Conectar_Ayudante.Alinea_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
    Print #1,
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
    Grid_Reporte.AddItem ""
    Print #1, "_________________________________________________________________________________________________________________________"
    Print #1, " TOTALES DOCUMENTOS"; Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Factura, "#,###,##0.00"), 13); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Abono, "#,###,##0.00"), 12); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Saldo, "#,###,##0.00"), 13); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_0_30, "#,###,##0.00"), 12); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_30_60, "#,###,##0.00"), 12); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_60_90, "#,###,##0.00"), 12); _
    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_90, "#,###,##0.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "TOTALES DOCUMENTOS" & Chr(9) & Chr(9) & Format(Total_Factura, "#,###,##0.00") & Chr(9) & Format(Total_Abono, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00") & Chr(9) & Format(Total_0_30, "#,###,##0.00") & Chr(9) & _
    Format(Total_30_60, "#,###,##0.00") & Chr(9) & Format(Total_60_90, "#,###,##0.00") & Chr(9) & Format(Total_90, "#,###,##0.00")
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 17415
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description
End Sub
Private Sub Btn_Exportar_Excel_Click()
Dim Columna As Integer     'Indica que columna del grid es la que se esta consultando
Dim Fila As Long           'Indica que fila del grid es la que se esta consultando
Dim Ruta_Archivo As String 'Obtiene la rura del archivo en donde se desea guardar
Dim CABECERA As String     'Indica el texto a exportar

On Error GoTo ErrHandler
    'Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    'Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    'Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.XLS|"
    'Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    'Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    'Display name of selected file
    Ruta_Archivo = MDIFrm_Apl_Principal.CommonDialog1.fileName
    Open Ruta_Archivo For Output As #1
        For Fila = 0 To Grid_Reporte.Rows - 1
            For Columna = 0 To Grid_Reporte.Cols - 1
                If Columna > 0 Then
                    CABECERA = CABECERA & Chr(9) & Grid_Reporte.TextMatrix(Fila, Columna)
                Else
                    CABECERA = Grid_Reporte.TextMatrix(Fila, Columna)
                End If
            Next Columna
            Print #1, CABECERA
            CABECERA = ""
        Next Fila
    Close #1
    MsgBox "Reporte Exportado a Excel", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Ha ocurrido un error en la exportación, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub
Private Sub Btn_Imprimir_Click()
''Dim Linea As String 'Obtiene el texto a imprimir
''
''    'Printer.PaperSize = vbPRPSLetter
''    'Printer.PaperBin = 8
''    'Printer.Orientation = vbPRORPortrait
''    Printer.FontSize = 8
''    Printer.Font = "COURIER NEW"
''    Printer.Print
''    Printer.FontSize = 11
''    Printer.Font = "COURIER NEW"
''    Printer.Print
''    Printer.FontSize = 8
''    Printer.Font = "Courier New"
''    'Valida que exista la ruta si no la crea
''    If Len(Dir(App.Path & "\REP\Reportes.txt", vbDirectory)) = 0 Then
''        MkDir App.Path & "\REP\Reportes.txt"
''    End If
''    Open App.Path & "\REP\Reportes.txt" For Input As #1
''    Do While Not EOF(1)
''        Line Input #1, Linea
''        Printer.Print Linea
''    Loop
''    Printer.EndDoc
''    Close #1
''    MsgBox "Reporte enviado a impresora", vbInformation
    
Dim linea As String 'Obtiene el texto a imprimir
Dim Impresora As String            'Tomna el nombre la impresora
Dim Mi_Impresora As Printer        'Toma el nombre de la impresora
Dim Ubicacion_Impresora As String  'Toma el valor de la ubicacion dela impresora
Dim Encabezado As Boolean          'Identifica si se ha encontrado el encabezado
Dim Lineas_Impresas As Integer     'Contador de lineas impresas
Dim Primer_Paso As Boolean         'Indica si es el inicio del archivop
Dim Escribir_Archivo As Boolean    'Indica si se escribe en el archivo temp
Dim Imprime_No_Pagina As Boolean   'Se imprime el no de pagina
Dim Control_Encabezado As Integer  'Control del encabezado
Const TW As Integer = 0
Dim Lineas_Encabezado As Integer


    Printer.FontSize = 8
    Printer.Font = "Courier New"
    
On Error GoTo handler
    Me.MousePointer = 11
    Primer_Paso = True
    Escribir_Archivo = True
    Lineas_Impresas = 0
    Imprime_No_Pagina = True
    'se va acambiar de impresaora y se Crea el encabezado
    Debug.Print Printer.DeviceName
    Debug.Print Printer.Page
    Printer.FontSize = 8
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 8
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 8
    Printer.Font = "Courier New"
    Printer.CurrentX = 0.01 * TW
    Open App.Path & "\REP\Reportes.txt" For Input As #1
    Do While Not EOF(1)
        'Localiza el inicio del archivo y la primera fecha encontrada delimitando el encabezado
        'Lee la primera linea
        Line Input #1, linea
        If linea = "NP" Then
            GoTo 6000
        End If
        If Imprime_No_Pagina = True Then
            If Len(linea) > 0 Then
                Printer.CurrentX = 0.01 * TW
                Printer.Print Spc(2); "Pag #" & Conectar_Ayudante.Alinea_Derecha(Printer.Page, 3)
                Printer.CurrentX = 0.01 * TW
                Printer.Print Spc(2); linea
                Imprime_No_Pagina = False
            Else
                Printer.CurrentX = 0.01 * TW
                Printer.Print Spc(2); linea
            End If
        Else
            Printer.CurrentX = 0.01 * TW
            Printer.Print Spc(2); linea
        End If
        'Se crea el archivo de encabezado para los reportes
        'Printer.Print Linea
        'Imprime el numero de pagina
        If Primer_Paso = True Then
            Open App.Path & "\Temp.txt" For Output As #2
            Primer_Paso = False
        End If
        Control_Encabezado = Control_Encabezado + 1
        If Escribir_Archivo = True And Lineas_Encabezado >= Control_Encabezado Then
            Printer.CurrentX = 0.01 * TW
            Print #2, Spc(2); linea
        End If
        Debug.Print linea
        If Lineas_Encabezado < Control_Encabezado _
          And Escribir_Archivo = True Then
            Close #2
            Escribir_Archivo = False
        End If
        'se incrementa el contador de lineas
        Lineas_Impresas = Lineas_Impresas + 1
        If Lineas_Impresas >= 89 Then
6000:       Imprime_No_Pagina = True
            If Lineas_Impresas < 79 Then
                Lineas_Impresas = Lineas_Impresas + 2
                For Lineas_Impresas = Lineas_Impresas To 60
                    Printer.CurrentX = 0.01 * TW
                    Printer.Print
                Next Lineas_Impresas
            End If
            Printer.FontSize = 8
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 8
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 8
            Printer.Font = "Courier New"
            Printer.CurrentX = 0.01 * TW
            'Imprime el encabezado
            Open App.Path & "\Temp.txt" For Input As #2
            Do While Not EOF(2)
                Line Input #2, linea
                If Imprime_No_Pagina = True Then
                    If Len(linea) > 0 Then
                        Printer.CurrentX = 0.01 * TW
                        Printer.Print Spc(2); " Pag #" & Conectar_Ayudante.Alinea_Derecha(Printer.Page, 3)
                        Printer.Print linea
                        Imprime_No_Pagina = False
                    Else
                        Printer.CurrentX = 0.01 * TW
                        Printer.Print linea
                    End If
                Else
                    Printer.CurrentX = 0.01 * TW
                    Printer.Print linea
                End If
                Debug.Print linea
            Loop
            Lineas_Impresas = Lineas_Encabezado
            Close #2
        End If
    Loop
    Printer.EndDoc
    Close #1
    Me.MousePointer = 0
    MsgBox "Reporte enviado a impresora"
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Btn_Inventario_General_Click()
    Me.MousePointer = 11
        Call Reporte_Inventario_General
    Me.MousePointer = 0
End Sub

Private Sub Btn_Kardex_Reporte_Click()
    Me.MousePointer = 11
        Call Kardex_Materia_Prima
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Anticipos_Clientes_Click()
    Me.MousePointer = 11
        Reporte_Anticipos_Proveedores   'Reporte de anticipos proveedores
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Click()

End Sub

Private Sub Btn_Reporte_Cobranza_Click()
    Me.MousePointer = 11
        Select Case Fra_Consulta_Cobranza.Caption
            Case "Reporte de Cobranza"
                Call Reporte_Cobranza
            End Select
    Me.MousePointer = 0
End Sub






'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Ventas_Clientes
'DESCRIPCIÓN            : Consulta las ventas
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Ventas_Clientes()
Dim Mi_SQL As String
Dim Total_Ventas As Double                              'Indica el total de las ventas obtenidas
Dim Gran_total As Double
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Ciudad As String                                    'Almacena el nombre de la ciudad
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Rs_Consulta_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Rs_Consulta_Detalles_Facturas As rdoResultset
Dim Encadezado_Reporte As Boolean

On Error GoTo handler

    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE VENTAS", Len("***************************************************************************************************************"))
    Mi_SQL = "SELECT * FROM Cat_Clientes "
    If Cmb_Producto_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.ListIndex), "00000") & "' "
    If Encadezado_Reporte = False Then
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Codigo" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
        Encadezado_Reporte = True
    End If
    Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Gran_total = 0
    While Not Rs_Consulta_Clientes.EOF
        With Rs_Consulta_Clientes
            Ciudad = Rs_Consulta_Clientes.rdoColumns("Ciudad")
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
            Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
            Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
        End With
        Total_Ventas = 0
        'Consulta las facturas del cliente y las fechas seleccionadas
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Fecha"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   No. Factura        Fecha              Importe        Estatus"
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem "FACTURAS"
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem ""
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Sin Pagar"
                    End If
                End If
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                    Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                        Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                        Spc(6); Estatus
                    Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                    Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                    Gran_total = Gran_total + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                Else
                    Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                        Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(5); Conectar_Ayudante.Alinea_Derecha("0.00", 14); Spc(6); Estatus
                    Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "0.00" & Chr(9) & Estatus
                End If
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Factura        Fecha              Importe        Estatus"
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                    Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                        Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                        Spc(6); Estatus
                Else
                    Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                        Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(5); Conectar_Ayudante.Alinea_Derecha("0.00", 14); Spc(6); Estatus
                End If
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                'Consulta el detalle de la factura
                Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas_Detalles"
                Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & "'"
                Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                While Not Rs_Consulta_Detalles_Facturas.EOF
                    Print #1, Spc(1); Rs_Consulta_Detalles_Facturas.rdoColumns("Producto_ID"); Spc(15 - Len(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto_ID"))); Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Descripcion"), Chr(13)), 1, 50); Spc(55 - Len(Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Descripcion"), Chr(13)), 1, 50))); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0"), 9); Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe"), "#,###,##0.00"), 13)
                    Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID") _
                        & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre") _
                        & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") _
                        & Chr(9) & "F -" & Empresa _
                        & Chr(9) & Estatus _
                        & Chr(9) & Rs_Consulta_Detalles_Facturas.rdoColumns("Producto_ID") _
                        & Chr(9) & Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Descripcion"), Chr(10)), Chr(13)) _
                        & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0") _
                        & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Precio"), "#,###,##0.00") _
                        & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe"), "#,###,##0.00")
                    Total_Ventas = Total_Ventas + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                    Gran_total = Gran_total + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                    Rs_Consulta_Detalles_Facturas.MoveNext
                Wend
                Rs_Consulta_Detalles_Facturas.Close
                
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
        Else
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   No. Factura        Fecha              Importe        Estatus                                                         "
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   SIN FACTURAS                                                                                                         "
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem "SIN FACTURAS"
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1,
        Print #1,
        Print #1, Spc(35); "-------------------"
        Print #1, Spc(18); "Total Venta"; Spc(6); Conectar_Ayudante.Alinea_Derecha(Format(Total_Ventas, "#,###,##0.00"), 14)
        Print #1,
        If Encadezado_Reporte = False Then
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Venta" & Chr(9) & Format(Total_Ventas, "#,###,##0.00")
            Encadezado_Reporte = True
        End If
        'Consulta las remisiones del cliente y las fechas seleccionadas
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Fecha_Salida BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Fecha_Salida"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
                Print #1,
                Print #1, Conectar_Ayudante.Centrar_Texto("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Remisiones     Fecha              Importe        Estatus                                                         "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem "REMISIONES"
                Grid_Reporte.AddItem "No. Remisiones" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
                If Encadezado_Reporte = False Then
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Codigo" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                    Encadezado_Reporte = True
                End If
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Sin Pagar"
                    End If
                End If
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                        Gran_total = Gran_total + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha("0.00", 14); Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                
                    Print #1,
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "   No. Remision       Fecha              Importe        Estatus"
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha("0.00", 14); Spc(6); Estatus
                    End If
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    'Consulta el detalle de la factura
                    Mi_SQL = "SELECT * FROM Adm_Detalles_Remisiones"
                    Mi_SQL = Mi_SQL & " WHERE No_Remision='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & "'"
                    Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")
                    Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    While Not Rs_Consulta_Detalles_Facturas.EOF
                        Print #1, Spc(1); Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"); Spc(15 - Len(Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"))); Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50); Spc(55 - Len(Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50))); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0"), 9); Spc(5); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe") + Rs_Consulta_Detalles_Facturas.rdoColumns("IVA"), "#,###,##0.00"), 13)
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID") _
                            & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") _
                            & Chr(9) & "R -" & Empresa _
                            & Chr(9) & Estatus _
                            & Chr(9) & Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo") _
                            & Chr(9) & Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(10)), Chr(13)) _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Precio"), "#,###,##0.00") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe"), "#,###,##0.00")
                        Total_Ventas = Total_Ventas + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Gran_total = Gran_total + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Rs_Consulta_Detalles_Facturas.MoveNext
                    Wend
                    Rs_Consulta_Detalles_Facturas.Close
                
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
        Else
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   No. Remision       Fecha              Importe        Estatus"
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   SIN REMISIONES                                                                                                         "
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem "SIN REMISIONES"
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        Rs_Consulta_Clientes.MoveNext
    Wend
    Rs_Consulta_Clientes.Close
    Print #1, Spc(35); "-------------------"
    Print #1, Spc(18); "Gran Total "; Spc(6); Conectar_Ayudante.Alinea_Derecha(Format(Gran_total, "#,###,##0.00"), 14)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Venta" & Chr(9) & Format(Gran_total, "#,###,##0.00")
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Reporte_Pedidos
'DESCRIPCION            : Muestra el reporte de los pedidos
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Reporte_Pedidos()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset        'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Ciudad As String                            'Almacena el nombre de la ciudad
Dim Total As Double                             'Almacena el total del pedido
Dim Total_Producto As Double                    'Almacena el total de productos
Dim Productos_Pedido As Double                  'Almacena el total del producto por pedido

On Error GoTo handler

    'Coloca el encabezado de los reportes
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE PEDIDOS CLIENTES", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("del " & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy HH:mm:ss") & " al " & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
    Print #1, ""
    Grid_Reporte.Rows = 0
    Grid_Reporte.Clear
    Grid_Reporte.AddItem ""
    'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
    Mi_SQL = "SELECT Ope_Pedidos_Detalles.Estatus, Ope_Pedidos.Pedido_ID,Ope_Pedidos.Total,Fecha_Pedido,Ope_Pedidos.Cliente_ID,Ope_Pedidos_Detalles.Producto_ID,Ope_Pedidos_Detalles.Descripcion, Ope_Pedidos_Detalles.Cantidad,Ope_Pedidos_Detalles.Importe,Cat_Productos.Clave,Cat_Productos.Producto_ID"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Ope_Pedidos_Detalles,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.Pedido_ID=Ope_Pedidos_Detalles.Pedido_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID=Ope_Pedidos_Detalles.Producto_ID"
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Estatus_Banco.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Ope_Pedidos_Detalles.Estatus ='" & Trim(Cmb_Estatus_Banco.text) & "'"
    If Cmb_Reporte_Pedido_Producto.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Cliente_ID='" & Format(Cmb_Reporte_Pedido_Producto.ItemData(Cmb_Reporte_Pedido_Producto.ListIndex), "00000") & "'"
    If Trim(Txt_Reporte_No_Pedido.text) <> "" Then Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Pedido_ID='" & Format(Trim(Txt_Reporte_No_Pedido.text), "0000000000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.Fecha_Pedido,Ope_Pedidos.Pedido_ID"
    Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Pedidos
        While Not .EOF
            If No_Pedido <> .rdoColumns("Pedido_ID") Then
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID = '" & Rs_Consulta_Pedidos.rdoColumns("Cliente_ID") & "'"
            Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Clientes
                Ciudad = Rs_Consulta_Clientes.rdoColumns("Ciudad")
                Print #1,
                If Productos_Pedido > 0 Then
                    Print #1, Spc(90); "Total Productos:  "; Format(Productos_Pedido, "###,###")
                    Grid_Reporte.AddItem "Total Producto:" & Chr(9) & Format(Productos_Pedido, "###,###")
                    Productos_Pedido = 0
                    Print #1,
                    Grid_Reporte.AddItem ""
                End If
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Print #1, Spc(2); "Cliente    :"; Trim(Rs_Consulta_Clientes.rdoColumns("Nombre"))
                Print #1, Spc(2); "Direccion  :"; Trim(.rdoColumns("Direccion"))
                Print #1, Spc(2); "Colonia    :"; Trim(.rdoColumns("Colonia")); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad     :"; Trim(Ciudad)
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente  :" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre")
                Grid_Reporte.AddItem "Direccion:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia  :" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad   :" & Chr(9) & Ciudad
            End With
                'Imprime los pedidos
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(1); "No. Pedido : "; .rdoColumns("Pedido_ID"); Spc(5); "Fecha : "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                Grid_Reporte.AddItem "Pedido :" & Chr(9) & .rdoColumns("Pedido_ID") & Chr(9) & "Fecha :" & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Total Pedido: " & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, " Codigo                 Descripcion                                                                            Cantidad "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad"
                Print #1, "________________________________________________________________________________________________________________________"
                Print #1,
                No_Pedido = .rdoColumns("Pedido_ID")
            End If
            Print #1, Spc(2); .rdoColumns("Clave"); Spc(22 - Len(.rdoColumns("Clave"))); Mid(.rdoColumns("Descripcion"), 1, 59); Spc(83 - Len(Mid(.rdoColumns("Descripcion"), 1, 75))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "###,###"), 10)
            Grid_Reporte.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Cantidad")
            Productos_Pedido = Productos_Pedido + .rdoColumns("Cantidad")
            Total_Producto = Total_Producto + .rdoColumns("Cantidad")
            .MoveNext
        Wend
    End With
    Rs_Consulta_Pedidos.Close
    Grid_Reporte.AddItem "Total Productos:" & Chr(9) & Format(Productos_Pedido, "###,###")
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(97); "Total Productos: "; Format(Total_Producto, "###,###")
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Productos:" & Chr(9) & Format(Total_Producto, "###,###")
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub




'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Anticipos_Proveedores
'DESCRIPCIÓN            : Muestra el reporte de los anticipos de proveedores
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 31-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Anticipos_Proveedores()
Dim Mi_SQL As String
Dim Rs_Consulta_Anticipos As rdoResultset
Dim Rs_Consulta_Movimientos As rdoResultset
Dim Comentarios_Presupuestos As String
Dim Cantidad_Cheque As Double
Dim Proveedor_ID As String
Dim Respuesta As String
Dim Estatus_Anticipo As String
Dim Pago As Double

    
On Error GoTo handler
    
    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    'Busca los anticipos del proveedor
    Mi_SQL = "SELECT Adm_Proveedores_Anticipos.*,Cat_Proveedores.Nombre AS Proveedor,Cat_Proveedores.Direccion,Cat_Proveedores.Colonia,Cat_Proveedores.Codigo_Postal,Cat_Proveedores.Ciudad,Cat_Bancos.Nombre AS Banco"
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Anticipos,Cat_Proveedores,Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Proveedores_Anticipos.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Anticipos.Banco_ID=Cat_Bancos.Banco_ID"
    If Cmb_Cliente_Anticipo.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Anticipos.Proveedor_ID='" & Format(Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.ListIndex), "00000") & "'"
    End If
    If Cmb_Estatus_Anticipo.ListIndex > 0 Then
        If Cmb_Estatus_Anticipo.text = "APLICADO" Then
            Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Anticipos.Aplicado='S'"
        Else
            Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Anticipos.Aplicado='N'"
        End If
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Anticipos.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Proveedores_Anticipos.Proveedor_ID"
    Set Rs_Consulta_Anticipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Anticipos.EOF Then
        Open App.Path & "\REP\Reportes.txt" For Output As #1
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
        Print #1,
        Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE ANTICIPOS DE PROVEEDORES", Len("***************************************************************************************************************"))
        While Not Rs_Consulta_Anticipos.EOF
            If Proveedor_ID <> Rs_Consulta_Anticipos.rdoColumns("Proveedor_ID") Then
                Print #1, ""
                Print #1, "________________________________________________________________________________________________________________________"
                Print #1, Spc(2); "Proveedor: "; Rs_Consulta_Anticipos.rdoColumns("Proveedor")
                Print #1, Spc(2); "Direccion: "; Rs_Consulta_Anticipos.rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; Rs_Consulta_Anticipos.rdoColumns("Colonia"); Spc(2); "CP: "; Rs_Consulta_Anticipos.rdoColumns("Codigo_Postal")
                Print #1, Spc(2); "Ciudad:    "; Rs_Consulta_Anticipos.rdoColumns("Ciudad")
                Print #1,
                Print #1, "  Fecha   Anticipo     Banco          F.Pago Referencia     Concepto                      Cantidad     Saldo   Estatus  "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Proveedor:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Proveedor")
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Codigo_Postal")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Ciudad")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Fecha" & Chr(9) & "Anticipo" & Chr(9) & "Banco" & Chr(9) & "Forma Pago" & Chr(9) & "Concepto" & Chr(9) & "Cantidad" & Chr(9) & "Saldo" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
                Proveedor_ID = Rs_Consulta_Anticipos.rdoColumns("Proveedor_ID")
            End If
            If IsNull(Rs_Consulta_Anticipos.rdoColumns("Pago")) Then
                Pago = 0
            Else
                Pago = Rs_Consulta_Anticipos.rdoColumns("Pago")
            End If
            If Val(Rs_Consulta_Anticipos.rdoColumns("Total")) > Pago Then
                Estatus_Anticipo = "PENDIENTE"
            Else
                Estatus_Anticipo = "APLICADO"
            End If
            Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(1); Conectar_Ayudante.Alinea_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                Spc(2); Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18); _
                Spc(20 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")), 1, 26); _
                Spc(23 - Len(Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")), 1, 26))); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total"), "#,###,##0.00"), 12); _
                Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Pago, "#,###,##0.00"), 12); _
                Spc(1); Estatus_Anticipo
            Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")) _
                & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total"), "#,###,##0.00") _
                & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Pago, "#,###,##0.00") _
                & Chr(9) & Estatus_Anticipo
        Rs_Consulta_Anticipos.MoveNext
        Wend
        Close #1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
    Else
        MsgBox "No se encontraron registros con la búsqueda especificada", vbExclamation
    End If
    Rs_Consulta_Anticipos.Close
    Exit Sub
handler:
   Close #1
   Me.MousePointer = 0
   MsgBox Er.Description, vbExclamation
End Sub





'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Facturas_Clientes
'DESCRIPCIÓN            : Consulta las factutas del cliente seleccionado
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Facturas_Clientes()
Dim Cliente As String                                       'Obtiene el nombre del cliente
Dim Rs_Consulta_Clientes As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_IVA As Double                                     'Almacena el total del IVA
Dim Total As Double                                         'Almacena el total de la Sin IVA
Dim Total_Facturas As Double

On Error GoTo handler

    Grid_Reporte.Clear
    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Cliente_Factura.ListIndex > -1 Then
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Cliente_Factura.ItemData(Cmb_Cliente_Factura.ListIndex), "00000") & "'"
        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Clientes.EOF Then
            With Rs_Consulta_Clientes
                If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad")) Then
                    Ciudad = Rs_Consulta_Clientes.rdoColumns("Ciudad")
                Else
                    Ciudad = ""
                End If
                Open App.Path & "\REP\Reportes.txt" For Output As #1
                Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
                Print #1,
                Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE FACTURAS CLIENTES", Len("***************************************************************************************************************"))
                 Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Factura.text
                Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Factura.text
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End With
            Grid_Reporte.AddItem ""
            'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Serie='A' ORDER BY Fecha"
            Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
                With Rs_Consulta_Cat_Facturas_Clientes
                    Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "  No. Factura      Fecha           Subtotal            IVA              Total      Estatus      UUID                         "
                    Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & _
                                         Chr(9) & "Subtotal" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus" & Chr(9) & "UUID"
                    Grid_Reporte.AddItem ""
                    While Not .EOF
                        'Consulta el estatus
                        If Trim(.rdoColumns("Cancelada")) = "S" Then
                            Estatus = "Cancelada"
                        Else
                            If Trim(.rdoColumns("Pagada")) = "S" Then
                                Estatus = "Pagada"
                            Else
                                Estatus = "Pendiente"
                            End If
                        End If
                        If Trim(.rdoColumns("Cancelada")) = "N" Then
                            Print #1, Spc(2); .rdoColumns("No_Factura"); _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Estatus, 10); Spc(4); .rdoColumns("Timbre_UUID")
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & Estatus & Chr(9) & .rdoColumns("Timbre_UUID") & Chr(13)
                            Total_IVA = Total_IVA + .rdoColumns("IVA")
                            Total = Total + .rdoColumns("Subtotal")
                            Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Else
                            Print #1, Spc(2); .rdoColumns("No_Factura"); _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Estatus, 10); Spc(4); .rdoColumns("Timbre_UUID")
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & Estatus & Chr(9) & .rdoColumns("Timbre_UUID")
                        End If
                        .MoveNext
                    Wend
                    Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Spc(10); "T O T A L E S "; Spc(9); Conectar_Ayudante.Alinea_Derecha(Format(Total, "#,###,##0.00"), 14); Conectar_Ayudante.Alinea_Derecha(Format(Total_IVA, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
                End With
                Rs_Consulta_Cat_Facturas_Clientes.Close
                Close #1
                Me.Edm_Reporte.Font.Size = 8
                Me.Edm_Reporte.Font = "Courier New"
                Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
                Me.Edm_Reporte.Visible = True
                Ancho_Frame = Me.Width
                Largo_Frame = Me.Height
                Me.Width = 17235
                Me.Height = 8175
                Me.Left = 0
                Me.Top = 0
            Else
                Close #1
                MsgBox "No hay facturas del cliente seleccionado", vbInformation
            End If
        End If
    Else
        MsgBox "Seleccione un cliente", vbExclamation
    End If
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Remisiones_Clientes
'DESCRIPCIÓN            : Consulta las remisiones del cliente seleccionado
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Remisiones_Clientes()
Dim Cliente As String                                       'Obtiene el nombre del cliente
Dim Rs_Consulta_Clientes As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_IVA As Double                                     'Almacena el total del IVA
Dim Total_Tasa_0 As Double                                  'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                                 'Almacena el total de la tasa 15
Dim Total_Facturas As Double                                'Almacena el total de la factura



On Error GoTo handler

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Cliente_Factura.ListIndex > -1 Then
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Cliente_Factura.ItemData(Cmb_Cliente_Factura.ListIndex), "00000") & "'"
        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Clientes.EOF Then
            With Rs_Consulta_Clientes
                If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad")) Then
                    Ciudad = Rs_Consulta_Clientes.rdoColumns("Ciudad")
                Else
                    Ciudad = ""
                End If
                Open App.Path & "\REP\Reportes.txt" For Output As #1
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
                Print #1,
                Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE REMISIONES CLIENTES", Len("***************************************************************************************************************"))
                Print #1,
                Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Factura.text
                Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Factura.text
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End With
            Grid_Reporte.AddItem ""
            'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
                With Rs_Consulta_Cat_Facturas_Clientes
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " No. Remision      Fecha            Subtotal             IVA           Total      Estatus                                "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Fecha" & Chr(9) & _
                                         Chr(9) & "Subtotal" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                    Grid_Reporte.AddItem ""
                    While Not .EOF
                        'Consulta el estatus
                        If Trim(.rdoColumns("Cancelada")) = "S" Then
                            Estatus = "Cancelada"
                        Else
                            If Trim(.rdoColumns("Pagada")) = "S" Then
                                Estatus = "Pagada"
                            Else
                                Estatus = "Pendiente"
                            End If
                        End If
                        If Trim(.rdoColumns("Cancelada")) = "N" Then
                            Print #1, Spc(2); .rdoColumns("No_Remision"); _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & Estatus & Chr(13)
                                Total_IVA = Total_IVA + .rdoColumns("IVA")
                                Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Subtotal")
                                Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Else
                            Print #1, Spc(2); .rdoColumns("No_Remision"); _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha("0.00", 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & Estatus
                        End If
                        .MoveNext
                    Wend
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Spc(10); "T O T A L E S "; Spc(9); Spc(4); Spc(4); Spc(4); Spc(4); Spc(15); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
                End With
            End If
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
            Exit Sub
        End If
        Rs_Consulta_Cat_Facturas_Clientes.Close
    Else
        MsgBox "Seleccione un cliente", vbExclamation
    End If
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Facturas_Pagadas
'DESCRIPCIÓN            : Consulta las factutas que fueron pagadas
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Facturas_Pagadas()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro

On Error GoTo handler

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Fecha >= " & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Pagada='S' AND Serie='A'"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE FACTURAS PAGADAS", Len("***************************************************************************************************************"))
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(2); "No. Factura"; Spc(12 - Len("No. Factura")); "Cliente"; _
                      Spc(42 - Len("Cliente")); "Fecha"; Spc(15 - Len("Fecha")); "Subtotal"; _
                      Spc(11 - Len("Subtotal")); "IVA"; _
                      Spc(10 - Len("IVA")); "Total"; Spc(6 - Len("Total"));
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Subtotal" & _
                                 Chr(9) & "IVA" & Chr(9) & "Total"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 40); _
                Spc(42 - Len(Mid(Cliente, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(12 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 9); _
                Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 9))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9); _
                Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9); _
                Spc(10 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9)))
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                Format(.rdoColumns("Total"), "##,##0.00") & Chr(13)
                .MoveNext
            Wend
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
            Exit Sub
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
        DTP_Fecha_Inicio_Historico_Factura.Value = Now
        DTP_Fecha_Final_Historico_Factura.Value = Now
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Notas_Cargo
'DESCRIPCIÓN            : Consulta las notas de cargo
'PARÁMETROS             :
'CREO                   : Abel Oros
'FECHA_CREO             : 14-Oct-2017
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Notas_Cargo()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro

On Error GoTo handler

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Fecha >= " & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Serie='NCA'"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE NOTAS DE CARGO", Len("***************************************************************************************************************"))
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(2); "No. N Cargo"; Spc(12 - Len("No. Factura")); "Cliente"; _
                      Spc(42 - Len("Cliente")); "Fecha"; Spc(15 - Len("Fecha")); "Subtotal"; _
                      Spc(11 - Len("Subtotal")); "IVA"; _
                      Spc(10 - Len("IVA")); "Total"; Spc(6 - Len("Total"));
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. N Cargo" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Subtotal" & _
                                 Chr(9) & "IVA" & Chr(9) & "Total"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 40); _
                Spc(42 - Len(Mid(Cliente, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(12 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 9); _
                Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 9))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9); _
                Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9))); _
                Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9); _
                Spc(10 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9)))
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                Format(.rdoColumns("Total"), "##,##0.00") & Chr(13)
                .MoveNext
            Wend
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
            Exit Sub
        End With
    Else
        MsgBox "No hay notas de cargo en las fechas que proporcionó", vbInformation
        DTP_Fecha_Inicio_Historico_Factura.Value = Now
        DTP_Fecha_Final_Historico_Factura.Value = Now
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Facturas_Canceladas
'DESCRIPCIÓN            : Consulta las factutas que fueron canceladas
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Facturas_Canceladas()
Dim Cliente As String                                 'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset          'Consulta el cliente al cual pertenece la factura emitida
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset 'Consulta las facturas que fueron emitidas en las fechas dadas por el usuario

    
On Error GoTo handler
    
    Grid_Reporte.Cols = 12
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Cancelada='S' AND Serie='A'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE FACTURAS CANCELADAS", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(2); "No. Factura"; Spc(12 - Len("No. Factura")); "Cliente"; _
                      Spc(42 - Len("Cliente")); "Fecha"; Spc(21 - Len("Fecha")); "Subtotal"; Spc(14 - Len("Subtotal")); "IVA"; _
                      Spc(12 - Len("IVA")); "Total"; Spc(12 - Len("Total"));
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Subtotal" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Fecha Cancelo"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    Else
                        Cliente = "SIN NOMBRE"
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                Else
                    Cliente = "SIN NOMBRE"
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 40); _
                    Spc(42 - Len(Mid(Cliente, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(20 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
                    Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Subtotal"), "##,##0.00"), 11))); _
                    Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 11))); _
                    Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 11)))
                Print #1, Spc(5); "FECHA CANCELO: "; Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy HH:mm:ss"); Spc(5); "CANCELO: "; .rdoColumns("Usuario_Cancelo")
                Print #1,
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Subtotal"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy HH:mm:ss") & Chr(9) & _
                    .rdoColumns("Usuario_Cancelo")
                .MoveNext
            Wend
                Close #1
                Me.Edm_Reporte.Font.Size = 8
                Me.Edm_Reporte.Font = "Courier New"
                Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
                Me.Edm_Reporte.Visible = True
                Ancho_Frame = Me.Width
                Largo_Frame = Me.Height
                Me.Width = 13250
                Me.Height = 8200
                Me.Left = 0
                Me.Top = 0
                Exit Sub
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
        DTP_Fecha_Inicio_Historico_Factura.Value = Now
        DTP_Fecha_Final_Historico_Factura.Value = Now
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub
Private Sub Btn_Reporte_Facturas_Vencidas_Click()
    Call Reporte_Facturas_Vencidas
End Sub

Private Sub Btn_Reporte_Entradas_Salidas_Click()
    Me.MousePointer = 11
        If Fra_Reporte_Entradas_Salidas.Caption = "Reporte de Entradas de Productos" Then
            Call Reporte_Entradas_Producto
        Else
            Call Reporte_Salidas_Producto
        End If
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_General_Facturas_Click()
    Me.MousePointer = 11
        Select Case Fra_Reporte_General_Facturas.Caption
         Case "Reporte General de Facturas"
            Call Reporte_General_Facturas
         Case "Reporte General de Remisiones"
            Call Reporte_General_Remisiones
        Case "Reporte de notas de cargo"
                Reporte_Notas_Cargo
        End Select
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Historico_Facturas_Click()
    Me.MousePointer = 11
        Select Case Fra_Historico_Facturas.Caption
            Case "Facturas Pagadas"
                Reporte_Facturas_Pagadas
            Case "Facturas Canceladas"
                Reporte_Facturas_Canceladas
            Case "Facturas Cliente"
                Reporte_Facturas_Clientes
            Case "Remisiones Cliente"
                Reporte_Remisiones_Clientes
            Case "Reporte de notas de cargo"
                Reporte_Notas_Cargo
        End Select
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Historico_Ventas_Click()
    Select Case Fra_Reporte_Ventas.Caption
        Case "Ventas por Cliente"
            ''Reporte_Ventas_Clientes
        Case "Distribución de Producto"
            ''Reporte_Ventas_Productos
        Case "Consulta de Materias Primas"
            ''Reporte_Materia_Prima
        Case "Consulta de Productos Terminados"
            ''Reporte_Producto_Terminado
    End Select
End Sub




Private Sub Btn_Reporte_Notas_Credito_Click()
    If Opt_Tipo_Reporte_Nota_Credito(0).Value = True Then
        Reporte_Notas_Credito_Listado
    Else
        Reporte_Notas_Credito_Agrupado
    End If
End Sub

Private Sub Btn_Reporte_Pagos_Cheques_Click()
    Me.MousePointer = 11
    Select Case Fra_Reporte_Pagos_Cheques.Caption
        Case "Reporte pagos con cheques"
            Reporte_Pagos_Cheques
    End Select
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Pedido_Factura_Click()
    Me.MousePointer = 11
        Reporte_Salidas_Estatus
    Me.MousePointer = 0
End Sub
Private Sub Btn_Reporte_Pedidos_Cliente_Click()
    Me.MousePointer = 11
        Select Case Fra_Reporte_Pedidos_Clientes.Caption
            Case "Facturas por Producto"
                ''Reporte_Facturas_Productos
            Case "Reporte de Pedidos"
                Reporte_Pedidos
        End Select
    Me.MousePointer = 0
End Sub
Private Sub Reporte_Facturas_Por_Proveedor()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset        'Manejo de registro para las facturas
Dim Rs_Consulta_Proveedores As rdoResultset                 'Manejo de registro de la tabla Clientes
Dim Total_Facturas As Double
Dim Estatus As String


On Error GoTo handler

    If Cmb_Facturas_Proveedores.ListIndex > -1 Then
        Grid_Reporte.Rows = 0
        Open App.Path & "\REP\Reportes.txt" For Output As #1
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
        Print #1,
        Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE FACTURAS PROVEEDORES", Len("***************************************************************************************************************"))
        Print #1,
        'Consulta el Proveedor seleccionado
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Facturas_Proveedores.ItemData(Cmb_Facturas_Proveedores.ListIndex), "00000") & "'"
        Set Rs_Consulta_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Proveedores.EOF Then
            Print #1, Spc(30); "Proveedor : "; Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Grid_Reporte.AddItem "Proveedor" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Print #1, Spc(30); "R.F.C.    : "; Rs_Consulta_Proveedores.rdoColumns("RFC")
            Grid_Reporte.AddItem "R.F.C." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("RFC")
            Print #1, Spc(30); "Direccion : "; Mid(Rs_Consulta_Proveedores.rdoColumns("Direccion"), 1, 70)
            Grid_Reporte.AddItem "Direccion" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Direccion")
            Print #1, Spc(30); "Colonia   : "; Rs_Consulta_Proveedores.rdoColumns("Colonia")
            Grid_Reporte.AddItem "Colonia" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Colonia") & Chr(9) & "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Codigo_Postal")
            Print #1, Spc(30); "Ciudad    : "; Rs_Consulta_Proveedores.rdoColumns("Ciudad")
            Grid_Reporte.AddItem "Ciudad" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Ciudad")
            Print #1, Spc(30); "C.P.      : "; Rs_Consulta_Proveedores.rdoColumns("Codigo_Postal")
            Grid_Reporte.AddItem "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Codigo_Postal")
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Factura              Fecha         Tipo                     Total      Estatus                                     "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Tipo" & Chr(9) & "Total" & Chr(9) & "Estatus"
            'Consulta las facturas registradas del proveedor
            Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "MM/dd/yyyy") & Par_Fecha
            Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Facturas_Proveedores.EOF
                With Rs_Consulta_Facturas_Proveedores
                    If .rdoColumns("Cancelada") = "S" Then
                        Estatus = "Cancelada"
                    Else
                        If .rdoColumns("Pagada") = "S" Then
                            Estatus = "Pagada"
                        Else
                            Estatus = "Pendiente"
                        End If
                    End If
                    If .rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                End With
                Rs_Consulta_Facturas_Proveedores.MoveNext
            Wend
            Rs_Consulta_Facturas_Proveedores.Close
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(30); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
        End If
    Else
        MsgBox "Seleccione un proveedor", vbInformation
    End If
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Pedidos_Documento_Click()
    Me.MousePointer = 11
        Call Reporte_Pedidos_Facturas
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Producto_Por_Documento_Click()
    Me.MousePointer = 11
        Call Reporte_Remisiones_Productos
    Me.MousePointer = 0
End Sub

Private Sub Btn_Reporte_Productos_Facturados_Cliente_Click()
Me.MousePointer = 11
    Select Case Fra_Facturas_Proveedor.Caption
        Case "Facturas por Proveedor"
            Reporte_Facturas_Por_Proveedor
        Case "Notas de Credito por Proveedor"
            ''Reporte_Notas_Credito
        Case "Reporte de Cambio de Ubicaciones"
            ''Reporte_Cambio_Ubicaciones
    End Select
Me.MousePointer = 0
End Sub
Private Sub Btn_Reporte_Saldos_Click()
    Me.MousePointer = 11
        Select Case Fra_Saldos.Caption
            Case "Reporte de Saldo por Cliente"
                Reporte_Saldos_Cliente
            Case "Reporte de cuentas por pagar"
                Reporte_Saldos_Proveedores
        End Select
    Me.MousePointer = 0
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCION   : Reporte_Saldos_Proveedores
'DESCRIPCION            : Muestra el reporte saldos por proveedor
'CREO                   : Julio Cruz
'FECHA_CREO             : 01-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Reporte_Saldos_Proveedores()
Dim Proveedor As String                                     'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Cont_Proveedores As Double
Dim Rs_Consulta_Facturas_Poveedor As rdoResultset
Dim Total_Importe As Double
Dim Total_IVA As Double
Dim Total_Total As Double
Dim Total_Saldo As Double
Dim Gran_Total_Importe As Double
Dim Gran_Total_IVA As Double
Dim Gran_Total_Total As Double
Dim Gran_Total_Saldo As Double


On Error GoTo handler
    
    Cont_Proveedores = 0
    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    Total_Importe = 0
    Total_IVA = 0
    Total_Total = 0
    Total_Saldo = 0
    Gran_Total_Importe = 0
    Gran_Total_IVA = 0
    Gran_Total_Total = 0
    Gran_Total_Saldo = 0
    ''Consulta las facturas que fueron emitidas al proveedor seleccionado por el usuario
    Mi_SQL = ""
    Mi_SQL = " SELECT DISTINCT Cat_Proveedores.Proveedor_ID,Cat_Proveedores.*  FROM Cat_Proveedores,Adm_Proveedores_Facturas "
    Mi_SQL = Mi_SQL & " WHERE Adm_Proveedores_Facturas.Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Facturas.Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Cat_Proveedores.Proveedor_ID= Adm_Proveedores_Facturas.Proveedor_ID"
    If Cmb_Cliente_Saldo.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Proveedores.Proveedor_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Proveedores.Proveedor_ID "
    Set Rs_Consulta_Adm_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Facturas_Proveedores.EOF Then
        While Not Rs_Consulta_Adm_Facturas_Proveedores.EOF
                If Cont_Proveedores = 0 Then
                    Open App.Path & "\REP\Reportes.txt" For Output As #1
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
                    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
                    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
                    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
                    Print #1,
                    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE CUENTAS POR PAGAR", Len("***************************************************************************************************************"))
                    Print #1, Conectar_Ayudante.Centrar_Texto("FECHA GENERACION:" & Format(Now, "MM/dd/yyyy"), Len("***************************************************************************************************************"))
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                End If
            'SE CONSULTAN LAS FACTURAS DEL PROVEEDOR
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Rs_Consulta_Adm_Facturas_Proveedores!Proveedor_ID & "'"
            If Chk_Imprimir_Saldos_ceros.Value = 0 Then Mi_SQL = Mi_SQL & " AND Saldo > 0 "
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Facturas_Poveedor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Facturas_Poveedor
                Print #1,
                If Not Rs_Consulta_Facturas_Poveedor.EOF Then
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, Spc(2); "Proveedor:   "; Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns!Nombre
                    ''Print #1, Spc(2); "Direccion:   "; Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Direccion")
                    ''Print #1, Spc(2); "Colonia:     "; Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Colonia"); Spc(2); "CP: "; Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Codigo_Postal")
                    ''Print #1, Spc(2); "Ciudad:      "; Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Ciudad")
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Proveedor:" & Chr(9) & Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns!Nombre
                    ''Grid_Reporte.AddItem "Dirección:" & Chr(9) & Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Direccion")
                    ''Grid_Reporte.AddItem "Colonia:" & Chr(9) & Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Codigo_Postal")
                    ''Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Ciudad")
                    Grid_Reporte.AddItem ""
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "No.Factura           Tipo             Fecha             Importe           IVA           Total          Saldo             "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Tipo" & Chr(9) & "Fecha" & Chr(9) & "Importe" & _
                    Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Saldo"
                End If
                While Not .EOF
                    If .rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(2); Mid(Trim(.rdoColumns("No_Factura")), 1, 20); _
                        Spc(17 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 20))); _
                        Spc(2); Mid(Trim(.rdoColumns("Tipo")), 1, 15); _
                        Spc(17 - Len(Mid(Trim(.rdoColumns("Tipo")), 1, 15))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 14); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "#,###,##0.00"), 14); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 14)
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") _
                        & Chr(9) & .rdoColumns("Tipo") _
                        & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & Format(.rdoColumns("Importe"), "#0.00") _
                        & Chr(9) & Format(.rdoColumns("IVA"), "#0.00") _
                        & Chr(9) & Format(.rdoColumns("Total"), "#0.00") _
                        & Chr(9) & Format(.rdoColumns("Saldo"), "#0.00")
                        'SE INCREMENTA EL TOTAL DEL CLIENTE
                        Total_Importe = Total_Importe + Val(.rdoColumns("Importe"))
                        Total_IVA = Total_IVA + Val(.rdoColumns("IVA"))
                        Total_Total = Total_Total + Val(.rdoColumns("Total"))
                        Total_Saldo = Total_Saldo + Val(.rdoColumns("Saldo"))
                        'SE INCREMENTA EL GRAN TOTAL
                        Gran_Total_Importe = Gran_Total_Importe + Val(.rdoColumns("Importe"))
                        Gran_Total_IVA = Gran_Total_IVA + Val(.rdoColumns("IVA"))
                        Gran_Total_Total = Gran_Total_Total + Val(.rdoColumns("Total"))
                        Gran_Total_Saldo = Gran_Total_Saldo + Val(.rdoColumns("Saldo"))
                    End If
                    .MoveNext
                Wend
                Rs_Consulta_Facturas_Poveedor.Close
                'SE CONSULTAN LAS FACTURAS DEL PROVEEDOR
                Mi_SQL = "SELECT * "
                Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas"
                Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
                Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & Par_Fecha
                Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Rs_Consulta_Adm_Facturas_Proveedores!Proveedor_ID & "'"
                If Chk_Imprimir_Saldos_ceros.Value = 0 Then Mi_SQL = Mi_SQL & " AND Saldo > 0 "
                Mi_SQL = Mi_SQL & " ORDER BY Fecha"
                Set Rs_Consulta_Facturas_Poveedor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Facturas_Poveedor.EOF Then
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, "TOTALES PROVEEDOR:                             "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Importe, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_IVA, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Total, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Saldo, "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "TOTALES PROVEEDOR:" & Chr(9) & Chr(9) & Chr(9) & Total_Importe & Chr(9) & Total_IVA & Chr(9) & Total_Total & Chr(9) & Total_Saldo
                End If
            Total_Importe = 0
            Total_IVA = 0
            Total_Total = 0
            Total_Saldo = 0
            End With
            Rs_Consulta_Facturas_Poveedor.Close
            Cont_Proveedores = Cont_Proveedores + 1
            Rs_Consulta_Adm_Facturas_Proveedores.MoveNext
        Wend
        Print #1, ""
        Print #1, "_________________________________________________________________________________________________________________________"
        Print #1, "GRAN TOTAL:                                    "; Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_Importe, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_IVA, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_Total, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_Saldo, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "GRAN TOTAL:" & Chr(9) & Chr(9) & Chr(9) & Gran_Total_Importe & Chr(9) & Gran_Total_IVA & Chr(9) & Gran_Total_Total & Chr(9) & Gran_Total_Saldo
        Close #1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
        Exit Sub
    Else
        MsgBox "No hay facturas que mostar", vbInformation
    End If
    Rs_Consulta_Adm_Facturas_Proveedores.Close
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
End Sub
Private Sub Btn_Salir_Click()
Dim Mi_Frame As Object    'Indica que objeto se esta consultando
Dim Frame_Visible As Boolean
    For Each Mi_Frame In Me.Controls
        If TypeOf Mi_Frame Is Frame Then
            If Mi_Frame.Visible = True Then
                Frame_Visible = True
            End If
        End If
    Next
    If Frame_Visible = True Then
        Edm_Reporte.Visible = False
        Pic_Reportes.Visible = True
        If Me.WindowState <> 2 Then
            Me.Width = Ancho_Frame
            Me.Height = Largo_Frame
            Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
            Me.Top = 100
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Cmb_Categorias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Categoria_ID,Nombre", "Cat_Categorias", Cmb_Categorias, 1, " Estatus='ACTIVO' AND Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Categorias.hwnd, &H14F, True, 0)
    End If
End Sub


''Private Sub Chk_Fechas_Saldos_Click()
''    If Chk_Fechas_Saldos.Value = 1 Then
''        Dtp_Fecha_Saldo_Inicio.Enabled = True
''        Dtp_Fecha_Saldo_Inicio.Value = Now
''        Dtp_Fecha_Saldo_Fin.Enabled = True
''        Dtp_Fecha_Saldo_Fin.Value = Now
''    Else
''        Dtp_Fecha_Saldo_Inicio.Enabled = False
''        Dtp_Fecha_Saldo_Fin.Enabled = False
''    End If
''End Sub
''Private Sub Chk_General_Click()
''    If Chk_General.Value = 0 Then
''        Dtp_Fecha_Inicio_Facturas_Directas.Enabled = True
''        Dtp_Fecha_Fin_Facturas_Directas.Enabled = True
''    Else
''        Dtp_Fecha_Inicio_Facturas_Directas.Enabled = False
''        Dtp_Fecha_Fin_Facturas_Directas.Enabled = False
''    End If
''End Sub
Public Sub Cmb_Cliente_Anticipo_KeyPress(KeyAscii As Integer)
Dim Mi_SQL As String
Dim Rs_Adm_Cuenta_Ajustes As rdoResultset

    If KeyAscii = 13 Then
        Select Case Fra_Anticipos_Clientes.Caption
            Case "Anticipos Clientes"
                Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Anticipo, 1, "Nombre")
            Case "Anticipos Proveedores"
                Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Cliente_Anticipo, 1, "Nombre")
            Case "Reporte de Pago de Saldos Pequeños"
                Mi_SQL = "SELECT No_Cuenta,Empresa,Cliente FROM Adm_Cuentas_Ajustes"
                Mi_SQL = Mi_SQL & " WHERE Cliente LIKE '%" & Cmb_Cliente_Anticipo.text & "%'"
                Set Rs_Adm_Cuenta_Ajustes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Cliente_Anticipo.Clear
                While Not Rs_Adm_Cuenta_Ajustes.EOF
                    If Rs_Adm_Cuenta_Ajustes.rdoColumns("Empresa") = 1 Then
                        Cmb_Cliente_Anticipo.AddItem "NHE - " & Rs_Adm_Cuenta_Ajustes.rdoColumns("Cliente")
                    Else
                        Cmb_Cliente_Anticipo.AddItem "PNC - " & Rs_Adm_Cuenta_Ajustes.rdoColumns("Cliente")
                    End If
                    Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.NewIndex) = Rs_Adm_Cuenta_Ajustes.rdoColumns("No_Cuenta")
                    Rs_Adm_Cuenta_Ajustes.MoveNext
                Wend
                Rs_Adm_Cuenta_Ajustes.Close
        End Select
        Cmb_Cliente_Anticipo.text = ""
    End If
End Sub
Private Sub Cmb_Cliente_Factura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_Id,Nombre", "Cat_Clientes", Cmb_Cliente_Factura, 1, "Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Cliente_Factura.hwnd, &H14F, True, 0)
    End If
End Sub
Private Sub Cmb_Cliente_Saldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Catalogo
            Case "Reporte de Saldo por Cliente"
                Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Saldo, 1, "Nombre")
            Case "Reporte de cuentas por pagar"
                Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Cliente_Saldo, 1, "Nombre")
        End Select
        Cmb_Cliente_Saldo.text = ""
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Cliente_Saldo.hwnd, &H14F, True, 0)
    End If
End Sub
Private Sub Cmb_Facturas_Proveedores_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Lbl_Proveedor_Reporte.Caption = "Proveedor" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Facturas_Proveedores, 1, "Nombre")
        End If
        Cmb_Facturas_Proveedores.text = ""
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Facturas_Proveedores.hwnd, &H14F, True, 0)
    End If
End Sub

Private Sub Cmb_Producto_Materia_Prima_Change()

End Sub

''Private Sub Cmb_Producto_Inventario_General_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then
''''        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Producto_Inventario_General, 1, "Nombre")
''''    End If
''End Sub


Private Sub Cmb_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Producto, 1, "Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Producto.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Producto_Reporte_Producto_Por_Documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Producto_Reporte_Producto_Por_Documento, 1, "Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Producto_Reporte_Producto_Por_Documento.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Producto_Venta_KeyPress(KeyAscii As Integer)
Dim Rs_Materias_Primas As rdoResultset
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Producto_Venta, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Proveedor_Entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor_Entrada, 2, "Nombre")
    End If
End Sub


Private Sub Cmb_Reporte_Pedido_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
''        If Lbl_Reporte_Pedido_Producto.Caption = "Clientes" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Reporte_Pedido_Producto, 1, "Nombre")
''        Else
''            Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Reporte_Pedido_Producto, 1, "Nombre")
''        End If
        Cmb_Reporte_Pedido_Producto.text = ""
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Reporte_Pedido_Producto.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Vendedor_Saldo_Click()
    Cmb_Cliente_Saldo.Clear
    Call Cmb_Cliente_Saldo_KeyPress(13)
End Sub

Private Sub Cmb_Reporte_ventas_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Reporte_ventas_Cliente, 1, "Nombre")
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Reporte_ventas_Cliente.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Tipo_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_ID,Nombre", "Cat_Productos_Tipo", Cmb_Tipo_Producto, 1, "Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Tipo_Producto.hwnd, &H14F, True, 0)
    End If
End Sub

Private Sub Cmb_Tipo_Producto_Producto_Por_Documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_ID,Nombre", "Cat_Productos_Tipo", Cmb_Tipo_Producto_Producto_Por_Documento, 1, " Estatus='ACTIVO' AND Nombre")
    Else
        'DESPLIEGA EL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Tipo_Producto_Producto_Por_Documento.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Form_Load()
    If Catalogo = "Reporte de Ventas" Then
        Me.Width = 5025
        Me.Height = 4300
        Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
        Me.Top = 100
    Else
        Me.Width = 5025
        Me.Height = 3570
        Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
        Me.Top = 100
    End If
    Call Cmb_Reporte_Pedido_Producto_KeyPress(13)
    Call Cmb_Reporte_Pedido_Producto_KeyPress(13)
    Call Cmb_Reporte_ventas_Cliente_KeyPress(13)
    Call Cmb_Producto_Reporte_Producto_Por_Documento_KeyPress(13)
    Call Cmb_Tipo_Producto_Producto_Por_Documento_KeyPress(13)
    Call Cmb_Categorias_KeyPress(13)
    Call Cmb_Tipo_Producto_KeyPress(13)
    Call Cmb_Producto_KeyPress(13)
    Call Cmb_Cliente_Factura_KeyPress(13)
    Call Cmb_Cliente_Factura_KeyPress(13)
    Call Cmb_Facturas_Proveedores_KeyPress(13)
    
    Call Cmb_Cliente_Saldo_KeyPress(13)
    If Catalogo = "Facturas_Vencidas" Then Reporte_Facturas_Vencidas
    If Catalogo = "Facturas_Por_Vencer" Then Reporte_Facturas_Vencer
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCION   : Reporte_Salidas_Estatus
'DESCRIPCION            : Muestra el reporte salidas segun sus estatus
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Reporte_Salidas_Estatus()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset        'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Ciudad As String                            'Almacena el nombre de la ciudad
Dim Total As Double                             'Almacena el total del pedido
Dim Total_Producto As Double                    'Almacena el total de productos
Dim Productos_Pedido As Double                  'Almacena el total del producto por pedido

On Error GoTo handler

    'Coloca el encabezado de los reportes
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE SALIDAS PENDIENTES POR DOCUMENTAR", Len("***************************************************************************************************************"))
    Grid_Reporte.AddItem ""
    'Consulta
    Mi_SQL = " SELECT Alm_Salidas_Almacen.*,Alm_Salidas_Almacen_Detalles.*, Cat_Clientes.Nombre as Nombre_Cliente    FROM Alm_Salidas_Almacen,Alm_Salidas_Almacen_Detalles,Cat_Clientes"
    Mi_SQL = Mi_SQL & " Where Alm_Salidas_Almacen.No_Salida = Alm_Salidas_Almacen_Detalles.No_Salida"
    Mi_SQL = Mi_SQL & " AND   Alm_Salidas_Almacen_Detalles.Facturado='NO'"
    Mi_SQL = Mi_SQL & " AND Fecha_Salida>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha_Salida<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedido_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID=Alm_Salidas_Almacen.Cliente_ID"
    Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Grid_Reporte.AddItem ""
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, " No. Salida                  Cliente                                                                 Fecha de la salida "
    Print #1, "________________________________________________________________________________________________________________________"
    With Rs_Consulta_Pedidos
        While Not .EOF
            'Imprime los pedidos
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No_Salida" & Chr(9) & "Nombre_Cliente" & Chr(9) & "Fecha_Salida"
            Print #1, Spc(2); .rdoColumns("No_Salida"); Spc(27 - Len(.rdoColumns("No_Salida"))); .rdoColumns("Nombre_Cliente"); ; Spc(73 - Len(.rdoColumns("Nombre_Cliente"))); Format(.rdoColumns("Fecha_Salida"), "MM/dd/yyyy")
            Grid_Reporte.AddItem .rdoColumns("No_Salida") & Chr(9) & .rdoColumns("Nombre_Cliente") & Chr(9) & Format(.rdoColumns("Fecha_Salida"), "MM/dd/yyyy")
            .MoveNext
        Wend
    End With
    Rs_Consulta_Pedidos.Close
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCION   : Reporte_Pagos_Cheques
'DESCRIPCION            : Muestra el reporte salidas segun sus estatus
'CREO                   : Julio Cruz
'FECHA_CREO             : 02-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Reporte_Pagos_Cheques()
Dim Rs_Consulta_Movimientos As rdoResultset         'Manejo de registro
Dim Total As Double                                 'Almacena el total del pedido
Dim Estatus As String
Dim Total_Cantidad As Double

On Error GoTo handler
    'Coloca el encabezado de los reportes
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Grid_Reporte.Rows = 0
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE PAGOS CHEQUES", Len("***************************************************************************************************************"))
    Grid_Reporte.AddItem ""
    'Consulta
    Mi_SQL = " SELECT * FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Forma_Pago = 'Cheque'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo = 'E'"
    Mi_SQL = Mi_SQL & " AND   Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pagos_Cheques.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND   Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pagos_Fechas.Value, "MM/dd/yyyy") & Par_Fecha
    Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Grid_Reporte.AddItem ""
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, " Fecha               No.Cheque            Banco                       Estatus                               Cantidad    "
    Print #1, "________________________________________________________________________________________________________________________"
    With Rs_Consulta_Movimientos
        While Not .EOF
            'Imprime los pedidos
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Fecha" & Chr(9) & "No.Cheque" & Chr(9) & "Banco" & Chr(9) & "Estatus" & Chr(9) & "Cantidad"
            If Trim(.rdoColumns("Estatus")) = "A" Then
                Estatus = "ACTIVO"
            Else
                Estatus = "CANCELADO"
            End If
            Total_Cantidad = Total_Cantidad + Val(.rdoColumns("Cantidad"))
            Print #1, Spc(2); Format(.rdoColumns("Fecha"), "MM/dd/yyyy"); Spc(20 - Len(.rdoColumns("Fecha"))); _
            .rdoColumns("Referencia"); Spc(19 - Len(.rdoColumns("Referencia"))); _
            .rdoColumns("Banco"); _
            Spc(29 - Len(.rdoColumns("Banco"))); Estatus; Spc(37 - Len(Estatus)) _
            ; Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#,###,###.00"), 10)
            
            Grid_Reporte.AddItem .rdoColumns("Fecha") & Chr(9) & .rdoColumns("Referencia") _
            & Chr(9) & .rdoColumns("Banco") & Chr(9) & Estatus & Chr(9) & Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#,###,###.00"), 15)
            .MoveNext
        Wend
    End With
    Rs_Consulta_Movimientos.Close
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1, Spc(99); "Total : "; Format(Total_Cantidad, "#,###,###.00")
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Total" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Format(Total_Cantidad, "#,###,###.00")
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Inventario_Materia_Prima
'DESCRIPCIÓN            : Muestra el inventario general
'CREO                   : Julio Cruz
'FECHA_CREO             : 08-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'PARÁMETROS             :
'*******************************************************************************
Private Sub Reporte_Inventario_General()
Dim Rs_Consulta_Tipo_Producto As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Producto As rdoResultset            'Manejo de registro
Dim Total As Double                                 'Almacena el total de la existencia
Dim Costo_Almacen As Double                         'Almacena el total del costo del almacen
Dim Nombre_Corto As String
Dim Existencia As Double


On Error GoTo handler
    
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("INVENTARIO GENERAL DE PRODUCTOS ", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
    Print #1, " "
    Print #1, "   Clave              Descripcion                                 Existencia          Costo        Total               "
    Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Existencia" & Chr(9) & "Unidad" & Chr(9) & "Ultimo Costo" & Chr(9) & "Total"
    Grid_Reporte.AddItem ""
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1,
    Total = 0
    Costo_Almacen = 0
    Mi_SQL = "SELECT * FROM Cat_Productos_Tipo"
    If Cmb_Tipo_Producto.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " WHERE Tipo_ID='" & Format(Cmb_Tipo_Producto.ItemData(Cmb_Tipo_Producto.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Tipo_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Hace el recorrido de los registros encontrados
    While Not Rs_Consulta_Tipo_Producto.EOF
        'Consulta los productos de la linea en Turno
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Productos "
        Mi_SQL = Mi_SQL & " WHERE Cat_Productos.Tipo_ID='" & Rs_Consulta_Tipo_Producto.rdoColumns("Tipo_ID") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Productos.Clave"
        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Imprime los productos encontrados
        If Not Rs_Consulta_Producto.EOF Then
            Print #1, Spc(40); "TIPO DE PRODUCTO :     "; Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Print #1,
            Grid_Reporte.AddItem "TIPO DE PRODUCTO : " & Chr(9) & Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Grid_Reporte.AddItem ""
            With Rs_Consulta_Producto
                While Not .EOF
                    If Not IsNull(.rdoColumns("Existencia")) Then
                        Existencia = .rdoColumns("Existencia")
                    Else
                        Existencia = 0
                    End If
                    Print #1, Mid(.rdoColumns("Clave"), 1, 20); _
                        Spc(22 - Len(Mid(.rdoColumns("Clave"), 1, 20))); Mid(.rdoColumns("Nombre"), 1, 40); _
                        Spc(43 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); Conectar_Ayudante.Alinea_Derecha(Format(Existencia, "#,###,###0.00"), 15); _
                        Spc(1); Nombre_Corto; _
                        Spc(4 - Len(Nombre_Corto)); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Costo"), "#,##0.00"), 8); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Costo")) * Val(Existencia), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem .rdoColumns("Clave") _
                        & Chr(9) & .rdoColumns("Nombre") _
                        & Chr(9) & Existencia _
                        & Chr(9) & Nombre_Corto _
                        & Chr(9) & .rdoColumns("Costo") _
                        & Chr(9) & Val(.rdoColumns("Costo")) * Existencia
                    Total = Val(Total) + Val(Existencia)
                    Costo_Almacen = Val(Costo_Almacen) + (Val(Existencia * .rdoColumns("Costo")))
                    Rs_Consulta_Producto.MoveNext
                Wend
            End With
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
        End If
        Rs_Consulta_Producto.Close
        Rs_Consulta_Tipo_Producto.MoveNext
    Wend
    Rs_Consulta_Tipo_Producto.Close
    Print #1,
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    'Imprime los totales de piezas y costo
    Print #1, "                                                                       TOTAL EXISTENCIA PIEZAS       "; Conectar_Ayudante.Alinea_Derecha(Format(Total, "#,###,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "TOTAL EXISTENCIA PIEZAS" & Chr(9) & Total
    Print #1, "                                                                       COSTO INVENTARIO PIEZAS       "; Conectar_Ayudante.Alinea_Derecha(Format(Costo_Almacen, "#,###,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "COSTO INVENTARIO PIEZAS" & Chr(9) & Costo_Almacen
    Grid_Reporte.AddItem ""
    Print #1,
    Print #1, "                                                                       COSTO TOTAL INVENTARIO        "; Conectar_Ayudante.Alinea_Derecha(Format(Val(Costo_Almacen), "##,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "COSTO TOTAL INVENTARIO" & Chr(9) & Costo_Almacen
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Entradas_Producto
'DESCRIPCIÓN            : Muestra el reporte de las entradas de Producto
'CREO                   : Julio Cruz
'FECHA_CREO             : 08-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'PARÁMETROS             :
'*******************************************************************************
Private Sub Reporte_Entradas_Producto()
Dim Mi_SQL As String
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles As rdoResultset   'Manejo de registro
Dim Proveedor As String                             'Almacena el nombre del proveedor
Dim Gran_total As Double                            'Almacena el total de las entadas
Dim Total_Entrada As Double                         'Almacena el total parcial de la entrada consultada en turno
Dim Rs_Consulta_Clasificacion_Proveedor As rdoResultset
Dim Tipo As String
Dim Fecha_Caducidad As Double


On Error GoTo handler

    'Consulta las entradas
    Mi_SQL = "SELECT Entrada_ID,Fecha_Recepcion_Factura,Tipo_Entrada,ISNULL(Alm_Entradas.Estatus,'ACTIVA') As Estatus,Alm_Entradas.Costo_Total,Alm_Entradas.Proveedor_ID,Cat_Proveedores.Clasificacion_ID,Cat_Proveedores.Nombre AS Proveedor"
    Mi_SQL = Mi_SQL & " FROM Alm_Entradas,Cat_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Alm_Entradas.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    If Cmb_Proveedor_Entrada.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Proveedor_ID='" & Format(Cmb_Proveedor_Entrada.ItemData(Cmb_Proveedor_Entrada.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Alm_Entradas.Fecha_Recepcion_Factura BETWEEN '" & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas.Fecha_Recepcion_Factura,Alm_Entradas.Entrada_ID"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE ENTRADAS DE PRODUCTO ", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, ""
    Gran_total = 0
    Grid_Reporte.Cols = 15
    While Not Rs_Consulta_Entradas.EOF
        If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Proveedor_ID")) Then
            Proveedor = Rs_Consulta_Entradas.rdoColumns("Proveedor")
        Else
            Proveedor = ""
        End If
        Mi_SQL = " SELECT * FROM Cat_Clasificacion_Proveedores WHERE Clasificacion_ID = '" & Rs_Consulta_Entradas!Clasificacion_ID & "'"
        Set Rs_Consulta_Clasificacion_Proveedor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Tipo = Rs_Consulta_Clasificacion_Proveedor!Nombre
        Print #1,
        Print #1, "__________________________________________________________________________________________________________________________"
        Print #1, "No. Entrada : "; Rs_Consulta_Entradas.rdoColumns("Entrada_ID"); "   Fecha : "; Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15))); "Tipo Entrada : "; Mid(Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada"), 1, 15)
        Print #1, "Proveedor   : "; Mid(Proveedor, 1, 33); Spc(35 - Len(Mid(Proveedor, 1, 33))); "Tipo    : "; Mid(UCase(Tipo), 1, 15); Spc(17 - Len(Mid(Tipo, 1, 15))); "Costo    : "; Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Entradas.rdoColumns("Costo_Total"), "#,###,##0.00"), 12)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Entrada_ID") & Chr(9) & "Fecha Entrada :" & Chr(9) & Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yyyy") & Chr(9) & "Estatus :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Estatus") & Chr(9) & "Tipo Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada") & Chr(9) & "Proveedor :" & Chr(9) & Proveedor & Chr(9) & "Tipo :" & Chr(9) & Tipo & Chr(9) & "Costo :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Costo_Total")
        'Consulta el detalle de las entradas
        Mi_SQL = "SELECT Alm_Entradas_Detalles.Entrada_ID,Alm_Entradas_Detalles.No_Lote,"
        Mi_SQL = Mi_SQL & " Alm_Entradas_Detalles.Fecha_Caducidad,Alm_Entradas_Detalles.Cantidad,"
        Mi_SQL = Mi_SQL & " Alm_Entradas_Detalles.Costo,Cat_Productos.Clave,Cat_Productos.Nombre"
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Detalles,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Detalles.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Detalles.Entrada_ID ='" & Rs_Consulta_Entradas.rdoColumns("Entrada_ID") & "'"
        Set Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        Print #1, "  Codigo         Descripcion                          No. Lote              Caducidad    Cantidad    Costo       Total    "
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "No. Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        With Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles
            While Not .EOF
                If Not IsNull(.rdoColumns("Fecha_Caducidad")) Then
                    Fecha_Caducidad = .rdoColumns("Fecha_Caducidad")
                Else
                    Fecha_Caducidad = 0
                End If
                If Not IsNull(.rdoColumns("No_Lote")) And Trim(.rdoColumns("No_Lote")) <> "" Then
                    Print #1, Mid(.rdoColumns("Clave"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Clave"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 35); _
                        Spc(37 - Len(Mid(.rdoColumns("Nombre"), 1, 35))); Mid(.rdoColumns("No_Lote"), 1, 20); _
                        Spc(22 - Len(.rdoColumns("No_Lote"))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
                        Spc(11 - Len(Fecha_Caducidad)); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "####0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Costo"), "#,##0.000"), 9); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Costo") * .rdoColumns("Cantidad")), "#,###,##0.00"), 12)
                Else
                    Print #1, Mid(.rdoColumns("Clave"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Clave"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 36); _
                        Spc(37 - Len(Mid(.rdoColumns("Nombre"), 1, 36))); ""; _
                        Spc(33); ""; _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "####0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Costo"), "#,##0.000"), 9); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Costo") * .rdoColumns("Cantidad")), "#,###,##0.00"), 12)
                End If
                Grid_Reporte.AddItem .rdoColumns("Clave") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("No_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo") _
                    & Chr(9) & .rdoColumns("Costo") * .rdoColumns("Cantidad")
                'Calcula el total de cada entrada
                Total_Entrada = Total_Entrada + Val(.rdoColumns("Costo") * .rdoColumns("Cantidad"))
                .MoveNext
            Wend
        End With
        'Calcula el gran total de las entradas
        Gran_total = Gran_total + Total_Entrada
        Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles.Close
        Rs_Consulta_Entradas.MoveNext
    Wend
    Print #1,
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1, Spc(61); "Gran Total : "; Spc(10); Conectar_Ayudante.Alinea_Derecha(Format(Gran_total, "#,###,##0.00"), 16)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_total
    Rs_Consulta_Entradas.Close
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Salidas_Producto
'DESCRIPCIÓN            : Muestra el reporte de las salidas
'CREO                   : Julio Cruz
'FECHA_CREO             : 08-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'PARÁMETROS             :
'*******************************************************************************
Private Sub Reporte_Salidas_Producto()
Dim Mi_SQL As String
Dim Rs_Consulta_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Alm_Salidas_Materia_Prima_Detalles As rdoResultset     'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset    'Manejo de registro
Dim Cliente As String                       'Almacena el nombre del cliente
Dim Total_Costos As Double                  'Almacena el total de los costos por salida
Dim Gran_Total_Costos As Double             'Almacena el gran total de los costos de las salidas


On Error GoTo handler

    'Consulta las salidas de acuerdo a las fechas proporcionadas
    Mi_SQL = "SELECT Alm_Salidas_Almacen.*, Cat_Clientes.Nombre as  Nombre_Solicito"
    Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen, Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Fecha_Salida BETWEEN '" & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Cliente_ID = Cat_Clientes.Cliente_ID "
    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Salida,No_Salida"
    Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE SALIDAS DE PRODUCTO ", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
    Print #1, ""
    Gran_Total_Costos = 0
    While Not Rs_Consulta_Salidas.EOF
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "No. Salida  : "; Rs_Consulta_Salidas.rdoColumns("No_Salida"); "   Fecha : "; Format(Rs_Consulta_Salidas.rdoColumns("Fecha_Salida"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Consulta_Salidas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Salidas.rdoColumns("Estatus"), 1, 15))); "Tipo Salida  : "; Mid(Rs_Consulta_Salidas.rdoColumns("Tipo_Salida"), 1, 15)
        Print #1, "Solicito    : "; Mid(Rs_Consulta_Salidas.rdoColumns("Nombre_Solicito"), 1, 50)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Salida : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("No_Salida") & Chr(9) & "Fecha : " & Chr(9) & Format(Rs_Consulta_Salidas.rdoColumns("Fecha_Salida"), "dd/MMM/yyyy") & Chr(9) & "Estatus : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Estatus") & Chr(9) & "Tipo Salida : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Tipo_Salida") & Chr(9) & "Solicito : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Nombre_Solicito")
        'Consulta la descripcion de las salidas
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen_Detalles,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen_Detalles.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND No_Salida ='" & Rs_Consulta_Salidas.rdoColumns("No_Salida") & "'"
        Set Rs_Alm_Salidas_Materia_Prima_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Costos = 0
        Print #1, "  Codigo          Descripcion                        Cantidad    Costo      Total                                          "
        Print #1, "________________________________________________________________________________________________________________________"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total Compra"
        With Rs_Alm_Salidas_Materia_Prima_Detalles
            While Not .EOF
                Print #1, Mid(.rdoColumns("Clave"), 1, 15); _
                    Spc(16 - Len(Mid(.rdoColumns("Clave"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 29); _
                    Spc(31 - Len(Mid(.rdoColumns("Nombre"), 1, 29))); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "####0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Costo"), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo")), "#,##0.00"), 10)
                Grid_Reporte.AddItem .rdoColumns("Clave") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo") _
                    & Chr(9) & Val(.rdoColumns("Cantidad") * .rdoColumns("Costo"))
                Total_Costos = Total_Costos + Val(.rdoColumns("Cantidad") * .rdoColumns("Costo"))
                .MoveNext
            Wend
        End With
        Print #1,
        'Calcula el gran total de todas las salidas
        Gran_Total_Costos = Gran_Total_Costos + Total_Costos
        Rs_Alm_Salidas_Materia_Prima_Detalles.Close
        Rs_Consulta_Salidas.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(44); "Gran Total   : "; Spc(11); Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_Costos, "#,###,##0.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total_Costos
    Rs_Consulta_Salidas.Close
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Kardex_Materia_Prima
'DESCRIPCIÓN            : Muestra el kardex de materia prima
'CREO                   : Julio Cruz
'FECHA_CREO             : 08-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'PARÁMETROS             :
'*******************************************************************************
Private Sub Kardex_Materia_Prima()
Dim Mi_SQL As String
Dim Rs_Consulta_Productos As rdoResultset           'Manejo de registro
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Salidas As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Unidad As rdoResultset
Dim Total_Entrada As Double                         'Almacena el costo total de las entradas
Dim Total_Salida As Double                          'Almacena el costo total de las salidas
Dim Total_Productos As Double                       'Almacena la cantidad total de piezas
Dim Costo As Double
Dim Fecha_Caducidad As String


On Error GoTo handler


If Cmb_Producto.ListIndex > -1 Then

        'Kardex de productos
        Open App.Path & "\REP\Reportes.txt" For Output As #1
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
        Print #1,
        Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE KARDEX DE PRODUCTO ", Len("***************************************************************************************************************"))
        Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, ""
        'Consulta los productos
        Mi_SQL = "SELECT * FROM Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(30); "Clave      : "; Rs_Consulta_Productos.rdoColumns("Clave")
            Grid_Reporte.AddItem "Clave :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Clave")
            Print #1, Spc(30); "Nombre      : "; Mid(Rs_Consulta_Productos.rdoColumns("Nombre"), 1, 70)
            Grid_Reporte.AddItem "Nombre :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Nombre")
            Grid_Reporte.AddItem ""
            Print #1,
        End If
        Print #1, "________________________________________________________________________________________________________________________"
        Print #1,
        Print #1, " Fecha    Referencia  Movimiento   Estatus   Cantidad  No.Lote      Caducidad         Costo       Costo Compra    Total "
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Referencia" & Chr(9) & "Movimiento" & Chr(9) & "No. Lote" & Chr(9) & "Fecha_Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Precio Venta" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        Grid_Reporte.AddItem ""
        Print #1, "________________________________________________________________________________________________________________________"
        'Consulta los movimientos del producto en entradas
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas,Alm_Entradas_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas.Entrada_ID=Alm_Entradas_Detalles.Entrada_ID"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Detalles.Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Fecha_Recepcion_Factura>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Fecha_Recepcion_Factura<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas.Fecha_Recepcion_Factura,Alm_Entradas.Entrada_ID"
        Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        Total_Productos = 0
        With Rs_Consulta_Entradas
            If Not .EOF Then
                Print #1,
                Print #1, Spc(60); "ENTRADAS"
                Grid_Reporte.AddItem "ENTRADAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las entradas y suma el costo de las entradas y el total de piezas y guarda la información en el grid para su exportación a Excel
                While Not .EOF
                    If Not IsNull(.rdoColumns("Fecha_Caducidad")) Then
                        Fecha_Caducidad = .rdoColumns("Fecha_Caducidad")
                    Else
                        Fecha_Caducidad = ""
                    End If
                    If Not IsNull(.rdoColumns("No_Lote")) Then
                        Print #1, Format(.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yy"); _
                            Spc(1); .rdoColumns("Entrada_ID"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Entrada"), 1, 15); _
                            Spc(17 - Len(Mid(.rdoColumns("Tipo_Entrada"), 1, 15))); Mid(.rdoColumns("Estatus"), 1, 3); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#0.00"), 10); _
                            Spc(2); Mid(.rdoColumns("No_Lote"), 1, 15); _
                            Spc(16 - Len(Mid(.rdoColumns("No_Lote"), 1, 15))); Format(Fecha_Caducidad, "dd/MMM/yy"); _
                            Spc(14 - Len(Fecha_Caducidad)); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Entradas("Costo"), "#,##0.00"), 8); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo")), "#,###,##0.00"), 12)
                    Else
                        Print #1, Format(.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yy"); _
                            Spc(1); .rdoColumns("Entrada_ID"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Entrada"), 1, 15); _
                            Spc(17 - Len(Mid(.rdoColumns("Tipo_Entrada"), 1, 15))); Mid(.rdoColumns("Estatus"), 1, 3); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#0.00"), 10); _
                            Spc(2); ""; _
                            Spc(16); ""; _
                            Spc(27); Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Entradas("Costo"), "#,##0.00"), 8); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo")), "#,###,##0.00"), 12)
                    End If
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Entrada_ID") _
                        & Chr(9) & .rdoColumns("Tipo_Entrada") _
                        & Chr(9) & .rdoColumns("Estatus") _
                        & Chr(9) & .rdoColumns("No_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & "" _
                        & Chr(9) & .rdoColumns("Costo") _
                        & Chr(9) & .rdoColumns("Cantidad") * .rdoColumns("Costo")
                    Total_Entrada = Total_Entrada + Val(.rdoColumns("Cantidad") * (.rdoColumns("Costo")))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Entradas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Entradas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(29); Conectar_Ayudante.Alinea_Derecha(Format(Total_Productos, "#######0.00"), 13); Spc(61); Conectar_Ayudante.Alinea_Derecha(Format(Total_Entrada, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL ENTRADAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Entrada
        'Consulta los movimientos del producto en salidas
        Mi_SQL = "SELECT DISTINCT Alm_Salidas_Almacen.No_Salida,Alm_Salidas_Almacen.Fecha_Salida,Alm_Salidas_Almacen.Tipo_Salida,Alm_Salidas_Almacen.Estatus,Alm_Salidas_Almacen_Detalles.No_Salida,Alm_Salidas_Almacen_Detalles.Producto_ID"
        Mi_SQL = Mi_SQL & " ,Alm_Salidas_Almacen_Detalles.Cantidad,Alm_Salidas_Almacen_Detalles.Precio_Venta"
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen,Alm_Salidas_Almacen_Detalles,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen.No_Salida=Alm_Salidas_Almacen_Detalles.No_Salida"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.Producto_ID ='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Fecha_Salida>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Fecha_Salida<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " ORDER BY Alm_Salidas_Almacen.Fecha_Salida,Alm_Salidas_Almacen.No_Salida"
        Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Salida = 0
        Total_Productos = 0
        With Rs_Consulta_Salidas
            If Not .EOF Then
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(60); "SALIDAS"
                Grid_Reporte.AddItem "SALIDAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las salidas y suma el costo de las salidas y el total de piezas
                While Not .EOF
                    If Not IsNull(.rdoColumns("Precio_Venta")) Then
                        Costo = Val(.rdoColumns("Precio_Venta"))
                    Else
                        Costo = 0
                    End If
                    Print #1, Format(.rdoColumns("Fecha_Salida"), "dd/MMM/yy"); _
                    Spc(1); .rdoColumns("No_Salida"); _
                    Spc(1); Mid(.rdoColumns("Tipo_Salida"), 1, 15); _
                    Spc(17 - Len(Mid(.rdoColumns("Tipo_Salida"), 1, 15))); Mid(.rdoColumns("Estatus"), 1, 3); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "######0.00"), 10); _
                    Spc(2); ""; _
                    Spc(16); ""; _
                    Spc(9); _
                    Spc(7); Conectar_Ayudante.Alinea_Derecha(Format(Costo, "#,##0.00"), 8); _
                    Spc(7); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Cantidad") * Costo), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha_Salida"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Salida") _
                        & Chr(9) & .rdoColumns("Tipo_Salida") _
                        & Chr(9) & "" _
                        & Chr(9) & "" _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & "" _
                        & Chr(9) & Costo _
                        & Chr(9) & Val(.rdoColumns("Cantidad")) * Costo
                    Total_Salida = Total_Salida + Val(.rdoColumns("Cantidad") * (Costo))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Salidas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Salidas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(29); Conectar_Ayudante.Alinea_Derecha(Format(Total_Productos, "#######0.00"), 13); Spc(61); Conectar_Ayudante.Alinea_Derecha(Format(Total_Salida, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL SALIDAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Salida
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime la existencia de los productos y saca la diferencia de costos de las entradas y salidas
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(34); "Existencia          : "; Conectar_Ayudante.Alinea_Derecha(Format(Rs_Consulta_Productos.rdoColumns("Existencia"), "#0.00"), 13); Spc(25); "Total    :"; Conectar_Ayudante.Alinea_Derecha(Format(Total_Entrada - Total_Salida, "#,###,###,##0.00"), 16)
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Existencia : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Existencia") & Chr(9) & Chr(9) & Chr(9) & Total_Entrada - Total_Salida
        End If
        Close #1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
        Exit Sub
        Rs_Consulta_Productos.Close
    Else
        MsgBox "Seleccione el producto a mostrar", vbExclamation
        Dtp_Kardex_Fecha_Inicial.Value = Now
        Dtp_Kardex_Fecha_Final.Value = Now
    End If
    Exit Sub
handler:
     Close #1
     Me.MousePointer = 0
     MsgBox Err.Description, vbInformation
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_Pedidos_Facturas
'DESCRIPCIÓN                : Muestra el reporte de las pedidos con las respectivas facturas
'CREO                       : Julio Cruz
'FECHA_CREO                 : 08-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'PARÁMETROS                 :
'*******************************************************************************
Private Sub Reporte_Pedidos_Facturas()
Dim Mi_SQL As String
Dim Rs_Consulta_Ope_Pedidos As rdoResultset
Dim Rs_Consulta_Salidas_Facturas As rdoResultset
Dim Empresa As String
Dim Total_General As Double
Dim Rs_Clientes_Nonbre As rdoResultset
Dim Estatus As String

On Error GoTo handler

    'Coloca el encabezado de los reportes
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE PEDIDOS POR DOCUMENTO", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
    Print #1, ""
    Grid_Reporte.AddItem ""
    'Consulta los pedidos de acuerdo al rango de fechas proporcionadas
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido BETWEEN " & Par_Fecha & Format(DTP_Fecha_Inicial_Reporte_documentos_Pedidos.Value, "MM/dd/yyyy") & Par_Fecha & " AND " & Par_Fecha & Format(DTP_Fecha_Fin_Reporte_documentos_Pedidos.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Pedido_ID"
    Set Rs_Consulta_Ope_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Ope_Pedidos.EOF
        Mi_SQL = "SELECT DISTINCT Alm_Salidas_Almacen.No_Salida,Adm_Clientes_Facturas.No_Factura,Adm_Clientes_Facturas.Total"
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen,Alm_Salidas_Almacen_Detalles,Adm_Clientes_Facturas"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen.No_Salida=Alm_Salidas_Almacen_Detalles.No_Salida"
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.No_Salida=Alm_Salidas_Almacen.No_Salida"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Pedido_ID ='" & Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID") & "'"
        Set Rs_Consulta_Salidas_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Salidas_Facturas.EOF Then
            'Imprime el encabezado del pedido
            Print #1,
            Grid_Reporte.AddItem ""
            If Not IsNull(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")) Then
                Estatus = Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")
             Else
              Estatus = ""
            End If
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Estatus, 1, 10); Spc(12 - Len(Mid(Estatus, 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Nombre"), 1, 30)
            Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Estatus & Chr(9) & "Cliente:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Nombre")
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "    No. Documento                                 Importe"
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Importe"
            Print #1, "________________________________________________________________________________________________________________________"
            With Rs_Consulta_Salidas_Facturas
                While Not .EOF
                    Print #1, Spc(2); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Factura"), 7); Spc(3), Empresa; Spc(20); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Total")), "#,###,##0.00"), 7)
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                    Total_General = Total_General + .rdoColumns("Total")
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Salidas_Facturas.Close
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(30); "GRAN TOTAL:       "; Format(Total_General, "#,###,##0.00")
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_General, "#,###,##0.00")
        Else    'Checa si lo tiene en remisiones
            Mi_SQL = "SELECT DISTINCT Alm_Salidas_Almacen.No_Salida,Adm_Clientes_Remisiones.No_Remision,Adm_Clientes_Remisiones.Total"
            Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen,Alm_Salidas_Almacen_Detalles,Adm_Clientes_Remisiones"
            Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen.No_Salida=Alm_Salidas_Almacen_Detalles.No_Salida"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.No_Salida= Alm_Salidas_Almacen_Detalles.No_Salida"
            Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Pedido_ID='" & Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID") & "'"
            Set Rs_Consulta_Salidas_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Salidas_Facturas.EOF Then
                'Imprime el encabezado del pedido
                Print #1,
                Grid_Reporte.AddItem ""
                If Not IsNull(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")) Then
                    Estatus = Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")
                 Else
                  Estatus = ""
                End If
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Estatus, 1, 10); Spc(12 - Len(Mid(Estatus, 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Nombre"), 1, 30)
                Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Estatus & Chr(9) & "Cliente:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Nombre")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Documento                              Importe"
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Importe"
                Print #1, "________________________________________________________________________________________________________________________"
                With Rs_Consulta_Salidas_Facturas
                    While Not .EOF
                        Print #1, Spc(2); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Remision"), 7); Spc(3), Empresa; Spc(20); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Total")), "#,###,##0.00"), 7)
                        Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                        Total_General = Total_General + .rdoColumns("Total")
                        .MoveNext
                    Wend
                End With
                Rs_Consulta_Salidas_Facturas.Close
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(30); "GRAN TOTAL:       "; Format(Total_General, "#,###,##0.00")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_General, "#,###,##0.00")
            Else
                'Imprime el encabezado del pedido
                 Print #1,
                 Grid_Reporte.AddItem ""
                 Print #1, "------------------------------------------------------------------------------------------------------------------------"
                 Mi_SQL = " SELECT * FROM Cat_Clientes WHERE Cliente_ID='" & Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente_ID") & "' "
                 Set Rs_Clientes_Nonbre = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                 If Not IsNull(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")) Then
                    Estatus = Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus")
                 Else
                  Estatus = ""
                End If
                 Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Estatus, 1, 10); Spc(12 - Len(Mid(Estatus, 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Clientes_Nonbre!Nombre, 1, 30)
                 Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Pedido_ID") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Estatus & Chr(9) & "Cliente:" & Chr(9) & Rs_Clientes_Nonbre!Nombre
                 Print #1, "------------------------------------------------------------------------------------------------------------------------"
                 Print #1, "    No. Factura                              Importe"
                 Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Importe"
                 Print #1, "________________________________________________________________________________________________________________________"
                 Print #1, Spc(2); "SIN FACTURAS NI REMISIONES"
                 Grid_Reporte.AddItem "SIN FACTURAS"
            End If
        End If
        Rs_Consulta_Ope_Pedidos.MoveNext
    Wend
    Rs_Consulta_Ope_Pedidos.Close
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 13250
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_General_Facturas
'DESCRIPCIÓN                : Consulta las factutas que fueron dadas de alta
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Reporte_General_Facturas()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Total_IVA As Double                                 'Almacena el total del IVA
Dim Total_Tasa_0 As Double                              'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                             'Almacena el total de la tasa 15
Dim Total_Facturas As Double                            'Almacena el total de la factura
Dim Iniciales As String

On Error GoTo handler
    Iniciales = Consulta_Usuario
    Grid_Reporte.Cols = 11
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Reporte_General_Facturas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"  'AND Serie='A' ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            'Valida que exista la ruta si no la crea
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE GENERAL DE FACTURAS ", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("DEL " & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas, "dd/MMM/yyyy HH:mm:ss") & " AL " & Format(Dtp_Fecha_Fin_Reporte_General_Facturas, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
            Print #1, ""
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "REPORTE DE GENERAL DE FACTURAS"
            Grid_Reporte.AddItem ""
            Print #1,
            Grid_Reporte.AddItem "DEL " & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Fecha_Fin_Reporte_General_Facturas.Value, "dd/MMM/yyyy")
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Factura     Serie    Documento     Cliente                         Fecha          Abono         Saldo          IVA         Total      Estatus               UUID"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Serie" & Chr(9) & "Documento" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus" & Chr(9) & "UUID"
            Grid_Reporte.AddItem ""
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            While Not .EOF
                Estatus = ""
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
                If Trim(.rdoColumns("Pagada")) = "S" Then
                    Estatus = "Pagada"
                End If
                If Trim(.rdoColumns("Cancelada")) = "S" Then
                    Estatus = "Cancelada"
                End If
                If Trim(.rdoColumns("Pagada")) = "N" And Trim(.rdoColumns("Cancelada")) = "N" Then
                    Estatus = "Pendiente"
                End If
                If Trim(.rdoColumns("Cancelada")) = "PC" Then
                    Estatus = "Proceso"
                End If
                If Trim(.rdoColumns("Cancelada")) = "N" Then
                  If Trim(.rdoColumns("Serie")) = "NCA" Then
                    Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(6); Conectar_Ayudante.Alinea_Derecha(Trim(.rdoColumns("Serie")), 1); Spc(8); Conectar_Ayudante.Alinea_Derecha(Trim(Trim(.rdoColumns("Tipo_Documento"))), 1); Spc(3); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Estatus, 10); Spc(4); .rdoColumns("Timbre_UUID")
                  Else
                    Print #1, Spc(1); .rdoColumns("No_Factura_Electronica"); Spc(8); Conectar_Ayudante.Alinea_Derecha(Trim(.rdoColumns("Serie")), 1); Spc(8); Conectar_Ayudante.Alinea_Derecha(Trim(Trim(.rdoColumns("Tipo_Documento"))), 1); Spc(6); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Estatus, 10); Spc(4); .rdoColumns("Timbre_UUID")
                 End If
                    Grid_Reporte.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Tipo_Documento") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("IVA"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus & Chr(9) & .rdoColumns("Timbre_UUID") & Chr(13)
                    Total_IVA = Total_IVA + .rdoColumns("IVA")
                    Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Abono")
                    Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Saldo")
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Else
                    Print #1, Spc(1); .rdoColumns("No_Factura_Electronica"); Spc(8); Trim(.rdoColumns("Serie")); Spc(8); Trim(.rdoColumns("Tipo_Documento")); Spc(6); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Estatus, 10); Spc(4); .rdoColumns("Timbre_UUID")
                    Grid_Reporte.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Tipo_Documento") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus & Chr(9) & .rdoColumns("Timbre_UUID")
                End If
                .MoveNext
            Wend
            Print #1,
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(55); Conectar_Ayudante.Alinea_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 13); Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 13); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_IVA, "###,##0.00"), 10); Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Grid_Reporte.AddItem "Usuario " & Iniciales
            Print #1,
            Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Alinea_Derecha("Usuario " & Iniciales, 10)
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 17235
            Me.Height = 8175
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

Public Function Consulta_Usuario()
Dim Rs_Consulta_Cat_Usuarios As rdoResultset            'Manejo de registro
Dim Cadena As String
Dim Iniciales() As String
Dim I As Long


On Error GoTo handler
    Mi_SQL = "SELECT Nombre "
    Mi_SQL = Mi_SQL & " FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & Usuario_ID & "'"
    Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Usuarios.EOF Then
        With Rs_Consulta_Cat_Usuarios
            Cadena = .rdoColumns("Nombre")
        End With
        Rs_Consulta_Cat_Usuarios.Close
        Iniciales = Split(Cadena, " ")
        Cadena = ""
        For I = 0 To UBound(Iniciales)
            Cadena = Cadena + UCase(Mid(Iniciales(I), 1, 1))
        Next
        Consulta_Usuario = Cadena
    Else
        Consulta_Usuario = ""
    End If
    Exit Function
handler:
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Function
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_Facturas_Vencidas
'DESCRIPCIÓN                : Consulta las factutas vencidas
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Reporte_Facturas_Vencidas()
Dim Mi_SQL As String
Dim Proveedor As String                                     'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Total_Facturas As Double                                'Almacena el total de las facturas
Dim Saldo_Facturas As Double
Dim Abono_Facturas As Double


On Error GoTo handler
    
    Catalogo = ""
    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Pagada='N'"
    Mi_SQL = Mi_SQL & " AND Fecha_Pago<'" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Proveedores.EOF Then
    With Rs_Consulta_Cat_Facturas_Proveedores
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE FACTURAS VENCIDAS", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
            Print #1, ""
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, " No Factura               Proveedor                      Fecha     Fecha Pago      Total         Saldo          Abono    "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1,
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Fecha Factura" & Chr(9) & "Fecha Pago" & Chr(9) & "Total" & Chr(9) & "Saldo" & Chr(9) & "Abono"
        Grid_Reporte.AddItem ""
        While Not .EOF
            If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores"
                Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Proveedores.EOF Then
                    Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                Else
                    Proveedor = ""
                End If
                Rs_Consulta_Cat_Proveedores.Close
            End If
            Print #1, Mid(Trim(.rdoColumns("No_Factura")), 1, 20); _
                Spc(22 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 20))); Mid(Proveedor, 1, 32); _
                Spc(34 - Len(Mid(Proveedor, 1, 32))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(2); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12)
            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Proveedor & Chr(9) & _
                Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                Format(.rdoColumns("Abono"), "#,###,##0.00")
            Total_Facturas = Total_Facturas + Val(.rdoColumns("Total"))
            Saldo_Facturas = Saldo_Facturas + Val(.rdoColumns("Saldo"))
            Abono_Facturas = Abono_Facturas + Val(.rdoColumns("Abono"))
            .MoveNext
        Wend
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(60); "TOTALES : "; Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Saldo_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Abono_Facturas, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTALES :" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Saldo_Facturas, "#,###,##0.00") & Chr(9) & Format(Abono_Facturas, "#,###,##0.00")
        Close #1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
        End With
    Else
        MsgBox "No hay facturas Vencidas", vbInformation
        Catalogo = "No_Facturas_Vencidas"
    End If
    Rs_Consulta_Cat_Facturas_Proveedores.Close
    Exit Sub
handler:
    Close #1
    MsgBox Err.Description, vbInformation
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_Facturas_Vencer
'DESCRIPCIÓN                : Consulta las factutas Por vencer
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Reporte_Facturas_Vencer()
Dim Mi_SQL As String
Dim Proveedor As String                                     'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Total_Facturas As Double                                 'Almacena el total de las facturas
Dim Saldo_Facturas As Double
Dim Abono_Facturas As Double

        
On Error GoTo handler

    Catalogo = ""
    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Pagada='N'"
    Mi_SQL = Mi_SQL & " AND Fecha_Pago= '" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Proveedores.EOF Then
        With Rs_Consulta_Cat_Facturas_Proveedores
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE FACTURAS POR VENCER", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto(Format(Now, "dd/MMM/yyyy HH:mm:ss"), Len("***************************************************************************************************************"))
            Print #1, ""
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " No Factura   Proveedor                                Fecha     Fecha Pago      Total         Saldo          Abono      "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Fecha Factura" & Chr(9) & "Fecha Pago" & Chr(9) & "Total" & Chr(9) & "Saldo" & Chr(9) & "Abono"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores"
                    Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                    Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Proveedores.EOF Then
                        Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                    Else
                        Proveedor = ""
                    End If
                    Rs_Consulta_Cat_Proveedores.Close
                End If
                Print #1, Mid(Trim(.rdoColumns("No_Factura")), 1, 12); Spc(14 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 12))); Mid(Proveedor, 1, 40); _
                    Spc(42 - Len(Mid(Proveedor, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12)
                Grid_Reporte.AddItem Trim(.rdoColumns("No_Factura")) & Chr(9) & Proveedor & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Abono"), "#,###,##0.00")
                Total_Facturas = Total_Facturas + Val(.rdoColumns("Total"))
                Saldo_Facturas = Saldo_Facturas + Val(.rdoColumns("Saldo"))
                Abono_Facturas = Abono_Facturas + Val(.rdoColumns("Abono"))
                .MoveNext
            Wend
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(60); "TOTALES : "; Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Saldo_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Abono_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTALES :" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Saldo_Facturas, "#,###,##0.00") & Chr(9) & Format(Abono_Facturas, "#,###,##0.00")
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay facturas por Vencer", vbInformation
        Catalogo = "No_Facturas_Vencer"
    End If
    Rs_Consulta_Cat_Facturas_Proveedores.Close
    Exit Sub
handler:
    Close #1
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_General_Remisiones
'DESCRIPCIÓN                : Consulta las remisiones
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Reporte_General_Remisiones()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Total_IVA As Double                                 'Almacena el total del IVA
Dim Total_Tasa_0 As Double                              'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                             'Almacena el total de la tasa 15
Dim Total_Facturas As Double                            'Almacena el total de la factura

On Error GoTo handler

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Reporte_General_Facturas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            Open App.Path & "\REP\Reportes.txt" For Output As #1
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
            Print #1,
            Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE GENERAL DE REMISIONES", Len("***************************************************************************************************************"))
            Print #1, Conectar_Ayudante.Centrar_Texto("DEL " & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Fecha_Fin_Reporte_General_Facturas.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Print #1, ""
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "REPORTE DE GENERAL DE REMISIONES"
            Grid_Reporte.AddItem ""
            Print #1,
            Grid_Reporte.AddItem "DEL " & Format(Dtp_Fecha_Inicio_Reporte_General_Facturas.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Fecha_Fin_Reporte_General_Facturas.Value, "dd/MMM/yyyy")
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " No. Remision     Cliente                       Fecha        Abono        Saldo          IVA         Total     Estatus   "
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem ""
           
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            While Not .EOF
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
                If Trim(.rdoColumns("Cancelada")) = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Trim(.rdoColumns("Pagada")) = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Pendiente"
                    End If
                End If
                If Trim(.rdoColumns("Cancelada")) = "N" Then
                    Print #1, Spc(1); .rdoColumns("No_Remision"); Spc(2); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("IVA"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                    Total_IVA = Total_IVA + .rdoColumns("IVA")
                    Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Abono")
                    Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Saldo")
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Else
                    Print #1, Spc(1); .rdoColumns("No_Remision"); Spc(2); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha("0.00", 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                End If
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(26); Conectar_Ayudante.Alinea_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 13); Spc(1); Conectar_Ayudante.Alinea_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 13); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_IVA, "###,##0.00"), 10); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13250
            Me.Height = 8200
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay remisiones en las fechas que proporcionó", vbExclamation
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Reporte_Remisiones_Productos
'DESCRIPCIÓN                : Consulta las remisiones y facturas por producto
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Febrero-2011
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Reporte_Remisiones_Productos()
Dim Mi_SQL As String
Dim Rs_Consulta_Datos_Factura As rdoResultset
Dim Rs_Consulta_Datos_Pedido As rdoResultset
Dim Empresa As String
Dim No_Pedido As String
Dim Total As Double
Dim Total_Producto As Double
Dim Rs_Consulta_Productos As rdoResultset
Dim Cont_Productos As Double
Dim Rs_Consulta_Categoria As rdoResultset
Dim Gran_total As Double
Dim Gran_Total_Piezaz As Double
Dim Cont_Partidas As Double

On Error GoTo handler

    Cont_Productos = 0
    Gran_total = 0
    Gran_Total_Piezaz = 0
    Mi_SQL = " SELECT"
    Mi_SQL = Mi_SQL & " Cat_Productos.*,Cat_Productos_Tipo.Nombre as Tipo_Producto FROM Cat_Productos,Cat_Productos_Tipo "
    Mi_SQL = Mi_SQL & " WHERE Cat_Productos.Tipo_ID = Cat_Productos_Tipo.Tipo_ID"
    If Cmb_Producto_Reporte_Producto_Por_Documento.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID='" & Format(Cmb_Producto_Reporte_Producto_Por_Documento.ItemData(Cmb_Producto_Reporte_Producto_Por_Documento.ListIndex), "00000") & "'"
    If Cmb_Tipo_Producto_Producto_Por_Documento.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Productos.Tipo_ID='" & Format(Cmb_Tipo_Producto_Producto_Por_Documento.ItemData(Cmb_Tipo_Producto_Producto_Por_Documento.ListIndex), "00000") & "'"
    If Cmb_Categorias.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Productos.Categoria_ID='" & Format(Cmb_Categorias.ItemData(Cmb_Categorias.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " Order by Cat_Productos_Tipo.Tipo_ID"
    Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Productos.EOF Then
        While Not Rs_Consulta_Productos.EOF
            Cont_Partidas = 0
            Total_Producto = 0
            Total = 0
            If Cont_Productos = 0 Then
                Open App.Path & "\REP\Reportes.txt" For Output As #1
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
                Print #1,
                Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE VENTAS", Len("***************************************************************************************************************"))
                Print #1, Conectar_Ayudante.Centrar_Texto("DEL " & Format(Dtp_Fecha_Inicio_Reporte_Producto_Por_Documento.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Fecha_Fin_Reporte_Producto_Por_Documento.Value, "dd/MMM/yyyy"), Len("***************************************************************************************************************"))
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
            End If
            'SE AGREGAN LAS REMISIONES
            Mi_SQL = "SELECT Adm_Clientes_Remisiones.No_Remision,Adm_Clientes_Remisiones.Fecha,Adm_Clientes_Remisiones.Cliente_ID"
            Mi_SQL = Mi_SQL & " ,Adm_Clientes_Remisiones_Detalles.Cantidad,Adm_Clientes_Remisiones_Detalles.Precio,Adm_Clientes_Remisiones_Detalles.Importe,Cat_Clientes.Nombre"
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones,Adm_Clientes_Remisiones_Detalles,Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Remisiones.No_Remision=Adm_Clientes_Remisiones_Detalles.No_Remision"
            If Cmb_Reporte_ventas_Cliente.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID='" & Format(Cmb_Reporte_ventas_Cliente.ItemData(Cmb_Reporte_ventas_Cliente.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones_Detalles.Producto_ID='" & Format(Rs_Consulta_Productos!Producto_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Fecha>='" & Format(Dtp_Fecha_Inicio_Reporte_Producto_Por_Documento.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Fecha<='" & Format(Dtp_Fecha_Fin_Reporte_Producto_Por_Documento.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cancelada='N'"
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Remisiones.Fecha"
            Set Rs_Consulta_Datos_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Datos_Factura
                While Not .EOF
                    If Cont_Partidas = 0 Then
                        'SE CONSULTA LA CATEGORIA DEL PRODUCTO
                        Mi_SQL = " SELECT Nombre as Categoria_Producto FROM Cat_Categorias WHERE Categoria_ID='" & Format(Rs_Consulta_Productos!Categoria_ID, "00000") & "' "
                        Set Rs_Consulta_Categoria = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        Print #1, "TIPO PRODUCTO: "; Rs_Consulta_Productos!Tipo_Producto
                        Print #1, "CATEGORIA PRODUCTO: "; Rs_Consulta_Categoria!Categoria_Producto
                        Print #1, ""
                        Print #1, "PRODUCTO : "; Rs_Consulta_Productos!Nombre
                        Grid_Reporte.AddItem "TIPO PRODUCTO : " & Chr(9) & Rs_Consulta_Productos!Tipo_Producto
                        Grid_Reporte.AddItem "CATEGORIA PRODUCTO:  " & Chr(9) & Rs_Consulta_Categoria!Categoria_Producto
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "PRODUCTO : " & Chr(9) & Rs_Consulta_Productos!Nombre
                        Rs_Consulta_Categoria.Close
                        Print #1,
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1, " Cliente                                    Documento        Fecha         Pedido     Cantidad    Precio       Importe  "
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "Cliente" & Chr(9) & "Documento" & Chr(9) & "Fecha" & Chr(9) & "Pedido" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                        Grid_Reporte.AddItem ""
                    End If
                    'Consulta el número del pedido
                    Mi_SQL = "SELECT DISTINCT Alm_Salidas_Almacen_Detalles.No_Salida,Alm_Salidas_Almacen.Pedido_ID"
                    Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen_Detalles,Alm_Salidas_Almacen"
                    Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen_Detalles.No_Salida=Alm_Salidas_Almacen.No_Salida"
                    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.No_Salida='" & .rdoColumns("No_Remision") & "'"
                    Set Rs_Consulta_Datos_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Datos_Pedido.EOF Then
                        If Not IsNull(Rs_Consulta_Datos_Pedido.rdoColumns("Pedido_ID")) Then
                            No_Pedido = Rs_Consulta_Datos_Pedido.rdoColumns("Pedido_ID")
                        Else
                            No_Pedido = "          "
                        End If
                    Else
                        No_Pedido = "          "
                    End If
                    Rs_Consulta_Datos_Pedido.Close
                    Print #1, Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); .rdoColumns("No_Remision"); " "; Empresa; _
                    Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(3); No_Pedido; _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0"), 7); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Precio")), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Importe")), "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("No_Remision") & " " & Empresa _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & No_Pedido _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0") _
                    & Chr(9) & Format(.rdoColumns("Precio"), "#,###,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Importe"), "#,###,##0.00")
                    Total_Producto = Total_Producto + Val(.rdoColumns("Cantidad"))
                    Total = Total + Val(.rdoColumns("Importe"))
                    Cont_Partidas = 1
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Datos_Factura.Close
            'SE AGREGAN LAS FACTURAS
            Mi_SQL = "SELECT Adm_Clientes_Facturas.No_Salida,Adm_Clientes_Facturas.No_Factura,Adm_Clientes_Facturas.Fecha,Adm_Clientes_Facturas.Cliente_ID"
            Mi_SQL = Mi_SQL & " ,Adm_Clientes_Facturas_Detalles.Cantidad,Adm_Clientes_Facturas_Detalles.Precio,Adm_Clientes_Facturas_Detalles.Importe,Cat_Clientes.Nombre"
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas,Adm_Clientes_Facturas_Detalles,Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Facturas.No_Factura=Adm_Clientes_Facturas_Detalles.No_Factura"
            If Cmb_Reporte_ventas_Cliente.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID='" & Format(Cmb_Reporte_ventas_Cliente.ItemData(Cmb_Reporte_ventas_Cliente.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas_Detalles.Producto_ID='" & Format(Rs_Consulta_Productos!Producto_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Fecha>='" & Format(Dtp_Fecha_Inicio_Reporte_Producto_Por_Documento.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Fecha<='" & Format(Dtp_Fecha_Fin_Reporte_Producto_Por_Documento.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.No_Salida IS NOT NULL "
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Facturas.Fecha"
            Set Rs_Consulta_Datos_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Datos_Factura
                While Not .EOF
                    If Cont_Partidas = 0 Then
                        'SE CONSULTA LA CATEGORIA DEL PRODUCTO
                        Mi_SQL = " SELECT Nombre as Categoria_Producto FROM Cat_Categorias WHERE Categoria_ID='" & Format(Rs_Consulta_Productos!Categoria_ID, "00000") & "' "
                        Set Rs_Consulta_Categoria = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        Print #1, "TIPO PRODUCTO: "; Rs_Consulta_Productos!Tipo_Producto
                        Print #1, "CATEGORIA PRODUCTO: "; Rs_Consulta_Categoria!Categoria_Producto
                        Print #1, ""
                        Print #1, "PRODUCTO : "; Rs_Consulta_Productos!Nombre
                        Grid_Reporte.AddItem "TIPO PRODUCTO : " & Chr(9) & Rs_Consulta_Productos!Tipo_Producto
                        Grid_Reporte.AddItem "CATEGORIA PRODUCTO:  " & Chr(9) & Rs_Consulta_Categoria!Categoria_Producto
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "PRODUCTO : " & Chr(9) & Rs_Consulta_Productos!Nombre
                        Rs_Consulta_Categoria.Close
                        Print #1,
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1, " Cliente                                    Documento        Fecha         Pedido     Cantidad    Precio       Importe  "
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "Cliente" & Chr(9) & "Documento" & Chr(9) & "Fecha" & Chr(9) & "Pedido" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                        Grid_Reporte.AddItem ""
                    End If
                    Grid_Reporte.AddItem ""
                    'Consulta los datos de No_Pedido y Orden_Compra
                    Mi_SQL = "SELECT DISTINCT Alm_Salidas_Almacen_Detalles.No_Salida,Alm_Salidas_Almacen.Pedido_ID"
                    Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Almacen_Detalles,Alm_Salidas_Almacen"
                    Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Almacen_Detalles.No_Salida=Alm_Salidas_Almacen.No_Salida"
                    Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen_Detalles.No_Salida='" & .rdoColumns("No_Salida") & "'"
                    Set Rs_Consulta_Datos_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Datos_Pedido.EOF Then
                        If Not IsNull(Rs_Consulta_Datos_Pedido.rdoColumns("Pedido_ID")) Then
                            No_Pedido = Rs_Consulta_Datos_Pedido.rdoColumns("Pedido_ID")
                        Else
                            No_Pedido = "          "
                        End If
                    Else
                        No_Pedido = "          "
                    End If
                    Rs_Consulta_Datos_Pedido.Close
                    Print #1, Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); .rdoColumns("No_Factura"); " "; Empresa; _
                    Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(3); No_Pedido; _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0"), 7); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Precio")), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Val(.rdoColumns("Importe")), "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("No_Factura") & " " & Empresa _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & No_Pedido _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0") _
                    & Chr(9) & Format(.rdoColumns("Precio"), "#,###,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Importe"), "#,###,##0.00")
                    Total_Producto = Total_Producto + Val(.rdoColumns("Cantidad"))
                    Total = Total + Val(.rdoColumns("Importe"))
                    Cont_Partidas = 1
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Datos_Factura.Close
            Gran_total = Gran_total + Total
            Gran_Total_Piezaz = Gran_Total_Piezaz + Total_Producto
            If Total_Producto > 0 Or Cont_Partidas = 1 Then
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(65); "TOTAL DE PIEZAS :  "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Producto, "#,###,##0"), 9)
                Print #1, Spc(65); "TOTAL :       "; Conectar_Ayudante.Alinea_Derecha(Format(Total, "#,###,##0.00"), 14)
                Print #1, "________________________________________________________________________________________________________________________"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL DE PIEZAS : " & Chr(9) & Format(Total_Producto, "###,##0")
                Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL :      " & Chr(9) & Format(Total, "#,###,##0.00")
            End If
            Total_Producto = 0
            Total = 0
            Cont_Productos = Cont_Productos + 1
            Rs_Consulta_Productos.MoveNext
        Wend
        Rs_Consulta_Productos.Close
        Print #1, ""
        ''Print #1, "________________________________________________________________________________________________________________________"
        Print #1, Spc(65); "GRAN TOTAL DE PIEZAS :  "; Conectar_Ayudante.Alinea_Derecha(Format(Gran_Total_Piezaz, "#,###,##0"), 9)
        Print #1, Spc(65); "GRAN TOTAL :       "; Conectar_Ayudante.Alinea_Derecha(Format(Gran_total, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL DE PIEZAS : " & Chr(9) & Format(Gran_Total_Piezaz, "###,##0")
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL :      " & Chr(9) & Format(Gran_total, "#,###,##0.00")
        Close #1
        Cmb_Producto_Reporte_Producto_Por_Documento.ListIndex = -1
        Cmb_Tipo_Producto_Producto_Por_Documento.ListIndex = -1
        Cmb_Categorias.ListIndex = -1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
    Else
        MsgBox "No hay datos que mostrar", vbInformation
    End If
    Exit Sub
handler:
    Close #1
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Reporte_Cobranza
'DESCRIPCIÓN            : Muestra la cobranza
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 12-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Reporte_Cobranza()
Dim Cliente As String                                       'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Remisiones_Clientes As rdoResultset     'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_Factura As Double                                 'Almacena el total de la facura
Dim Total_Saldo As Double                                   'Almacena el total del saldo
Dim Total_Abono As Double                                   'Almacena el total de los abonos
Dim Fecha_Vencimiento As Date                               'Almacena la fecha de vencimiento de la factura
Dim Vencimiento As Integer                                  'Almacena los dias de vnecimiento de la factura
Dim Total_Vencimiento As Integer                            'Almacena los espacios para colocar la información
Dim Espacio As String
Dim Total_0_30 As Double
Dim Total_30_60 As Double
Dim Total_60_90 As Double
Dim Total_90 As Double
Dim Total_Cliente As Double
Dim Abono_Cliente As Double
Dim Saldo_Cliente As Double
Dim Cliente_ID As String
Dim Tipo_Pago As String
Dim Tem_Cliente As String

On Error GoTo handler

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    Open App.Path & "\REP\Reportes.txt" For Output As #1
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***************************************************************************************************************"))
    Print #1, Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***************************************************************************************************************"))
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto("REPORTE DE COBRANZA", Len("***************************************************************************************************************"))
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    
    'SE CONSULTAN LOS CLIENTES
    Mi_SQL = " SELECT Distinct Cat_Clientes.Cliente_ID, Cat_Clientes.* "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Facturas.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Pagada='N'"
    Mi_SQL = Mi_SQL & " AND Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " UNION"
    Mi_SQL = Mi_SQL & " SELECT Distinct Cat_Clientes.Cliente_ID, Cat_Clientes.* "
    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones,Cat_Clientes "
    Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Remisiones.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Pagada='N'"
    Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre "
    Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Adm_Remisiones_Clientes.EOF
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID='" & Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") & "'"
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Clientes.EOF Then
            With Rs_Consulta_Cat_Clientes
                If Cliente_ID <> Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") Then
                    If Not IsNull(Rs_Consulta_Cat_Clientes!Ciudad) Then Ciudad = Rs_Consulta_Cat_Clientes!Ciudad
                    If Cliente_ID <> "" Then
                        Print #1, "_________________________________________________________________________________________________________________________"
                        Print #1, Spc(47); " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13)
                        Print #1,
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                        Grid_Reporte.AddItem ""
                    End If
                    Print #1,
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
                    Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion"); " "; .rdoColumns("Colonia")
                    Print #1, Spc(2); "Ciudad:    "; Ciudad; Spc(2); "CP: "; .rdoColumns("CP"); Spc(2); "Telefono: "; .rdoColumns("Telefono")
                    Print #1, Spc(2); "(*) Documentos de contado"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                    Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion") & Chr(9) & .rdoColumns("Colonia")
                    Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP") & Chr(9) & "Telefono" & Chr(9) & .rdoColumns("Telefono")
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " Documento             Fecha            Fecha Vencimiento                   Total                                                    "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Documento" & Chr(9) & "Fecha" & Chr(9) & "Fecha Vencimiento" & Chr(9) & "Total"
                    Cliente_ID = Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")
                    Total_Cliente = 0
                End If
            End With
        End If
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Facturas.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Pagada='N'"
        Mi_SQL = Mi_SQL & " AND Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cliente_ID='" & Format(Rs_Consulta_Adm_Remisiones_Clientes!Cliente_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Facturas.Cliente_ID"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
            With Rs_Consulta_Adm_Facturas_Clientes
                Print #1, Format(.rdoColumns("No_Factura_Electronica"), "0000000000"); Spc(10); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(15); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                Spc(15); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13)
                Grid_Reporte.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & _
                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Fecha_Pago"), "dd/MMMM/yyyy") & Chr(9) & _
                Format(.rdoColumns("Total"), "#,###,##0.00")
                Total_Factura = Total_Factura + .rdoColumns("Total")
                Total_Cliente = Total_Cliente + .rdoColumns("Total")
            End With
            Rs_Consulta_Adm_Facturas_Clientes.MoveNext
        Wend
        Rs_Consulta_Adm_Facturas_Clientes.Close
        If Not Trim(Tem_Cliente) = Trim(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")) Then
            'SE ASIGNA EL CLIENTE DEL QUE SE ESTA HACIENDO LA CONSULTA
            Tem_Cliente = Trim(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID"))
            'SE AGREGAN LAS REMISIONNES
            Mi_SQL = "SELECT * "
            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones,Cat_Clientes "
            Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Remisiones.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Pagada='N'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cliente_ID='" & Format(Rs_Consulta_Adm_Remisiones_Clientes!Cliente_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Clientes_Remisiones.Cliente_ID"
            Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                With Rs_Consulta_Adm_Facturas_Clientes
                    Print #1, Format(.rdoColumns("No_Remision"), "0000000000"); Spc(10); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(15); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                    Spc(15); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13)
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Fecha_Pago"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Total"), "#,###,##0.00")
                    Total_Factura = Total_Factura + .rdoColumns("Total")
                    Total_Cliente = Total_Cliente + .rdoColumns("Total")
                End With
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
            Rs_Consulta_Adm_Facturas_Clientes.Close
        End If
        Rs_Consulta_Adm_Remisiones_Clientes.MoveNext
    Wend
    Rs_Consulta_Adm_Remisiones_Clientes.Close
    Print #1, "_________________________________________________________________________________________________________________________"
    Print #1, Spc(47); " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13)
    Print #1,
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
    Grid_Reporte.AddItem ""
    Print #1, "_________________________________________________________________________________________________________________________"
    Print #1, Spc(47); " TOTALES DOCUMENTOS"; Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Factura, "#,###,##0.00"), 13)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "TOTALES DOCUMENTOS" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Factura, "#,###,##0.00")
    Close #1
    Me.Edm_Reporte.Font.Size = 8
    Me.Edm_Reporte.Font = "Courier New"
    Me.Edm_Reporte.fileName = App.Path & "\REP\Reportes.txt"
    Me.Edm_Reporte.Visible = True
    Ancho_Frame = Me.Width
    Largo_Frame = Me.Height
    Me.Width = 17450
    Me.Height = 8200
    Me.Left = 0
    Me.Top = 0
    Exit Sub
handler:
    MsgBox Err.Description
    Close #1
    Me.MousePointer = 0
End Sub
Private Sub Reporte_Cliente_Estadistico()
''Dim Mi_SQL As String
''Dim Rs_Consulta_Cat_Clientes As rdoResultset
''Dim Rs_Cat_Clientes As rdoResultset
''Dim Rs_Consulta_Fecha_Factura As rdoResultset
''Dim Rs_Adm_Facturas_Clientes As rdoResultset
''Dim Rs_Adm_Notas_Credtio_Clientes As rdoResultset
''Dim Rs_Consulta_Clientes As rdoResultset
''Dim Rs_Consulta_Ciudades As rdoResultset
''Dim Ciudad As String
''Dim Total_Nota_Credito As Double
''Dim Total_Pagos As Double
''Dim Compra_Total As Double
''Dim Compra_Maxima As Double
''Dim Compra_Minima As Double
''Dim Cheque_Devuelto As Double
''Dim Meses_Trasncurridos As Integer
''Dim Puntos_Gestion As Integer
''Dim Nivel_Gestion As String
''
''    If Cmb_Cliente_Estadistico.ListIndex > -1 Then
''        Grid_Reporte.Cols = 7
''        Grid_Reporte.Rows = 0
''        Call Encabezado_Reportes("ESTADISTICO DE CLIENTES", Format(Now, "dd/MMM/yyyy"), "", True)
''        Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad"
''        Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Ciudades"
''        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Nombre='" & Cmb_Cliente_Estadistico.Text & "'"
''        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''        With Rs_Consulta_Clientes
''            If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
''                Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
''                Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
''                Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
''                Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
''                Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''                If Not Rs_Consulta_Ciudades.EOF Then
''                    Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
''                Else
''                    Ciudad = ""
''                End If
''                Rs_Consulta_Ciudades.Close
''            Else
''                Ciudad = ""
''            End If
''            Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Estadistico.Text
''            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
''            Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
''            Print #1, Spc(2); "Ciudad:    "; Ciudad
''            Grid_Reporte.AddItem ""
''            Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Estadistico.Text
''            Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
''            Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
''            Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
''        End With
''        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
''        Print #1, "     Estadístico         Limite            Notas          Compra           Compra           Cargos          Nivel  "
''        Print #1, "     Pago X Mes          Crédito          Crédito         Máxima           Mínima          Bancarios       Gestión "
''        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
''        Grid_Reporte.AddItem ""
''        Grid_Reporte.AddItem "Estadístico Pago X Mes" & Chr(9) & "Límte de Crédito" & Chr(9) & "Notas de Crédito" & Chr(9) & "Compra Máxima" & Chr(9) & "Compra Mínima" & Chr(9) & "Cargos Bancarios" & Chr(9) & "Nivel de Gestión"
''        Grid_Reporte.AddItem ""
''        'Consulta los clientes
''        Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes "
''        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente_Estadistico.ItemData(Cmb_Cliente_Estadistico.ListIndex), "00000") & "' ORDER BY Nombre"
''        Set Rs_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''        While Not Rs_Cat_Clientes.EOF
''            Puntos_Gestion = 0
''            'Consulta la fecha de la primera compra del cliente y calcula el total de meses transcurridos
''            Mi_SQL = "SELECT TOP 1 Fecha FROM Adm_Facturas_Clientes"
''            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
''            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
''            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
''            Set Rs_Consulta_Fecha_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''            If Not Rs_Consulta_Fecha_Factura.EOF Then
''                Meses_Trasncurridos = DateDiff("M", Format(Rs_Consulta_Fecha_Factura.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
''            Else
''                Meses_Trasncurridos = 0
''            End If
''            Rs_Consulta_Fecha_Factura.Close
''            'Consulta el total pagado, el máximo y el mínimo de compra del cliente
''            Mi_SQL = "SELECT SUM(Abono) AS Total_Pagos,MAX(Total) AS Maxima_Compra,MIN(Total) AS Minima_Compra,SUM(Total) AS Total_Compras"
''            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
''            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
''            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
''            Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''            If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")) Then
''                Total_Pagos = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")
''                Compra_Maxima = Rs_Adm_Facturas_Clientes.rdoColumns("Maxima_Compra")
''                Compra_Minima = Rs_Adm_Facturas_Clientes.rdoColumns("Minima_Compra")
''                Compra_Total = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Compras")
''                'Valida que el límite supere la cantidad adeudada para sumar un punto a nivel de gestion
''                If Meses_Trasncurridos > 0 Then
''                    If Val(Compra_Total - Total_Pagos) > (Total_Pagos / Meses_Trasncurridos * 2) Then
''                        Puntos_Gestion = Puntos_Gestion + 1
''                    End If
''                End If
''            Else
''                Total_Pagos = 0
''                Compra_Maxima = 0
''                Compra_Minima = 0
''            End If
''            Rs_Adm_Facturas_Clientes.Close
''            'Consulta las notas de crédito
''            Mi_SQL = "SELECT SUM(Importe) FROM Adm_Notas_Credito_Clientes"
''            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
''            Mi_SQL = Mi_SQL & " AND Autorizado='SI' AND Aplicada='SI'"
''            Set Rs_Adm_Notas_Credtio_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''            If Not IsNull(Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)) Then
''                Total_Nota_Credito = Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)
''                'Valida el 20 % de compras superen el total de notas para agregar un punto a nivel de gestion
''                If Val(Total_Nota_Credito / Compra_Total * 100) > 20 Then
''                    Puntos_Gestion = Puntos_Gestion + 1
''                End If
''            Else
''                Total_Nota_Credito = 0
''            End If
''            Rs_Adm_Notas_Credtio_Clientes.Close
''            'Consulta las comisiones por cheques devueltos
''            Mi_SQL = "SELECT SUM(Adm_Detalles_Facturas.Importe) AS Importe,SUM(Adm_Detalles_Facturas.IVA) AS Iva"
''            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas"
''            Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
''            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
''            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
''            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
''            Mi_SQL = Mi_SQL & " AND Producto='COMISION CHEQUE DEVUELTO'"
''            Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''            If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Importe")) Then
''                Cheque_Devuelto = Rs_Adm_Facturas_Clientes.rdoColumns("Importe") + Rs_Adm_Facturas_Clientes.rdoColumns("Iva")
''                Puntos_Gestion = Puntos_Gestion + 1
''            Else
''                Cheque_Devuelto = 0
''            End If
''            Rs_Adm_Facturas_Clientes.Close
''            'Valida el nivel de gestión
''            Select Case Puntos_Gestion
''                Case 0
''                    Nivel_Gestion = "BAJO"
''                Case 1
''                    Nivel_Gestion = "BAJO"
''                Case 2
''                    Nivel_Gestion = "MEDIO"
''                Case 3
''                    Nivel_Gestion = "ALTO"
''            End Select
''            If Meses_Trasncurridos > 0 Then
''                Print #1, Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12); _
''                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12); _
''                    Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
''                    Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
''                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
''                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "#,###,##0.00"), 12); _
''                    Spc(10); Nivel_Gestion
''                Grid_Reporte.AddItem Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "###,##0.00"), 9) _
''                    & Chr(9) & Nivel_Gestion
''            Else
''                Print #1, Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
''                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
''                    Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
''                    Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
''                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
''                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
''                    Spc(10); Nivel_Gestion
''                Grid_Reporte.AddItem Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
''                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "###,##0.00"), 9) _
''                    & Chr(9) & Nivel_Gestion
''            End If
''            Rs_Cat_Clientes.MoveNext
''        Wend
''        Rs_Cat_Clientes.Close
''        Call Encabezado_Reportes("", "", "", False)
''    Else
''        MsgBox "Seleccione un cliente de la lista", vbExclamation
''        Exit Sub
''    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Reporte_Notas_Credito_Listado
'DESCRIPCIÓN: Muestra el reporte de las notas de credito emitidas como listado
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 8-Septiembre-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'PARÁMETROS        :
'*******************************************************************************
Private Sub Reporte_Notas_Credito_Listado()
Dim Rs_Consulta_Ope_Notas_Credito_Clientes As rdoResultset        'Consultas las notas de credito pertenecientes al cliente que se esta consultando
Dim Total_Notas_Credito As Double         'Es el importe total de las notas de credito
Dim Total_Notas_Credito_Canceladas As Double 'Es el importe total de las notas de credito canceladas

If Valida_Fechas(Dtp_Fecha_Inicio_Nota_Credito.Value, Dtp_Fecha_Fin_Nota_Credito.Value) = False Then
    MsgBox "Las fechas no son validas," & Chr(13) & Chr(13) & "favor de verificarlas", vbCritical
    Dtp_Fecha_Inicio_Nota_Credito.SetFocus
    Exit Sub
End If

If DateDiff("d", Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "yyyy/MM/dd"), Format(Dtp_Fecha_Fin_Nota_Credito.Value, "yyyy/MM/dd")) < 0 Then
    MsgBox "Las fechas no son validas," & Chr(13) & Chr(13) & "favor de verificarlas", vbCritical
    Dtp_Fecha_Inicio_Nota_Credito.SetFocus
    Exit Sub
End If

On Error GoTo handler
    MDIFrm_Apl_Principal.MousePointer = 11
    Grid_Reporte.Rows = 0
    Grid_Reporte.Cols = 6
    Call Encabezado_Reportes("REPORTE DE NOTAS DE CRÉDITO", Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Nota_Credito.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    Lineas_Encabezado = 6
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "Clave Cliente                                         No. Nota       Fecha Nota           Total"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Clave" & Chr(9) & "Cliente" & Chr(9) & "No. Nota" & Chr(9) & "Fecha Nota" & Chr(9) & "Total"
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Total_Notas_Credito = 0
    Total_Notas_Credito_Canceladas = 0
    
    'Consulta las facturas realizadas en la sucursal de acuerdo a las fechas proporcionadas por el usuario
    Mi_SQL = "SELECT No_Nota_Credito, Serie, Adm_Notas_Credito.Cliente_ID, Cancelada, Fecha, Total, Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito, Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito.Cliente_ID = Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Fecha >= " & Par_Fecha & Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "yyyy/MM/dd") & " 00:00:00" & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(Dtp_Fecha_Fin_Nota_Credito.Value, "yyyy/MM/dd") & " 11:59:59" & Par_Fecha
    If Cmb_Cliente_Nota_Credito.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito.Cliente_ID = '" & Format(Cmb_Cliente_Nota_Credito.ItemData(Cmb_Cliente_Nota_Credito.ListIndex), "00000") & "'"
    End If
    If Chk_Nota_Credito_Activa.Value = 0 And Chk_Nota_Credito_Cancelada.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cancelada = 'S'"
    ElseIf Chk_Nota_Credito_Activa.Value = 1 And Chk_Nota_Credito_Cancelada.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Serie, No_Nota_Credito, Nombre"
    Set Rs_Consulta_Ope_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Ope_Notas_Credito_Clientes
        If Not .EOF Then
            While Not .EOF
                'Imprime el listado
                If Trim(.rdoColumns("Cancelada")) = "N" Or Trim(.rdoColumns("Cancelada")) = "PC" Then
                    Print #1, Mid(.rdoColumns("Cliente_ID"), 1, 5); Spc(6 - Len(Mid(.rdoColumns("Cliente_ID"), 1, 5)));
                    Print #1, Mid(.rdoColumns("Nombre"), 1, 45); Spc(50 - Len(Mid(.rdoColumns("Nombre"), 1, 45)));
                    Print #1, Mid(Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))), 1, 10); Spc(13 - Len(Mid(Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))), 1, 10)));
                    Print #1, Format(.rdoColumns("Fecha"), "dd/MM/yyyy"); Spc(5);
'                    Print #1, Mid(.rdoColumns("Tipo_Nota_Credito"), 1, 10); Spc(13 - Len(Mid(.rdoColumns("Tipo_Nota_Credito"), 1, 10)));
                    Print #1, Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12)
                    Total_Notas_Credito = Total_Notas_Credito + Val(.rdoColumns("Total"))
                    Grid_Reporte.AddItem .rdoColumns("Cliente_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))) & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MM/yyyy") & Chr(9) & Format(.rdoColumns("Total"), "###,###,##0.00")
                Else
                    Print #1, Mid(.rdoColumns("Cliente_ID"), 1, 5); Spc(6 - Len(Mid(.rdoColumns("Cliente_ID"), 1, 5)));
                    Print #1, Mid("C A N C E L A D A", 1, 45); Spc(50 - Len(Mid("C A N C E L A D A", 1, 45)));
                    Print #1, Mid(Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))), 1, 10); Spc(13 - Len(Mid(Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))), 1, 10)));
                    Print #1, Format(.rdoColumns("Fecha"), "dd/MM/yyyy"); Spc(5);
'                    Print #1, Mid(.rdoColumns("Tipo_Nota_Credito"), 1, 10); Spc(13 - Len(Mid(.rdoColumns("Tipo_Nota_Credito"), 1, 10)));
                    Print #1, Conectar_Ayudante.Alinea_Derecha(Format("0", "#,###,##0.00"), 12)
'                    Total_Notas_Credito = Total_Notas_Credito + Val(.rdoColumns("Total"))
                    Grid_Reporte.AddItem .rdoColumns("Cliente_ID") & Chr(9) & "C A N C E L A D A" & Chr(9) & Trim(.rdoColumns("Serie")) & Val(Trim(.rdoColumns("No_Nota_Credito"))) & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MM/yyyy") & Chr(9) & Format("0", "###,###,##0.00")
                End If
                .MoveNext
            Wend
        End If
    End With
    Rs_Consulta_Ope_Notas_Credito_Clientes.Close
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1,
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem ""
    'Imprime los totales
    Print #1, "T O T A L                   "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito, "###,###,###.00"), 15)
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "TOTAL" & Chr(9) & Chr(9) & Total_Notas_Credito
    Print #1, "C A N C E L A D A S         "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito_Canceladas, "###,###,###.00"), 15)
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "CANCELADAS" & Chr(9) & Chr(9) & Total_Notas_Credito_Canceladas
    Call Encabezado_Reportes("", "", "", False)
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
    
handler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Close #1
    MsgBox Mi_SQL
    MsgBox Err.Description
End Sub

Public Sub Encabezado_Reportes(NombreReporte As String, Optional FechaInicio As String, _
Optional FechaFinal As String, Optional Abrir As Boolean)
Dim Carpeta_Reportes As String

    Carpeta_Reportes = App.Path & "\REP"
    If Abrir = True Then
        'Valida que exista la carpeta donde se almacenara el archivo
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Carpeta_Reportes, "CARPETA") = False Then
            MkDir Carpeta_Reportes & "\"
        End If
        Grid_Reporte.Rows = 0
        Open Carpeta_Reportes & "\Reportes.txt" For Output As #1
        Print #1, "************************************************************************************************************************"
        Print #1, Spc(2); Format(Date, "dd/MM/yyyy"); Spc(7); Conectar_Ayudante.Centrar_Texto(Nombre_Emisor, 70)
        Grid_Reporte.AddItem Format(Date, "dd/MM/yyyy") & Chr(9) & Nombre_Emisor
        Print #1, Spc(2); Format(Now, "HH:mm:ss") '; Spc(5); Empresa_Abreviatura; Spc(5); "Version "; App.Major; "." & App.Minor; "."; App.Revision '; Spc(10); Conectar_Ayudante.Centra("FARMACIA " & Space(10) & "Ciudad ", 30)
        Grid_Reporte.AddItem Format(Now, "dd/MMM/yyyy HH:mm:ss") ' & Chr(9) & Empresa_Abreviatura & Chr(9) & "Version " & App.Major & "." & App.Minor & "." & App.Revision '& Chr(9) & "FARMACIA " & Chr(9) & "Ciudad "
        Print #1,
        Grid_Reporte.AddItem ""
        Print #1, Spc(2); Conectar_Ayudante.Centrar_Texto(NombreReporte, Len("************************************************************************************************************************"))
        Grid_Reporte.AddItem NombreReporte
        If FechaInicio <> "" And FechaFinal <> "" Then
            Print #1, Conectar_Ayudante.Centrar_Texto("DEL " & Format(FechaInicio, "dd MMMM yyyy") & " AL " & Format(FechaFinal, "dd MMMM yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(FechaInicio, "dd MMMM yyyy") & " AL " & Format(FechaFinal, "dd MMMM yyyy")
        ElseIf FechaInicio <> "" Then
            Print #1, Conectar_Ayudante.Centrar_Texto("FECHA GENERACION " & Format(FechaInicio, "dd MMMM yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "FECHA GENERACION " & Format(FechaInicio, "dd MMMM yyyy")
            Grid_Reporte.AddItem ""
        End If
        Print #1, "************************************************************************************************************************"
        Print #1, Spc(2); ""
        Grid_Reporte.AddItem ""
    Else
        Print #1,
        Grid_Reporte.AddItem ""
        Close #1
        Pic_Reportes.Visible = False
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Printer.FontSize = 10
        Printer.FontSize = 8
        Me.Edm_Reporte.fileName = Carpeta_Reportes & "\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
''        Me.Width = 13600
''        Me.Edm_Reporte.Left = 0
''        Me.Edm_Reporte.Width = Me.Width - 100
''        If Screen.Height > 8340 Then
''            Me.Height = Screen.Height - 2500
''        Else
''            Me.Height = 8340
''        End If
''        Me.Edm_Reporte.Height = Me.Height - 1395
''        Btn_Imprimir.Top = Edm_Reporte.Height + 200
''        Btn_Exportar_Excel.Top = Btn_Imprimir.Top
''        Btn_Guardar_Disco.Top = Btn_Imprimir.Top
''        Btn_Regresar.Top = Btn_Imprimir.Top
''        Btn_Salir.Top = Btn_Imprimir.Top
''        Me.Left = 0
''        Me.Top = 0
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Reporte_Notas_Credito_Agrupado
'DESCRIPCIÓN: Muestra el reporte de las notas de credito emitidas con su detalle
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 8-Septiembre-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'PARÁMETROS        :
'*******************************************************************************
Private Sub Reporte_Notas_Credito_Agrupado()
Dim Total_Notas_Credito As Double         'Es el importe total de las notas de credito realizadas al cliente
Dim Total_Notas_Credito_Canceladas As Double 'Es el importe total de las notas de credito canceladas
Dim Total_Impuestos_Notas_Credito As Double           'Es el importe total de los impuestos
Dim Total_Impuestos_Canceladas As Double           'Es el importe total de los impuestos
Dim TTotal_Notas_Credito As Double         'Es el importe total de las notas de credito realizadas al cliente
Dim TTotal_Notas_Credito_Canceladas As Double 'Es el importe total de las notas de credito canceladas
Dim TTotal_Impuestos_Notas_Credito As Double           'Es el importe total de los impuestos
Dim TTotal_Impuestos_Canceladas As Double           'Es el importe total de los impuestos
Dim Rs_Consulta_Ope_Notas_Credito_Clientes As rdoResultset        'Consultas las notas de credito pertenecientes al cliente que se esta consultando
Dim Rs_Consulta_Detalles_Notas_Credito_Clientes As rdoResultset   'Consulta los detalles de las notas de credito que se le entrego al cliente
Dim RFC_Compara As String               'Para realizar la comparacion del ciclo
Dim Bandera_Monto As Boolean            'Para poner los totales por cliente

If Valida_Fechas(Dtp_Fecha_Inicio_Nota_Credito.Value, Dtp_Fecha_Fin_Nota_Credito.Value) = False Then
    MsgBox "Las fechas no son validas," & Chr(13) & Chr(13) & "favor de verificarlas", vbCritical
    Dtp_Fecha_Inicio_Nota_Credito.SetFocus
    Exit Sub
End If

If DateDiff("d", Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "yyyy/MM/dd"), Format(Dtp_Fecha_Fin_Nota_Credito.Value, "yyyy/MM/dd")) < 0 Then
    MsgBox "Las fechas no son validas," & Chr(13) & Chr(13) & "favor de verificarlas", vbCritical
    Dtp_Fecha_Inicio_Nota_Credito.SetFocus
    Exit Sub
End If

On Error GoTo handler
    MDIFrm_Apl_Principal.MousePointer = 11
    Grid_Reporte.Rows = 0
    Grid_Reporte.Cols = 7
    Call Encabezado_Reportes("REPORTE DE NOTAS DE CRÉDITO DE CLIENTES", Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Nota_Credito.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    Lineas_Encabezado = 9
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "        No. Nota Crédito    Fecha Emisión                                                Total   Estatus"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "No. Nota Crédito" & Chr(9) & "Fecha Emisión" & Chr(9) & "Total" & Chr(9) & "Estatus"
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Total_Notas_Credito = 0
    Total_Notas_Credito_Canceladas = 0
    Total_Impuestos_Notas_Credito = 0
    Total_Impuestos_Canceladas = 0
    TTotal_Notas_Credito = 0
    TTotal_Notas_Credito_Canceladas = 0
    TTotal_Impuestos_Notas_Credito = 0
    TTotal_Impuestos_Canceladas = 0
    'Consulta las notas de credito realizadas en la sucursal de acuerdo a las fechas proporcionadas por el usuario
    Mi_SQL = "SELECT Serie, No_Nota_Credito, Adm_Notas_Credito.Cliente_ID, Fecha_Creo_Xml, RFC, Nombre, Cancelada, Total"
    Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito, Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito.Cliente_ID = Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Fecha_Creo_Xml >= " & Par_Fecha & Format(Dtp_Fecha_Inicio_Nota_Credito.Value, "yyyy/MM/dd") & " 00:00:00" & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha_Creo_Xml <= " & Par_Fecha & Format(Dtp_Fecha_Fin_Nota_Credito.Value, "yyyy/MM/dd") & " 11:59:59" & Par_Fecha
    If Cmb_Cliente_Nota_Credito.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito.Cliente_ID = '" & Format(Cmb_Cliente_Nota_Credito.ItemData(Cmb_Cliente_Nota_Credito.ListIndex), "00000") & "'"
    End If
    If Chk_Nota_Credito_Activa.Value = 1 And Chk_Nota_Credito_Cancelada.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
    ElseIf Chk_Nota_Credito_Activa.Value = 0 And Chk_Nota_Credito_Cancelada.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cancelada = 'S'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY RFC, Serie, No_Nota_Credito"
    Set Rs_Consulta_Ope_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Ope_Notas_Credito_Clientes
        If Not .EOF Then
            Bandera_Monto = False
            While Not .EOF
                'Imprime los datos del cliente
                If RFC_Compara <> .rdoColumns("RFC") Then
                    If Bandera_Monto = True Then
                        Print #1, Spc(78); "------------------------------"
'                        Print #1, Spc(64); Conectar_Ayudante.Alinea_Derecha(Format(Total_Impuestos_Notas_Credito + Total_Impuestos_Canceladas, "###,###,##0.00"), 14);
                        Print #1, Spc(81); Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito + Total_Notas_Credito_Canceladas, "###,###,##0.00"), 14)
                        Print #1,
                        Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Total_Impuestos_Notas_Credito + Total_Impuestos_Canceladas, "###,###,##0.00") & Chr(9) & Format(Total_Notas_Credito + Total_Notas_Credito_Canceladas, "###,###,##0.00")
                        Grid_Reporte.AddItem ""
                    End If
                    'Asigna el nombre del cliente
                    
                    Print #1, "--------------------------------------------------------------------------------------------------------"
                    Print #1, "Cliente ---> " & .rdoColumns("RFC"); Spc(2); .rdoColumns("Nombre")
                    Print #1, "--------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem "Cliente ---> " & Chr(9) & .rdoColumns("RFC") & "  " & .rdoColumns("Nombre")
                    RFC_Compara = .rdoColumns("RFC")
                    Bandera_Monto = True
                    'Calcula los totales
                    TTotal_Notas_Credito = TTotal_Notas_Credito + Total_Notas_Credito
                    TTotal_Notas_Credito_Canceladas = TTotal_Notas_Credito_Canceladas + Total_Notas_Credito_Canceladas
'                    TTotal_Impuestos_Notas_Credito = TTotal_Impuestos_Notas_Credito + Total_Impuestos_Notas_Credito
'                    TTotal_Impuestos_Canceladas = TTotal_Impuestos_Canceladas + Total_Impuestos_Canceladas
                    'Reinica los totales
                    Total_Notas_Credito = 0
                    Total_Notas_Credito_Canceladas = 0
                    Total_Impuestos_Notas_Credito = 0
                    Total_Impuestos_Canceladas = 0
                End If
                'Imprime los detalles
                If Trim(.rdoColumns("Cancelada")) = "N" Then
                    Print #1, Spc(11); Conectar_Ayudante.Alinea_Derecha(Trim(.rdoColumns("Serie")) & " " & Val(Trim(.rdoColumns("No_Nota_Credito"))), 6); Spc(3); Format(.rdoColumns("Fecha_Creo_Xml"), "dd/MMM/yyyy HH:mm:ss"); Spc(41); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "###,###,##0.00"), 14); Spc(3); "ACTIVA"
                    Total_Notas_Credito = Total_Notas_Credito + Val(.rdoColumns("Total"))
'                    Total_Impuestos_Notas_Credito = Total_Impuestos_Notas_Credito + Val(.rdoColumns("Impuestos"))
                    Grid_Reporte.AddItem Trim(.rdoColumns("Serie")) & " " & Val(Trim(.rdoColumns("No_Nota_Credito"))) & Chr(9) & Format(.rdoColumns("Fecha_Creo_Xml"), "dd/MMM/yyyy HH:mm:ss") & Chr(9) & Format(.rdoColumns("Total"), "###,###,##0.00") & Chr(9) & "ACTIVA"
                Else
                    Print #1, Spc(11); Conectar_Ayudante.Alinea_Derecha(Trim(.rdoColumns("Serie")) & " " & Val(Trim(.rdoColumns("No_Nota_Credito"))), 6); Spc(3); Format(.rdoColumns("Fecha_Creo_Xml"), "dd/MMM/yyyy HH:mm:ss"); Spc(41); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "###,###,##0.00"), 14); Spc(3); "CANCELADA"
                    Total_Notas_Credito_Canceladas = Total_Notas_Credito_Canceladas + Val(.rdoColumns("Total"))
'                    Total_Impuestos_Canceladas = Total_Impuestos_Canceladas + Val(.rdoColumns("Impuestos"))
                    Grid_Reporte.AddItem Trim(.rdoColumns("Serie")) & " " & Val(Trim(.rdoColumns("No_Nota_Credito"))) & Chr(9) & Format(.rdoColumns("Fecha_Creo_Xml"), "dd/MMM/yyyy HH:mm:ss") & Chr(9) & Format(.rdoColumns("Total"), "###,###,##0.00") & Chr(9) & "CANCELADA"
                End If
                .MoveNext
            Wend
        End If
    End With
    Rs_Consulta_Ope_Notas_Credito_Clientes.Close
    Print #1, Spc(78); "------------------------------"
'    Print #1, Spc(64); Conectar_Ayudante.Alinea_Derecha(Format(Total_Impuestos_Notas_Credito + Total_Impuestos_Canceladas, "###,###,##0.00"), 14);
    Print #1, Spc(81); Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito + Total_Notas_Credito_Canceladas, "###,###,##0.00"), 14)
    Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Total_Notas_Credito + Total_Notas_Credito_Canceladas, "###,###,##0.00")
    Grid_Reporte.AddItem ""
    Print #1, Spc(7); "-----------------------------------------------------------------------------------------------------------------"
    Print #1,
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem ""
    'Imprime los totales
    Print #1, Spc(46); "T O T A L  N O T A S  C R E D I T O"; Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito, "###,###,###.00"), 14)
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "SUBTOTAL" & Chr(9) & Chr(9) & Total_Notas_Credito
'    Print #1, Spc(44); "I M P U E S T O S                  "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Impuestos_Notas_Credito, "###,###,###.00"), 15)
'    Grid_Reporte.AddItem Chr(9) & Chr(9) & "IMPUESTOS" & Chr(9) & Chr(9) & Total_Impuestos_Notas_Credito
    Print #1,
    Print #1, Spc(46); "C A N C E L A D A S                "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Notas_Credito_Canceladas, "###,###,###.00"), 14)
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "CANCELADAS" & Chr(9) & Chr(9) & Total_Notas_Credito_Canceladas
'    Print #1, Spc(44); "I M P U E S T O S                  "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Impuestos_Canceladas, "###,###,###.00"), 15)
'    Grid_Reporte.AddItem Chr(9) & Chr(9) & "IMPUESTOS" & Chr(9) & Chr(9) & Total_Impuestos_Canceladas
    Print #1,
    Call Encabezado_Reportes("", "", "", False)
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
    
handler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Close #1
    MsgBox Mi_SQL
    MsgBox Err.Description
End Sub

