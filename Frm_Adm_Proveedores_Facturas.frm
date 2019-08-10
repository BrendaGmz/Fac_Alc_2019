VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Proveedores_Facturas 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   7260
   Begin VB.PictureBox Pic_Facturas_Proveedores 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6630
      Left            =   -15
      ScaleHeight     =   6630
      ScaleWidth      =   7230
      TabIndex        =   0
      Top             =   -15
      Width           =   7230
      Begin VB.CommandButton Btn_Pagos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar Pagos"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5610
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Frame Fra_Anticipos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anticipos"
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
         Height          =   1305
         Left            =   3495
         TabIndex        =   51
         Top             =   2460
         Width           =   3615
         Begin VB.ListBox Lst_Anticipos 
            Height          =   960
            Left            =   75
            Style           =   1  'Checkbox
            TabIndex        =   52
            Top             =   255
            Width           =   3450
         End
      End
      Begin VB.Frame Fra_Cantidades 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cantidades"
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
         Height          =   2715
         Left            =   105
         TabIndex        =   26
         Top             =   3750
         Width           =   3300
         Begin VB.TextBox Txt_Subtotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            TabIndex        =   89
            Top             =   157
            Width           =   1770
         End
         Begin VB.TextBox Txt_Impuesto_Cedular 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1590
            Width           =   1725
         End
         Begin VB.TextBox Txt_Flete 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   32
            Top             =   525
            Width           =   1770
         End
         Begin VB.TextBox Txt_Retencion_ISR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1230
            Width           =   645
         End
         Begin VB.TextBox Txt_Retencion_IVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1230
            Width           =   615
         End
         Begin VB.TextBox Txt_Saldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2310
            Width           =   1725
         End
         Begin VB.TextBox Txt_Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1950
            Width           =   1725
         End
         Begin VB.TextBox Txt_IVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   870
            Width           =   1725
         End
         Begin VB.Label Lbl_Fetes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Flete"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   585
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ISR"
            Height          =   195
            Left            =   2160
            TabIndex        =   40
            Top             =   1290
            Width           =   270
         End
         Begin VB.Label Lbl_Impuesto_Cedular 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Impuesto Cedular"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1650
            Width           =   1230
         End
         Begin VB.Label Lbl_Retencion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Retencion     IVA"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1290
            Width           =   1215
         End
         Begin VB.Label Lbl_Subtotal 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "SubTotal"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   2370
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   35
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I.V.A."
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   34
            Top             =   930
            Width           =   390
         End
      End
      Begin VB.Frame Fra_Nota_Credito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nota de Credito"
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
         Height          =   1305
         Left            =   105
         TabIndex        =   19
         Top             =   2460
         Width           =   3300
         Begin VB.TextBox Txt_Anticipo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   915
            Width           =   1710
         End
         Begin VB.TextBox Txt_Monto_Nota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1500
            MaxLength       =   20
            TabIndex        =   21
            Top             =   570
            Width           =   1710
         End
         Begin VB.TextBox Txt_Nota_Credito 
            Height          =   315
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   20
            Top             =   225
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Anticipo"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   25
            Top             =   975
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   24
            Top             =   630
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Credito"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   23
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.CommandButton Btn_Imprime_Contra_Recibo 
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
         Height          =   315
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3975
         Visible         =   0   'False
         Width           =   915
      End
      Begin TabDlg.SSTab Tab_Facturas 
         Height          =   2340
         Left            =   105
         TabIndex        =   1
         Top             =   90
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   4128
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos Factura"
         TabPicture(0)   =   "Frm_Adm_Proveedores_Facturas.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Datos_Factura"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Compras"
         TabPicture(1)   =   "Frm_Adm_Proveedores_Facturas.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Compras"
         Tab(1).ControlCount=   1
         Begin VB.Frame Fra_Compras 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   1875
            Left            =   -74910
            TabIndex        =   82
            Top             =   375
            Width           =   6825
            Begin VB.CheckBox Chk_Seleccionar 
               BackColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   180
               TabIndex        =   83
               Top             =   375
               Visible         =   0   'False
               Width           =   210
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Compras 
               Height          =   1590
               Left            =   75
               TabIndex        =   84
               Top             =   150
               Width           =   6660
               _ExtentX        =   11748
               _ExtentY        =   2805
               _Version        =   393216
               Rows            =   0
               Cols            =   6
               FixedRows       =   0
               BackColorBkg    =   16777215
               AllowUserResizing=   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Fra_Datos_Factura 
            BackColor       =   &H00FFFFFF&
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
            Height          =   1920
            Left            =   75
            TabIndex        =   2
            Top             =   330
            Width           =   6825
            Begin VB.ComboBox Cmb_Proveedor 
               Height          =   315
               Left            =   1470
               TabIndex        =   88
               Top             =   315
               Width           =   5250
            End
            Begin VB.ComboBox Cmb_Estatus 
               Height          =   315
               ItemData        =   "Frm_Adm_Proveedores_Facturas.frx":0038
               Left            =   4410
               List            =   "Frm_Adm_Proveedores_Facturas.frx":0048
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   705
               Width           =   2310
            End
            Begin VB.TextBox Txt_Tipo_Cambio 
               Height          =   315
               Left            =   6180
               MaxLength       =   8
               TabIndex        =   6
               Text            =   "1"
               Top             =   1440
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.ComboBox Cmb_Tipo 
               Height          =   315
               ItemData        =   "Frm_Adm_Proveedores_Facturas.frx":0073
               Left            =   4410
               List            =   "Frm_Adm_Proveedores_Facturas.frx":0089
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1065
               Width           =   2310
            End
            Begin VB.ComboBox Cmb_Moneda 
               Height          =   315
               ItemData        =   "Frm_Adm_Proveedores_Facturas.frx":00CC
               Left            =   4395
               List            =   "Frm_Adm_Proveedores_Facturas.frx":00D6
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox Txt_No_Factura 
               Height          =   315
               Left            =   1470
               MaxLength       =   15
               TabIndex        =   3
               Top             =   705
               Width           =   1920
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Factura 
               Height          =   315
               Left            =   1470
               TabIndex        =   7
               Top             =   1065
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMMM yyyy"
               Format          =   58458115
               CurrentDate     =   38038
            End
            Begin MSComCtl2.DTPicker DTP_Fecha_Recepcion 
               Height          =   315
               Left            =   1470
               TabIndex        =   8
               Top             =   1440
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMMM yyyy"
               Format          =   58458115
               CurrentDate     =   38038
            End
            Begin VB.Label Lbl_Tipo_Cambio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cambio"
               Height          =   195
               Left            =   5610
               TabIndex        =   16
               Top             =   1500
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label Lbl_Estatus 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Estatus"
               Height          =   195
               Left            =   3405
               TabIndex        =   15
               Top             =   765
               Width           =   525
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   10
               Left            =   3480
               TabIndex        =   14
               Top             =   1500
               Width           =   585
            End
            Begin VB.Label Lbl_No_Factura 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "No. Factura"
               Height          =   195
               Left            =   135
               TabIndex        =   13
               Top             =   765
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Factura"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   12
               Top             =   1125
               Width           =   1035
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Factura"
               Height          =   195
               Index           =   13
               Left            =   3405
               TabIndex        =   11
               Top             =   1125
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Recepcion"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   10
               Top             =   1500
               Width           =   1275
            End
            Begin VB.Label Lbl_Proveedor 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Proveedor"
               Height          =   195
               Left            =   135
               TabIndex        =   9
               Top             =   375
               Width           =   735
            End
         End
      End
      Begin VB.Frame Fra_Condiciones_Pago 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Condiciones de Pago"
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
         Height          =   2715
         Left            =   3495
         TabIndex        =   42
         Top             =   3750
         Width           =   3615
         Begin VB.TextBox Txt_Comentarios 
            Height          =   1290
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   1335
            Width           =   2190
         End
         Begin VB.ComboBox Cmb_Forma_Pago 
            Height          =   315
            ItemData        =   "Frm_Adm_Proveedores_Facturas.frx":00EA
            Left            =   1320
            List            =   "Frm_Adm_Proveedores_Facturas.frx":00F7
            TabIndex        =   43
            Top             =   975
            Width           =   2190
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Pago 
            Height          =   315
            Left            =   1320
            TabIndex        =   46
            Top             =   615
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58458115
            CurrentDate     =   38038
         End
         Begin VB.TextBox Txt_Contra_Recibo 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   225
            Width           =   2160
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contra Recibo"
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   50
            Top             =   285
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   49
            Top             =   675
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   48
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   47
            Top             =   1395
            Width           =   870
         End
      End
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
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":011C
      Style           =   1  'Graphical
      TabIndex        =   57
      Tag             =   "A"
      Top             =   6585
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
      Left            =   1455
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":3653
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "M"
      Top             =   6585
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
      Left            =   2850
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":6D84
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "C"
      Top             =   6585
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
      Left            =   4245
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":A310
      Style           =   1  'Graphical
      TabIndex        =   54
      Tag             =   "B"
      Top             =   6585
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
      Left            =   5655
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":D8CA
      Style           =   1  'Graphical
      TabIndex        =   53
      Tag             =   "A"
      Top             =   6585
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Regresar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Regresar"
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
      Left            =   5655
      Picture         =   "Frm_Adm_Proveedores_Facturas.frx":10FC9
      Style           =   1  'Graphical
      TabIndex        =   85
      Tag             =   "A"
      Top             =   6585
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox Pic_Busqueda_Facturas_Proveedores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6510
      Left            =   135
      ScaleHeight     =   6480
      ScaleWidth      =   6930
      TabIndex        =   58
      Top             =   15
      Visible         =   0   'False
      Width           =   6960
      Begin VB.Frame Fra_Busqueda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busqueda de Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6420
         Left            =   90
         TabIndex        =   59
         Top             =   45
         Width           =   6795
         Begin VB.CommandButton Btn_Busqueda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Busqueda"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4545
            Style           =   1  'Graphical
            TabIndex        =   86
            Tag             =   "C"
            Top             =   1650
            Width           =   2115
         End
         Begin VB.CheckBox Chk_Busqueda_Fecha_Factura 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Por Fecha Factura"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   975
            Width           =   1830
         End
         Begin VB.ComboBox Cmb_Con_Estatus 
            Height          =   315
            ItemData        =   "Frm_Adm_Proveedores_Facturas.frx":14529
            Left            =   2100
            List            =   "Frm_Adm_Proveedores_Facturas.frx":14533
            TabIndex        =   72
            Top             =   1620
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.CheckBox Chk_Busqueda_Estatus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Por Estatus"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CommandButton Btn_Excel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Excel"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   375
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   2100
            Width           =   840
         End
         Begin VB.TextBox Txt_Saldos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   2100
            Width           =   1215
         End
         Begin VB.TextBox Txt_Abonos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   2100
            Width           =   1215
         End
         Begin VB.TextBox Txt_Totales 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2100
            Width           =   1215
         End
         Begin VB.CommandButton Btn_Mas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton Btn_Seleccionar 
            Caption         =   "Ver Detalles"
            Height          =   315
            Left            =   5475
            TabIndex        =   65
            Top             =   7560
            Width           =   1215
         End
         Begin VB.ComboBox Cmb_Con_Proveedor 
            Height          =   315
            Left            =   2100
            TabIndex        =   64
            Top             =   555
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.TextBox Txt_Con_No_Factura 
            Height          =   285
            Left            =   2100
            MaxLength       =   10
            TabIndex        =   63
            Top             =   225
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.CheckBox Chk_Busqueda_Fecha_Recepcion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Por Fecha Recepcion"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   1350
            Width           =   1875
         End
         Begin VB.CheckBox Chk_Busqueda_Proveedor 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Por Proveedor"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   615
            Width           =   1755
         End
         Begin VB.CheckBox Chk_Busqueda_No_Factura 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Por No. de Factura"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   270
            Width           =   1710
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Factura_Inicial 
            Height          =   315
            Left            =   2100
            TabIndex        =   74
            Top             =   915
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58458115
            CurrentDate     =   38038
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Recepcion_Inicial 
            Height          =   315
            Left            =   2100
            TabIndex        =   75
            Top             =   1290
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58458115
            CurrentDate     =   38038
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Recepcion_Final 
            Height          =   315
            Left            =   4560
            TabIndex        =   76
            Top             =   1290
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58458115
            CurrentDate     =   38038
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Factura_Final 
            Height          =   315
            Left            =   4560
            TabIndex        =   77
            Top             =   915
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58458115
            CurrentDate     =   38038
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Facturas 
            Height          =   3885
            Left            =   75
            TabIndex        =   78
            Top             =   2445
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   6853
            _Version        =   393216
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
            BackColorBkg    =   16777215
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Saldos"
            Height          =   195
            Left            =   4800
            TabIndex        =   81
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Abonos"
            Height          =   195
            Left            =   2925
            TabIndex        =   80
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total"
            Height          =   195
            Left            =   1245
            TabIndex        =   79
            Top             =   2160
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "Frm_Adm_Proveedores_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Buscar_Click()
    Pic_Busqueda_Facturas_Proveedores.Visible = True
    Pic_Facturas_Proveedores.Visible = False
    Btn_Pagos.Visible = False
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Con_Proveedor, 1, "Nombre")
End Sub

Private Sub Btn_Busqueda_Click()
    Call Consulta_Facturas
End Sub

Private Sub Btn_Eliminar_Click()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset        'Manejo de Registro
Dim RS_Consulta_Vales As rdoResultset                       'Manejo de Registro
Dim Rs_Consulta_Movimientos As rdoResultset                 'Manejo de Registro
Dim Rs_Consulta_Anticipos_Proveedores As rdoResultset       'Manejo de Registro
Dim RS_Consulta_Notas_Credito As rdoResultset               'Manejo de Registro
Dim Rs_Actualiza_Presupuesto As rdoResultset
Dim Respuesta As Integer                                    'Almacena el valor de la respuesta
Dim Fecha As String, Cuenta As String
Dim Tipo_Presupuesto As String

On Error GoTo Handler
    If Btn_Eliminar.Caption = "Eliminar" Then
        If MsgBox("¿Esta seguro de eliminar la factura?", vbYesNo + vbCritical) = vbYes Then
            Conexion_Base.BeginTrans
            'Consulta la factura a Eliminar
            Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
            Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
            Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
            With Rs_Consulta_Facturas_Proveedores
                If Not Rs_Consulta_Facturas_Proveedores.EOF Then
                    .Delete
                End If
            End With
            Rs_Consulta_Facturas_Proveedores.Close
            Conexion_Base.CommitTrans
            MsgBox "Factura eliminada exitosamente", vbInformation
            Call Conectar_Ayudante.Limpiar_Textos(Me)
            Btn_Modificar.Enabled = True
            Btn_Eliminar.Caption = "Eliminar"
       End If
    Else
        If MsgBox("¿Está seguro de cancelar la factura?", vbYesNo + vbQuestion) = vbYes Then
            Conexion_Base.BeginTrans
            'Cambia el Status de la Factura a Cancelada
            Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
            Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "' "
            Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Consulta_Facturas_Proveedores.EOF Then
                With Rs_Consulta_Facturas_Proveedores
                    .Edit
                        .rdoColumns("Cancelada") = "S"
                        .rdoColumns("Pagada") = "N"
                        .rdoColumns("Abono") = 0
                        '.rdoColumns("Saldo") = .rdoColumns("Total")
                        .rdoColumns("Saldo") = 0
                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                        .rdoColumns("Fecha_Modifico") = Now
                        .rdoColumns("Aplicado_Presupuesto") = "NO"
                    .Update
                End With
            End If
            Rs_Consulta_Facturas_Proveedores.Close
            'Cancela los movimientos generados
            Mi_SQL = "SELECT * FROM Adm_Movimientos"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
            Mi_SQL = Mi_SQL & " AND Proveedor_Cliente='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
            Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Movimientos.EOF
                Fecha = Rs_Consulta_Movimientos.rdoColumns("Fecha")
                Cuenta = Rs_Consulta_Movimientos.rdoColumns("Cuenta")
                Call Conectar_Ayudante.Actualiza_Saldo(Fecha, Cuenta)
                'Elimina el anticipo generado por el movimiento
                Mi_SQL = "SELECT Aplicado,Usuario_Modifico,Fecha_Modifico FROM Adm_Anticipos_Proveedores"
                Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Rs_Consulta_Movimientos.rdoColumns("No_Movimiento") & "'"
                Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
                If Not Rs_Consulta_Anticipos_Proveedores.EOF Then
                    Rs_Consulta_Anticipos_Proveedores.Delete
                End If
                Rs_Consulta_Anticipos_Proveedores.Close
                Rs_Consulta_Movimientos.MoveNext
            Wend
            Rs_Consulta_Movimientos.Close
            Mi_SQL = "UPDATE Adm_Movimientos"
            Mi_SQL = Mi_SQL & " SET Estatus='C',Concepto='CANCELADO  '+ Concepto,Cantidad=0"
            Mi_SQL = Mi_SQL & " ,Usuario_Modifico='" & Nombre_Usuario & "',Fecha_Modifico='" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "'"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
            Mi_SQL = Mi_SQL & " AND Proveedor_Cliente='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
            Conexion_Base.Execute Mi_SQL
            'Cambia el estatus del anticipo
'            Mi_SQL = "SELECT Aplicado,Usuario_Modifico,Fecha_Modifico FROM Adm_Anticipos_Proveedores"
'            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
'            Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
'            Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
'            With Rs_Consulta_Anticipos_Proveedores
'                If Not Rs_Consulta_Anticipos_Proveedores.EOF Then
'                    .Edit
'                        .rdoColumns("Aplicado") = "N"
'                        .rdoColumns("Usuario_Modifico") = Usuario
'                        .rdoColumns("Fecha_Modifico") = Now
'                    .Update
'                End If
'            End With
'            Rs_Consulta_Anticipos_Proveedores.Close
            'Elimina la nota de credito
            Mi_SQL = "SELECT * FROM Adm_Proveedores_Notas_Credito"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)) & "'"
            Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
            Set RS_Consulta_Notas_Credito = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
            With RS_Consulta_Notas_Credito
                If Not RS_Consulta_Notas_Credito.EOF Then
                    .Delete
                End If
            End With
            RS_Consulta_Notas_Credito.Close
            Conexion_Base.CommitTrans
            MsgBox "Factura cancelada exitosamente", vbInformation
            Call Conectar_Ayudante.Limpiar_Textos(Me)
            Btn_Modificar.Enabled = True
            Btn_Eliminar.Caption = "Eliminar"
        End If
    End If
    Exit Sub
Handler:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Excel_Click()
Dim RutaArchivo As String

    ' Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.Xls|"
    ' Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    ' Display name of selected file

    RutaArchivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    
    Open RutaArchivo For Output As #1
        For I = 0 To Grid_Consulta_Facturas.Rows - 1
            For J = 0 To Grid_Consulta_Facturas.Cols - 1
                CABECERA = CABECERA & Grid_Consulta_Facturas.TextMatrix(I, J) & Chr(9)
            Next J
            Print #1, CABECERA
            CABECERA = ""
        Next I
    Close #1
    MsgBox "Reporte Importado"
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Btn_Imprime_Contra_Recibo_Click()
Dim I As Integer

On Error GoTo Handler

    If MsgBox("¿Desea reimprimir el contrarrecibo?", vbYesNo + vbQuestion) = vbYes Then
        Contrarecibo = Val(InputBox("Cuantas copias desea imprimir del Contra recibo", "Impresión de contrarrecibo"))
        If IsNumeric(Contrarecibo) Then
            If Contrarecibo > 0 And Contrarecibo <= 10 Then
                For I = 1 To Contrarecibo
                    Imprime_Contra_Recibo
                Next I
                MsgBox "Contra recibo enviado a impresión", vbInformation
            End If
        End If
    End If
    
Handler:
    If Err.Number = "482" Then
        MsgBox "Operación Cancelada"
    End If
    Exit Sub
End Sub

Private Sub Btn_Mas_Click()
    If Btn_Mas.Caption = "+" Then
        Me.Top = 0
        Me.Left = 0
        Me.Width = MDIFrm_Apl_Principal.ScaleWidth
        Me.Height = MDIFrm_Apl_Principal.ScaleHeight
        Grid_Consulta_Facturas.Width = MDIFrm_Apl_Principal.ScaleWidth - 1000
        Grid_Consulta_Facturas.Height = MDIFrm_Apl_Principal.ScaleHeight - 5000
        Fra_Busqueda.Width = MDIFrm_Apl_Principal.ScaleWidth - 400
        Fra_Busqueda.Height = MDIFrm_Apl_Principal.ScaleHeight - 1000
        Pic_Busqueda_Facturas_Proveedores.Height = MDIFrm_Apl_Principal.ScaleHeight - 1000
        Pic_Busqueda_Facturas_Proveedores.Width = MDIFrm_Apl_Principal.ScaleWidth - 400
        Btn_Buscar.Visible = False
        Btn_Eliminar.Visible = False
        Btn_Modificar.Visible = False
        Btn_Nuevo.Visible = False
        Btn_Salir.Visible = False
        Btn_Seleccionar.Visible = False
        Btn_Mas.Caption = "-"
    Else
        Me.Left = 1000
        Me.Height = 9015
        Me.Width = 7215
        Grid_Consulta_Facturas.Width = 6630
        Grid_Consulta_Facturas.Height = 3900
        Fra_Busqueda.Width = 6795
        Fra_Busqueda.Height = 7185
        Pic_Busqueda_Facturas_Proveedores.Height = 7440
        Pic_Busqueda_Facturas_Proveedores.Width = 6960
        Btn_Buscar.Visible = True
        Btn_Eliminar.Visible = True
        Btn_Modificar.Visible = True
        Btn_Nuevo.Visible = True
        Btn_Salir.Visible = True
        Btn_Mas.Caption = "+"
    End If
End Sub

Private Sub Btn_Modificar_Click()
Dim Rs_Modifica_Facturas_Proveedores As rdoResultset
Dim Rs_Modifica_Anticipos_Proveedores As rdoResultset
Dim Rs_Modifica_Movimientos As rdoResultset
Dim Rs_Agrega_Notas_Credito_Proveedores As rdoResultset
Dim Rs_Consulta_Tipos_Pagos As rdoResultset
Dim Rs_Actualiza_Presupuesto As rdoResultset
Dim Rs_Actualiza_Compras As rdoResultset
Dim Respuesta As Integer
Dim Gasto_Anterior As String
Dim I As Integer
Dim Total_Factura As Double

On Error GoTo Handler
    If Btn_Modificar.Caption = "Modificar" Then
        If Trim(Txt_No_Factura.Text) = "" Then Exit Sub
        Pic_Busqueda_Facturas_Proveedores.Visible = False
        Pic_Facturas_Proveedores.Visible = True
        Fra_Datos_Factura.Enabled = True
        Fra_Nota_Credito.Enabled = True
        Fra_Anticipos.Enabled = True
        Fra_Cantidades.Enabled = True
        Fra_Condiciones_Pago.Enabled = True
        Fra_Compras.Enabled = True
        Btn_Imprime_Contra_Recibo.Visible = False
        Btn_Nuevo.Enabled = False
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Salir.Visible = False
        Btn_Regresar.Visible = True
    Else
        If Cmb_Proveedor.ListIndex > -1 And Txt_No_Factura.Text <> "" Then
            If MsgBox("¿Esta seguro de modificar la factura?", vbYesNo + vbCritical) = vbYes Then
                Conexion_Base.BeginTrans
                'Consulta la factura para actualizar su información
                Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
                Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Txt_No_Factura & "'"
                Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
                Set Rs_Modifica_Facturas_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Not Rs_Modifica_Facturas_Proveedores.EOF Then
                    With Rs_Modifica_Facturas_Proveedores
                        .Edit
                            .rdoColumns("Tipo") = Cmb_Tipo.Text
                            .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                            .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                            .rdoColumns("Fecha_Recepcion") = Format(DTP_Fecha_Recepcion.Value, "MM/dd/yyyy")
                            .rdoColumns("Moneda") = Cmb_Moneda.Text
                            .rdoColumns("Tipo_Cambio") = Val(Txt_Tipo_Cambio.Text)
                            .rdoColumns("Cantidad") = 1
                            .rdoColumns("Precio_Unitario") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
                            .rdoColumns("Flete") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Flete.Text, ","))
                            .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
                            .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))
                            .rdoColumns("Retencion_IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Retencion_IVA.Text, ","))
                            .rdoColumns("Retencion_ISR") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Retencion_ISR.Text, ","))
                            If Cmb_Tipo.Text = "Honorarios" Then
                                .rdoColumns("Impuesto_Cedular") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Impuesto_Cedular.Text, ","))
                                .rdoColumns("Retencion_Fletes") = 0
                            Else
                                .rdoColumns("Retencion_Fletes") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Impuesto_Cedular.Text, ","))
                                .rdoColumns("Impuesto_Cedular") = 0
                            End If
                            .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
                            .rdoColumns("Abono") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ","))
                            .rdoColumns("Saldo") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Saldo.Text, ","))
                            If Val(Txt_Saldo.Text) = 0 Or Cmb_Estatus.Text = "Pagada" Then
                                .rdoColumns("Pagada") = "S"
                            Else
                                .rdoColumns("Pagada") = "N"
                            End If
                            If Cmb_Estatus.Text = "Cancelada" Then
                                .rdoColumns("Cancelada") = "S"
                            Else
                                .rdoColumns("Cancelada") = "N"
                            End If
                            .rdoColumns("Fecha_Pago") = Format(DTP_Fecha_Pago.Value, "MM/dd/yyyy")
                            .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.Text
                            .rdoColumns("Comentarios") = Txt_Comentarios.Text
                            .rdoColumns("Contra_Recibo") = Txt_Contra_Recibo.Text
                            '.rdoColumns("No_Anticipo") = Val(Mid(Cmb_Anticipos, 1, 6))
                            'Datos presupuestales
                            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                            .rdoColumns("Fecha_Modifico") = Now
                        .Update
                    End With
                    'Les quita la relación a las compras para agregar las nuevas
                    Mi_SQL = "UPDATE Tmp_Proveedores_Facturas"
                    Mi_SQL = Mi_SQL & " SET No_Factura_Proveedor=NULL"
                    Mi_SQL = Mi_SQL & " ,Aplicada='NO'"
                    Mi_SQL = Mi_SQL & " ,Usuario_Aplico=NULL"
                    Mi_SQL = Mi_SQL & " WHERE No_Factura_Proveedor='" & Txt_No_Factura & "'"
                    Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
                    Conexion_Base.Execute Mi_SQL
                    'Actualiza las compras asignándoles la factura capturada
                    For I = 1 To Grid_Compras.Rows - 1
                        If Trim(Grid_Compras.TextMatrix(I, 1)) = "SI" Then
                            Mi_SQL = "SELECT No_Control,No_Factura_Proveedor,Aplicada,Usuario_Aplico"
                            Mi_SQL = Mi_SQL & " FROM Tmp_Proveedores_Facturas"
                            Mi_SQL = Mi_SQL & " WHERE No_Control='" & Trim(Grid_Compras.TextMatrix(I, 0)) & "'"
                            Set Rs_Actualiza_Compras = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            If Not Rs_Actualiza_Compras.EOF Then
                                Rs_Actualiza_Compras.Edit
                                    Rs_Actualiza_Compras.rdoColumns("No_Factura_Proveedor") = UCase(Trim(Txt_No_Factura.Text))
                                    Rs_Actualiza_Compras.rdoColumns("Aplicada") = "SI"
                                    Rs_Actualiza_Compras.rdoColumns("Usuario_Aplico") = Nombre_Usuario
                                Rs_Actualiza_Compras.Update
                            End If
                            Rs_Actualiza_Compras.Close
                        End If
                    Next I
                    'Consulta los anticipos que se aplicarán
                    For I = 0 To Lst_Anticipos.ListCount - 1
                        If Lst_Anticipos.Selected(I) = True Then
                            'Consulta el anticipo si esta aplicado
                            Mi_SQL = "SELECT * FROM Adm_Anticipos_Proveedores"
                            Mi_SQL = Mi_SQL & " WHERE No_Anticipo=" & Val(Mid(Lst_Anticipos.List(I), 1, 7))
                            Set Rs_Modifica_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            With Rs_Modifica_Anticipos_Proveedores
                                If Not .EOF Then
                                    .Edit
                                        .rdoColumns("Aplicado") = "S"
                                        .rdoColumns("No_Factura") = Txt_No_Factura.Text
                                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                        .rdoColumns("Fecha_Modifico") = Now
                                    .Update
                                End If
                            End With
                            'Pone el número de factura en el movimiento administrativo
                            Mi_SQL = "SELECT * FROM Adm_Movimientos "
                            Mi_SQL = Mi_SQL & " WHERE No_Movimiento=" & Rs_Modifica_Anticipos_Proveedores.rdoColumns("No_Movimiento")
                            Set Rs_Modifica_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            With Rs_Modifica_Movimientos
                                If Not .EOF Then
                                    .Edit
                                        .rdoColumns("No_Factura") = Txt_No_Factura.Text
                                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                        .rdoColumns("Fecha_Modifico") = Now
                                    .Update
                                End If
                            End With
                            Rs_Modifica_Anticipos_Proveedores.Close
                            Rs_Modifica_Movimientos.Close
                        End If
                    Next I
                    'Captura la nota de crédito
                    If Txt_Nota_Credito.Text <> "" And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ",")) > 0 Then
                        Set Rs_Agrega_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Agregar("Adm_Notas_Credito_Proveedores")
                        With Rs_Agrega_Notas_Credito_Proveedores
                            .AddNew
                                .rdoColumns("No_Nota_Credito") = Txt_Nota_Credito.Text
                                .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                                .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                                .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ","))
                                .rdoColumns("No_Factura") = Txt_No_Factura.Text
                                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                                .rdoColumns("Fecha_Creo") = Now
                            .Update
                        End With
                        Rs_Agrega_Notas_Credito_Proveedores.Close
                    End If
                    Btn_Nuevo.Enabled = True
                    Btn_Modificar.Caption = "Modificar"
                    Btn_Eliminar.Enabled = True
                    Btn_Buscar.Enabled = True
                    Btn_Salir.Caption = "Salir"
                    Fra_Datos_Factura.Enabled = False
                    Fra_Nota_Credito.Enabled = False
                    Fra_Anticipos.Enabled = False
                    Fra_Cantidades.Enabled = False
                    Fra_Condiciones_Pago.Enabled = False
                    Fra_Compras.Enabled = False
                    ''Btn_Imprime_Contra_Recibo.Visible = True
                    Btn_Imprime_Contra_Recibo.Visible = False
                    MsgBox "La factura ha sido modificada exitosamente", vbInformation
                Else
                    MsgBox "No existe el número de factura", vbExclamation
                End If
                Rs_Modifica_Facturas_Proveedores.Close
                Conexion_Base.CommitTrans
            End If
        Else
            MsgBox "Faltan datos para capturar la factura", vbExclamation
        End If
    End If
    Exit Sub
Handler:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Nuevo_Click()
Dim No_Movimiento As String
Dim Fila As Integer
Dim Rs_Editar_Pedido_Detalles As rdoResultset
Dim Rs_Editar_Entrada_Detalles As rdoResultset
Dim Rs_Consulta As rdoResultset
Dim No_Factura As String

If Btn_Nuevo.Caption = "Nuevo" Then
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        DTP_Fecha_Factura.Value = Now
        DTP_Fecha_Recepcion.Value = Now
        DTP_Fecha_Pago.Value = Now
        Txt_Impuesto_Cedular.Text = ""
        Grid_Compras.Rows = 0
        Chk_Seleccionar.Visible = False
        Pic_Busqueda_Facturas_Proveedores.Visible = False
        Pic_Facturas_Proveedores.Visible = True
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Salir.Visible = False
        Btn_Regresar.Visible = True
        Cmb_Estatus.Visible = True
        Cmb_Tipo.ListIndex = 0
        Cmb_Moneda.ListIndex = 0
        Cmb_Estatus.ListIndex = 0
        Btn_Pagos.Visible = False
        Lbl_Estatus.Visible = True
        Txt_No_Factura.Locked = False
        Fra_Datos_Factura.Enabled = True
        Fra_Nota_Credito.Enabled = True
        Fra_Anticipos.Enabled = True
        Fra_Cantidades.Enabled = True
        Fra_Condiciones_Pago.Enabled = True
        Fra_Compras.Enabled = True
        Btn_Imprime_Contra_Recibo.Visible = False
        Txt_Contra_Recibo.Text = Conectar_Ayudante.Maximo_Catalogo("Adm_Proveedores_Facturas", "Contra_Recibo")
        Btn_Nuevo.Caption = "Dar de Alta"
        Cmb_Orden_Compra_KeyPress 13
        Cmb_Proveedor.SetFocus
    Else
         If Txt_No_Factura.Text <> "" And Val(Txt_Saldo.Text) >= 0 And Cmb_Proveedor.ListIndex > -1 Then
            If Cmb_Moneda.Text <> "PESOS" And Val(Txt_Tipo_Cambio.Text) <= 0 Then
                MsgBox "Falta capturar el tipo de cambio", vbExclamation
                Txt_Tipo_Cambio.SetFocus
                Exit Sub
            End If
            If Txt_Nota_Credito.Text = "" And Val(Txt_Monto_Nota.Text) > 0 Then
                MsgBox "Falta el número de nota de credito para aplicarla", vbInformation
                Exit Sub
            End If
            'Busca las facturas con número o monto parecido para ver si la debe dar de alta
            Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND (No_Factura LIKE '%" & Val(Txt_No_Factura.Text) & "%'"
            Mi_SQL = Mi_SQL & " OR Total=" & Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) & ")"
            Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Facturas.EOF Then
                Mensaje_Facturas = "Las siguientes facturas conciden con folio o el monto"
                Mensaje_Facturas = Mensaje_Facturas & Chr(13) & "Folio" & Chr(9) & "" & Chr(9) & "Total"
                While Not Rs_Consulta_Facturas.EOF
                    Mensaje_Facturas = Mensaje_Facturas & Chr(13) & Rs_Consulta_Facturas.rdoColumns("No_Factura") & Chr(9) & "" & Chr(9) & Format(Rs_Consulta_Facturas.rdoColumns("Total"), "#,###,##0.00")
                    No_Factura = Rs_Consulta_Facturas.rdoColumns("No_Factura")
                    Rs_Consulta_Facturas.MoveNext
                Wend
                MsgBox Mensaje_Facturas, vbExclamation
                If Format(No_Factura, "0000000000") = Format(Txt_No_Factura.Text, "0000000000") Then Exit Sub
                If MsgBox("¿Desea ingresarla de todas formas?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Rs_Consulta_Facturas.Close
            Conexion_Base.BeginTrans
            Set Rs_Agrega_Facturas_Proveedores = Conectar_Ayudante.Recordset_Agregar("Adm_Proveedores_Facturas")
            With Rs_Agrega_Facturas_Proveedores
                .AddNew
                    .rdoColumns("No_Factura") = UCase(Trim(Txt_No_Factura.Text))
                    .rdoColumns("Tipo") = Cmb_Tipo.Text
                    .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                    .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Recepcion") = Format(DTP_Fecha_Recepcion.Value, "MM/dd/yyyy")
                    .rdoColumns("Moneda") = Cmb_Moneda.Text
                    .rdoColumns("Tipo_Cambio") = Val(Txt_Tipo_Cambio.Text)
                    .rdoColumns("Cantidad") = 1
                    .rdoColumns("Precio_Unitario") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
                    .rdoColumns("Flete") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Flete.Text, ","))
                    .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
                    .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))
                    .rdoColumns("Retencion_IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Retencion_IVA.Text, ","))
                    .rdoColumns("Retencion_ISR") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Retencion_ISR.Text, ","))
                    If Cmb_Tipo.Text = "Honorarios" Then
                        .rdoColumns("Impuesto_Cedular") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Impuesto_Cedular.Text, ","))
                        .rdoColumns("Retencion_Fletes") = 0
                    Else
                        .rdoColumns("Retencion_Fletes") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Impuesto_Cedular.Text, ","))
                        .rdoColumns("Impuesto_Cedular") = 0
                    End If
                    .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
                    .rdoColumns("Abono") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ","))
                    .rdoColumns("Saldo") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Saldo.Text, ","))
                    If Val(Txt_Saldo.Text) = 0 Then
                        .rdoColumns("Pagada") = "S"
                    Else
                        'Si es de autoservicio la pone como pagada
                        If Cmb_Tipo.Text = "Autoservicio" Then
                            .rdoColumns("Pagada") = "S"
                            .rdoColumns("Abono") = .rdoColumns("Total")
                            .rdoColumns("Saldo") = 0
                        Else
                            .rdoColumns("Pagada") = "N"
                        End If
                    End If
                    .rdoColumns("Cancelada") = "N"
                    .rdoColumns("Fecha_Pago") = Format(DTP_Fecha_Pago.Value, "MM/dd/yyyy")
                    .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.Text
                    .rdoColumns("Comentarios") = Txt_Comentarios.Text
                    .rdoColumns("Contra_Recibo") = Txt_Contra_Recibo.Text
                         'Actualiza el Estatus de los detalles del pedido
                         For Fila = 1 To Grid_Compras.Rows - 1 Step 1
                            If Grid_Compras.TextMatrix(Fila, 1) = "SI" Then
                                'Actualiza estatus de los detalles de la entrada
                                Mi_SQL = " SELECT Alm_Entradas.Entrada_ID  FROM Alm_Entradas  "
                                Mi_SQL = Mi_SQL & " WHERE Alm_Entradas.No_Control = '" & Grid_Compras.TextMatrix(Fila, 0) & "' "
                                Mi_SQL = Mi_SQL & " AND Alm_Entradas.Proveedor_ID = '" & Grid_Compras.TextMatrix(Fila, 10) & "' "
                                Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                If Not Rs_Consulta.EOF Then
                                    While Not Rs_Consulta.EOF
                                        Mi_SQL = " SELECT Alm_Entradas_Detalles.Estatus  FROM Alm_Entradas_Detalles  "
                                        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Detalles.Entrada_ID='" & Rs_Consulta.rdoColumns("Entrada_ID") & "'  "
                                        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Detalles.Producto_ID = '" & Grid_Compras.TextMatrix(Fila, 11) & "' "
                                        Set Rs_Editar_Entrada_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                                        If Not Rs_Editar_Entrada_Detalles.EOF Then
                                            With Rs_Editar_Entrada_Detalles
                                                .Edit
                                                    .rdoColumns("Estatus") = "FACTURADA"
                                                .Update
                                            End With
                                        End If
                                        Rs_Editar_Entrada_Detalles.Close
                                    Rs_Consulta.MoveNext
                                    Wend
                                End If
                                Rs_Consulta.Close
                            End If
                         Next
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
            Rs_Agrega_Facturas_Proveedores.Close
            'Valida el total de la factura siendo en dolares o en pesos
            If Cmb_Moneda.Text = "PESOS" Then
                Total_Factura = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
            Else
                Total_Factura = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) * Val(Txt_Tipo_Cambio.Text)
            End If
            'Actualiza las compras asignándoles la factura capturada
            For I = 1 To Grid_Compras.Rows - 1
                If Trim(Grid_Compras.TextMatrix(I, 1)) = "SI" Then
                    Mi_SQL = "SELECT No_Control,No_Factura_Proveedor,Aplicada,Usuario_Aplico"
                    Mi_SQL = Mi_SQL & " FROM Tmp_Proveedores_Facturas"
                    Mi_SQL = Mi_SQL & " WHERE No_Control='" & Trim(Grid_Compras.TextMatrix(I, 0)) & "'"
                    Set Rs_Actualiza_Compras = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Actualiza_Compras.EOF Then
                        Rs_Actualiza_Compras.Edit
                            Rs_Actualiza_Compras.rdoColumns("No_Factura_Proveedor") = UCase(Trim(Txt_No_Factura.Text))
                            Rs_Actualiza_Compras.rdoColumns("Aplicada") = "SI"
                            Rs_Actualiza_Compras.rdoColumns("Usuario_Aplico") = Nombre_Usuario
                        Rs_Actualiza_Compras.Update
                    End If
                    Rs_Actualiza_Compras.Close
                End If
            Next I
            'Aplica los anticipos
            For I = 0 To Lst_Anticipos.ListCount - 1
                If Lst_Anticipos.Selected(I) = True Then
                    'Consulta el anticipo para ver si esta aplicado
                    Mi_SQL = "SELECT * FROM Adm_Proveedores_Anticipos"
                    Mi_SQL = Mi_SQL & " WHERE No_Anticipo=" & Format(Mid(Lst_Anticipos.List(I), 1, 10), "0000000000")
                    Set Rs_Modifica_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Modifica_Anticipos_Proveedores
                        If Not .EOF Then
                            No_Movimiento = Rs_Modifica_Anticipos_Proveedores.rdoColumns("No_Movimiento")
                            .Edit
                                .rdoColumns("Pago") = .rdoColumns("Total")
                                .rdoColumns("Aplicado") = "S"
                                .rdoColumns("No_Factura") = Txt_No_Factura.Text
                                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                .rdoColumns("Fecha_Modifico") = Now
                            .Update
                        End If
                    End With
                    
                    'Pone el número de factura en el movimiento administrativo
                    Mi_SQL = "SELECT * FROM Adm_Movimientos "
                    Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & No_Movimiento & "'"
                    Set Rs_Modifica_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Modifica_Movimientos
                        If Not .EOF Then
                            .Edit
                                .rdoColumns("No_Factura") = Txt_No_Factura.Text
                                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                .rdoColumns("Fecha_Modifico") = Now
                            .Update
                        End If
                    End With
                    Rs_Modifica_Anticipos_Proveedores.Close
                    Rs_Modifica_Movimientos.Close
                End If
            Next I
            
            
            'Captura la nota de crédito
            If Txt_Nota_Credito.Text <> "" And Val(Txt_Monto_Nota.Text) > 0 Then
            Set Rs_Agrega_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Agregar("Adm_Proveedores_Notas_Credito")
                With Rs_Agrega_Notas_Credito_Proveedores
                    .AddNew
                        .rdoColumns("No_Nota_Credito") = Txt_Nota_Credito.Text
                        .rdoColumns("Fecha") = Format(DTP_Fecha_Factura.Value, "MM/dd/yyyy")
                        .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                        .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ","))
                        .rdoColumns("No_Factura") = Txt_No_Factura.Text
                        .rdoColumns("Usuario_Creo") = Nombre_Usuario
                        .rdoColumns("Fecha_Creo") = Now
                    .Update
                End With
                Rs_Agrega_Notas_Credito_Proveedores.Close
            End If

            Conexion_Base.CommitTrans
            If MsgBox("Factura Capturada" & Chr(13) & Chr(13) & "¿Desea capturar otra factura del mismo proveedor con el mismo contra recibo?", vbYesNo + vbQuestion) = vbYes Then
                Txt_Subtotal.Text = ""
                Txt_No_Factura = ""
                Txt_Nota_Credito.Text = ""
                Txt_Monto_Nota.Text = ""
                Txt_Saldo.Text = ""
                Txt_Anticipo.Text = ""
                ''Txt_Cuenta_Proveedor.Text = ""
                ''Txt_Cuenta_Gasto.Text = ""
                Call Cmb_Proveedor_Click
            Else
''                If MsgBox("¿Desea imprimir el Contra Recibo?", vbYesNo + vbQuestion) = vbYes Then
''                    Respuesta = Val(InputBox("¿Cuantas copias desea imprimir del contrarrecibo?", "Impresión Contra Recibo"))
''                    If Respuesta > 0 And Respuesta <= 10 Then
''                        For I = 1 To Respuesta
''                            Imprime_Contra_Recibo
''                        Next I
''                    End If
''                End If
                Fra_Datos_Factura.Enabled = False
                Fra_Nota_Credito.Enabled = False
                Fra_Anticipos.Enabled = False
                Fra_Cantidades.Enabled = False
                Fra_Condiciones_Pago.Enabled = False
                Fra_Compras.Enabled = False
                Btn_Nuevo.Caption = "Nuevo"
                Btn_Modificar.Enabled = True
                Btn_Buscar.Enabled = True
                Btn_Eliminar.Enabled = True
                Btn_Salir.Caption = "Salir"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Cmb_Tipo.ListIndex = -1
                Cmb_Moneda.ListIndex = -1
                Lst_Anticipos.Clear
                Cmb_Proveedor.Text = ""
                Cmb_Proveedor.Clear
                Grid_Compras.Rows = 0
                Chk_Seleccionar.Visible = False
            End If
        Else
            MsgBox "Faltan datos para capturar la factura", vbExclamation
            Exit Sub
        End If
    End If
    Exit Sub
Handler:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Pagos_Click()
    If Grid_Consulta_Facturas.RowSel > 0 Then
        Abrir_Movimiento = True
        Movimiento_Factura = Txt_No_Factura.Text
        Load Frm_Adm_Movimientos_Consulta
        'Frm_Adm_Consulta_Movimientos.No_Factura = Trim(Txt_No_Factura.Text)
        Call Frm_Adm_Movimientos_Consulta.Consulta_Movimientos(Trim(Txt_No_Factura.Text))
        'Frm_Adm_Consulta_Movimientos.Btn_Consultar_Click
    End If
End Sub

Private Sub Btn_Regresar_Click()
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Lst_Anticipos.Clear
        Grid_Compras.Rows = 0
        Chk_Seleccionar.Visible = False
        If Btn_Nuevo.Caption = "Dar de Alta" Then
            Btn_Nuevo.Caption = "Nuevo"
        Else
            Btn_Nuevo.Enabled = True
        End If
        If Btn_Modificar.Caption = "Actualizar" Then
            Btn_Modificar.Caption = "Modificar"
        Else
            Btn_Modificar.Enabled = True
        End If
        Btn_Buscar.Enabled = True
        Btn_Eliminar.Enabled = True
        Cmb_Tipo.ListIndex = -1
        Cmb_Moneda.ListIndex = -1
        Cmb_Proveedor.Clear
        Btn_Salir.Visible = True
        Btn_Regresar.Visible = False
        Fra_Datos_Factura.Enabled = False
        Fra_Nota_Credito.Enabled = False
        Fra_Anticipos.Enabled = False
        Fra_Cantidades.Enabled = False
        Fra_Condiciones_Pago.Enabled = False
        Fra_Compras.Enabled = False
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    End If
End Sub

Private Sub Btn_Seleccionar_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset
Dim Rs_Anticipos_Proveedores As rdoResultset
Dim Rs_Consulta_Notas_Credito_Proveedores As rdoResultset
Dim Rs_Consulta_Tipo_Pagos As rdoResultset
Dim Rs_Consulta_Ope_Presupuestos As rdoResultset
Dim Rs_Consulta_Compras As rdoResultset
Dim Rs_Consulta_Pedidos As rdoResultset
Dim I As Integer
Dim Indice As Integer
Dim Cantidad_Pedida As Double

    'Consulta las facturas
    Indice = 8
    Mi_SQL = "SELECT Adm_Proveedores_Facturas.*,Cat_Proveedores.Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas,Cat_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Adm_Proveedores_Facturas.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    Mi_SQL = Mi_SQL & " AND No_Factura='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1) & "'"
    Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Facturas.Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, Indice) & "'"
    Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Facturas_Proveedores
        If Not .EOF Then
            Txt_No_Factura.Text = .rdoColumns("No_Factura")
            'Consulta el tipo de factura
            Select Case .rdoColumns("Tipo")
                Case "Factura"
                    Cmb_Tipo.ListIndex = 0
                Case "Honorarios"
                    Cmb_Tipo.ListIndex = 1
                Case "Fletes"
                    Cmb_Tipo.ListIndex = 2
                Case "Sin IVA"
                    Cmb_Tipo.ListIndex = 3
                Case "Gasolina"
                    Cmb_Tipo.ListIndex = 4
                Case "Arrendamiento"
                    Cmb_Tipo.ListIndex = 5
                Case "Otros"
                    Cmb_Tipo.ListIndex = 6
                Case "Entrada"
                    Cmb_Tipo.ListIndex = 7
                Case "Autoservicio"
                    Cmb_Tipo.ListIndex = 8
            End Select
            Cmb_Orden_Compra = .rdoColumns("Orden_Compra")
            Cmb_Proveedor.Clear
            Cmb_Proveedor.AddItem .rdoColumns("Nombre")
            Cmb_Proveedor.ItemData(Cmb_Proveedor.NewIndex) = .rdoColumns("Proveedor_ID")
            Cmb_Proveedor.ListIndex = 0
            DTP_Fecha_Factura.Value = .rdoColumns("Fecha")
            DTP_Fecha_Recepcion.Value = .rdoColumns("Fecha_Recepcion")
            If UCase(Trim(.rdoColumns("Moneda"))) = "PESOS" Then
                Cmb_Moneda.ListIndex = 0
            Else
                Cmb_Moneda.Text = Trim(.rdoColumns("Moneda"))
                Txt_Tipo_Cambio.Text = .rdoColumns("Tipo_Cambio")
            End If
            Txt_Subtotal.Text = Format(.rdoColumns("Importe"), "#,###,##0.00")
            If Cmb_Tipo.Text = "Gasolina" Or Cmb_Tipo.Text = "Otros" Then
                Txt_IVA.Text = Format(.rdoColumns("IVA"), "#,###,##0.00")
                Txt_Retencion_IVA.Text = Format(.rdoColumns("Retencion_IVA"), "#,###,##0.00")
                Txt_Retencion_ISR.Text = Format(.rdoColumns("Retencion_ISR"), "#,###,##0.00")
                Txt_Impuesto_Cedular.Text = Format(.rdoColumns("Retencion_Fletes"), "#,###,##0.00")
            End If
            Txt_Flete.Text = Format(.rdoColumns("Flete"), "#,###,##0.00")
            DTP_Fecha_Pago.Value = .rdoColumns("Fecha_Pago")
            If Not IsNull(.rdoColumns("Forma_Pago")) Then Cmb_Forma_Pago.Text = .rdoColumns("Forma_Pago") Else Cmb_Forma_Pago.Text = ""
            Txt_Comentarios.Text = .rdoColumns("Comentarios")
            If Not IsNull(.rdoColumns("Contra_Recibo")) Then Txt_Contra_Recibo.Text = .rdoColumns("Contra_Recibo") Else Txt_Contra_Recibo.Text = ""
            Btn_Nuevo.Visible = True
            If .rdoColumns("Pagada") = "S" Then
                Cmb_Estatus.Text = "Pagada"
                Btn_Modificar.Enabled = False
            Else
                If .rdoColumns("Cancelada") = "S" Then
                    Cmb_Estatus.Text = "Cancelada"
                    Btn_Modificar.Enabled = False
                Else
                    If .rdoColumns("Abono") > 0 Then
                        Cmb_Estatus.Text = "Abonada"
                        Btn_Modificar.Enabled = False
                    Else
                        Cmb_Estatus.Text = "Sin Pagar"
                        If Rol = "ADMINISTRADOR" Then
                            Btn_Eliminar.Visible = True
                            Btn_Modificar.Enabled = True
                        End If
                        If Cmb_Tipo.ListIndex <> 2 Then
                            Btn_Modificar.Visible = True
                        End If
                    End If
                End If
            End If
            'CONSULTA LOS ANTICIPOS DE LA FACTURA
            Mi_SQL = "SELECT No_Anticipo,Concepto,Total FROM Adm_Proveedores_Anticipos"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND No_Factura='" & Txt_No_Factura.Text & "'"
            Mi_SQL = Mi_SQL & " AND Aplicado='S'"
            Set Rs_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Anticipos_Proveedores
                Lst_Anticipos.Clear
                I = 0
                While Not .EOF
                    Lst_Anticipos.AddItem .rdoColumns("No_Anticipo") & "         " & Format(.rdoColumns("Total"), "#0.00")
                    Lst_Anticipos.Selected(I) = True
                    I = I + 1
                    .MoveNext
                Wend
            End With
            'CONSULTA LA NOTA DE CREDITO SI APLICA
            Mi_SQL = "SELECT No_Nota_Credito,Importe FROM Adm_Proveedores_Notas_Credito "
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND No_Factura='" & Txt_No_Factura.Text & "'"
            Set Rs_Consulta_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Notas_Credito_Proveedores
                If Not .EOF Then
                    Txt_Nota_Credito.Text = .rdoColumns("No_Nota_Credito")
                    Txt_Monto_Nota.Text = Format(.rdoColumns("Importe"), "#,###,###.00")
                End If
            End With
            Pic_Busqueda_Facturas_Proveedores.Visible = False
            Pic_Facturas_Proveedores.Visible = True
            Cmb_Estatus.Visible = True
            Lbl_Estatus.Visible = True
            Btn_Pagos.Visible = True
            If (Cmb_Tipo.Text = "Orden Compra") And (Trim(.rdoColumns("No_Factura")) = Trim(.rdoColumns("Orden_Compra"))) Then
                Txt_No_Factura.Locked = False
            Else
                Txt_No_Factura.Locked = True
            End If
            ''Btn_Imprime_Contra_Recibo.Visible = True
            Btn_Imprime_Contra_Recibo.Visible = False
            Txt_Saldo.Text = Format(.rdoColumns("Saldo"), "#0.00")
            'Consulta las compras asignadas a la factura
            Grid_Compras.Rows = 0
            Grid_Compras.Cols = 12
            'Encabezado
            Grid_Compras.AddItem "No_Control" & Chr(9) & "Aplicada" & Chr(9) & "Documento" & Chr(9) & "Tipo" & Chr(9) & "Recepcion" & Chr(9) & "Total" & Chr(9) & "Estatus" & Chr(9) & "Surtida" & Chr(9) & "Subtotal" & Chr(9) & "IVA" & Chr(9) & "Proveedor_ID" & Chr(9) & "Producto_ID"
            Mi_SQL = "SELECT Tmp_Proveedores_Facturas.No_Control,Tmp_Proveedores_Facturas.No_Factura,Tmp_Proveedores_Facturas.Tipo_Recepcion,Tmp_Proveedores_Facturas.Fecha_Recepcion,Alm_Entradas_Detalles.Importe,Tmp_Proveedores_Facturas.Moneda,Tmp_Proveedores_Facturas.Aplicada,Alm_Entradas_Detalles.Estatus,Alm_Entradas_Detalles.Cantidad,Alm_Entradas_Detalles.Producto_ID,Tmp_Proveedores_Facturas.Subtotal,Tmp_Proveedores_Facturas.IVA,Alm_Entradas.Proveedor_ID"
            Mi_SQL = Mi_SQL & " FROM Tmp_Proveedores_Facturas,Alm_Entradas,Alm_Entradas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Tmp_Proveedores_Facturas.No_Control=Alm_Entradas.No_Control"
            Mi_SQL = Mi_SQL & " AND Alm_Entradas.Entrada_ID=Alm_Entradas_Detalles.Entrada_ID"
            Mi_SQL = Mi_SQL & " AND Tmp_Proveedores_Facturas.Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Tmp_Proveedores_Facturas.No_Factura_Proveedor ='" & Trim(Txt_No_Factura.Text) & "'"
            Set Rs_Consulta_Compras = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Compras.EOF
                Grid_Compras.AddItem Rs_Consulta_Compras.rdoColumns("No_Control") _
                    & Chr(9) & "SI" _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("No_Factura") _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("Tipo_Recepcion") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Recepcion"), "dd/MMM/yy") _
                    & Chr(9) & Format(Val(.rdoColumns("Importe")) + Val(.rdoColumns("IVA")), "#,###,##0.00") _
                    & Chr(9) & Trim(Rs_Consulta_Compras.rdoColumns("Estatus")) _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("Cantidad") _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("Importe") _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("IVA") _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("Proveedor_ID") _
                    & Chr(9) & Rs_Consulta_Compras.rdoColumns("Producto_ID")
                Grid_Compras.FixedRows = 1
                Rs_Consulta_Compras.MoveNext
            Wend
            Grid_Compras.ColWidth(0) = 0        'No_Control
            Grid_Compras.ColWidth(1) = 350      'Aplica
            Grid_Compras.ColWidth(2) = 1100     'Factura
            Grid_Compras.ColAlignment(2) = flexAlignCenterCenter
            Grid_Compras.ColWidth(3) = 1000     'Tipo
            Grid_Compras.ColWidth(4) = 1000     'Recepcion
            Grid_Compras.ColAlignment(4) = flexAlignCenterCenter
            Grid_Compras.ColWidth(5) = 1200     'Total
            Grid_Compras.ColAlignment(5) = flexAlignCenterCenter
            Grid_Compras.ColWidth(6) = 1000     'Estatus
            Grid_Compras.ColWidth(7) = 800      'Surtida
            Grid_Compras.ColWidth(8) = 0        'Subtotal
            Grid_Compras.ColWidth(9) = 0       'IVA
            Grid_Compras.ColWidth(10) = 0       'Proveedor_ID
            Grid_Compras.ColWidth(11) = 0       'Producto_ID
            Rs_Consulta_Compras.Close
        End If
    End With
    Rs_Consulta_Facturas_Proveedores.Close
    Chk_Seleccionar.Value = 1
    Call Chk_Seleccionar_Click
End Sub

Private Sub Chk_Busqueda_Estatus_Click()
    If Chk_Busqueda_Estatus.Value = 1 Then
        Cmb_Con_Estatus.Visible = True
        Cmb_Con_Estatus.SetFocus
    Else
        Cmb_Con_Estatus.Visible = False
    End If
End Sub

Private Sub Chk_Busqueda_Fecha_Factura_Click()
    If Chk_Busqueda_Fecha_Factura.Value = 1 Then
        DTP_Fecha_Factura_Inicial.Visible = True
        DTP_Fecha_Factura_Final.Visible = True
        DTP_Fecha_Factura_Inicial.SetFocus
    Else
        DTP_Fecha_Factura_Inicial.Visible = False
        DTP_Fecha_Factura_Final.Visible = False
    End If
End Sub

Private Sub Chk_Busqueda_Fecha_Recepcion_Click()
    If Chk_Busqueda_Fecha_Recepcion.Value = 1 Then
        DTP_Fecha_Recepcion_Inicial.Visible = True
        DTP_Fecha_Recepcion_Final.Visible = True
        DTP_Fecha_Recepcion_Inicial.SetFocus
    Else
        DTP_Fecha_Recepcion_Inicial.Visible = False
        DTP_Fecha_Recepcion_Final.Visible = False
    End If
End Sub


Private Sub Chk_Busqueda_No_Factura_Click()
    If Chk_Busqueda_No_Factura.Value = 1 Then
        Txt_Con_No_Factura.Visible = True
        Txt_Con_No_Factura.SetFocus
    Else
        Txt_Con_No_Factura.Visible = False
    End If
End Sub

Private Sub Chk_Busqueda_Proveedor_Click()
    If Chk_Busqueda_Proveedor.Value = 1 Then
        Cmb_Con_Proveedor.Visible = True
        Cmb_Con_Proveedor.SetFocus
    Else
        Cmb_Con_Proveedor.Visible = False
    End If
End Sub


Private Sub Chk_Seleccionar_Click()
Dim Fila As Integer
Dim Suma_IVA As Double
Dim Suma_Subtotal As Double

    If Grid_Compras.Rows > 1 Then
        If Chk_Seleccionar.Value = 0 Then
            Grid_Compras.TextMatrix(Grid_Compras.RowSel, 1) = "NO"
        Else
            If Trim(Grid_Compras.TextMatrix(Grid_Compras.RowSel, 6)) <> "" Then
                If Trim(Grid_Compras.TextMatrix(Grid_Compras.RowSel, 6)) = "RECEPCION" Then
                    Grid_Compras.TextMatrix(Grid_Compras.RowSel, 1) = "SI"
                    Call Calcula_Cantidades
                End If
            End If
        End If
        Txt_Subtotal.Text = ""
        Txt_IVA.Text = ""
        Txt_Total.Text = ""
        Txt_Saldo.Text = ""
        For Fila = 1 To Grid_Compras.Rows - 1 Step 1
            If Grid_Compras.TextMatrix(Fila, 1) = "SI" Then
                Suma_IVA = Suma_IVA + Val(Grid_Compras.TextMatrix(Fila, 9))
                Suma_Subtotal = Suma_Subtotal + Val(Grid_Compras.TextMatrix(Fila, 8))
            End If
        Next Fila
        ''Txt_Subtotal.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Suma_Subtotal
        ''Txt_IVA.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ",")) + Suma_IVA
        ''Txt_Total.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))
        ''Txt_Saldo.Text = (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))) - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ","))
        Txt_Subtotal.Text = Suma_Subtotal
        Txt_IVA.Text = Suma_IVA
        Txt_Total.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))
        Txt_Saldo.Text = (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))) - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ","))
    End If
End Sub

Private Sub Cmb_Moneda_Click()
    If Cmb_Moneda.Text = "DOLARES" Then
        Lbl_Tipo_Cambio.Visible = True
        Txt_Tipo_Cambio.Visible = True
    End If
    If Cmb_Moneda.Text = "PESOS" Then
        Lbl_Tipo_Cambio.Visible = False
        Txt_Tipo_Cambio.Visible = False
    End If
End Sub


Private Sub Cmb_Orden_Compra_Click()

End Sub

Private Sub Cmb_Orden_Compra_KeyPress(KeyAscii As Integer)

   

End Sub


Private Sub Cmb_Proveedor_Change()
       Grid_Compras.Rows = 0
End Sub

Private Sub Cmb_Proveedor_Click()
Dim Rs_Consulta_Anticipos_Proveedores As rdoResultset
Dim Rs_Consulta_Compras As rdoResultset
Dim Rs_Consulta_Pedidos As rdoResultset
Dim Cantidad_Pedida As Double

    If Cmb_Proveedor.ListIndex > -1 Then
        If Btn_Nuevo.Caption = "Dar de Alta" Then
            'Consulta los anticipos
            Mi_SQL = "SELECT No_Anticipo,Concepto,Total,Pago"
            Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Anticipos"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
            Mi_SQL = Mi_SQL & " AND Aplicado='N'"
            Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Anticipos_Proveedores
                Lst_Anticipos.Clear
                While Not .EOF
                    Lst_Anticipos.AddItem .rdoColumns("No_Anticipo") & "         " & Format(.rdoColumns("Total"), "#0.00")
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Anticipos_Proveedores.Close
        End If
        'Consulta la factura de compras para mostrar en pantalla la información
        Mi_SQL = "SELECT Tmp_Proveedores_Facturas.No_Control,Tmp_Proveedores_Facturas.No_Factura,Tmp_Proveedores_Facturas.Tipo_Recepcion,Tmp_Proveedores_Facturas.Fecha_Recepcion,Tmp_Proveedores_Facturas.Total,Tmp_Proveedores_Facturas.Moneda,Alm_Entradas_Detalles.Estatus,Alm_Entradas_Detalles.Cantidad,Tmp_Proveedores_Facturas.Subtotal,Tmp_Proveedores_Facturas.IVA as IVA_Factura,Alm_Entradas.Proveedor_ID,Alm_Entradas_Detalles.Importe,Alm_Entradas_Detalles.IVA,Alm_Entradas_Detalles.Cantidad,Alm_Entradas_Detalles.Producto_ID"
        Mi_SQL = Mi_SQL & " FROM Tmp_Proveedores_Facturas,Alm_Entradas,Alm_Entradas_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Tmp_Proveedores_Facturas.No_Control=Alm_Entradas.No_Control"
        Mi_SQL = Mi_SQL & " AND Tmp_Proveedores_Facturas.Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Entrada_ID = Alm_Entradas_Detalles.Entrada_ID"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Detalles.Estatus='RECEPCION' "
        If Btn_Modificar.Caption = "Actualizar" Then
            Mi_SQL = Mi_SQL & " AND (Tmp_Proveedores_Facturas.Aplicada='NO' "
        End If
        Set Rs_Consulta_Compras = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Grid_Compras.Rows = 0
        Grid_Compras.Cols = 13
        'Encabezado
        Grid_Compras.AddItem "No_Control" & Chr(9) & "Aplicada" & Chr(9) & "Documento" & Chr(9) & "Tipo" & Chr(9) & "Recepcion" & Chr(9) & "Total" & Chr(9) & "Estatus" & Chr(9) & "Surtida" & Chr(9) & "Subtotal" & Chr(9) & "IVA" & Chr(9) & "Proveedor_ID" & Chr(9) & "Producto_ID"
        With Rs_Consulta_Compras
            While Not Rs_Consulta_Compras.EOF
                Grid_Compras.AddItem .rdoColumns("No_Control") _
                    & Chr(9) & "NO" _
                    & Chr(9) & .rdoColumns("No_Factura") _
                    & Chr(9) & .rdoColumns("Tipo_Recepcion") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Recepcion"), "dd/MMM/yy") _
                    & Chr(9) & Format(Val(.rdoColumns("Importe")) + Val(.rdoColumns("IVA")), "#,###,##0.00") _
                    & Chr(9) & Trim(.rdoColumns("Estatus")) _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Importe") _
                    & Chr(9) & .rdoColumns("IVA") _
                    & Chr(9) & .rdoColumns("Proveedor_ID") _
                    & Chr(9) & .rdoColumns("Producto_ID")
                Grid_Compras.FixedRows = 1
                .MoveNext
            Wend
        End With
        Rs_Consulta_Compras.Close
        Grid_Compras.ColWidth(0) = 0        'No_Control
        Grid_Compras.ColWidth(1) = 350      'Aplica
        Grid_Compras.ColWidth(2) = 1100     'Factura
        Grid_Compras.ColAlignment(2) = flexAlignCenterCenter
        Grid_Compras.ColWidth(3) = 1000     'Tipo
        Grid_Compras.ColWidth(4) = 1000     'Recepcion
        Grid_Compras.ColAlignment(4) = flexAlignCenterCenter
        Grid_Compras.ColWidth(5) = 1200     'Total
        Grid_Compras.ColAlignment(6) = flexAlignCenterCenter
        Grid_Compras.ColWidth(7) = 1000     'Estatus
        Grid_Compras.ColWidth(8) = 800      'Surtida
        Grid_Compras.ColWidth(9) = 0        'Subtotal
        Grid_Compras.ColWidth(10) = 0       'IVA
        Grid_Compras.ColWidth(11) = 0       'Proveedor_ID
        Grid_Compras.ColWidth(12) = 0       'Producto_ID
   End If
End Sub


Private Sub Cmb_Proveedor_KeyPress(KeyAscii As Integer)
Dim Despliega_Lista As Long

    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Proveedor.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Tipo_Click()
    If Cmb_Tipo.Text = "Honorarios" Or Cmb_Tipo.Text = "Arrendamiento" Then
        Lbl_Impuesto_Cedular.Caption = "Impuesto Cedular"
        Lbl_Subtotal.Caption = Cmb_Tipo.Text
    Else
        Lbl_Impuesto_Cedular.Caption = "Retencion Fletes"
        Lbl_Subtotal.Caption = "SubTotal"
    End If
    If Cmb_Tipo.Text = "Gasolina" Then
        Lbl_Fetes.Caption = "IEPS"
    Else
        Lbl_Fetes.Caption = "Flete"
    End If
    Calcula_Cantidades
End Sub


Private Sub Form_Load()
    Me.Height = 7830
    Me.Width = 7230
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
    Txt_Contra_Recibo.Text = Conectar_Ayudante.Maximo_Catalogo("Adm_Proveedores_Facturas", "Contra_Recibo")
    DTP_Fecha_Factura.Value = Now
    DTP_Fecha_Recepcion.Value = Now
    DTP_Fecha_Pago.Value = Now
    Call Cmb_Proveedor_KeyPress(13)
End Sub


Private Sub Grid_Compras_Click()
    If Grid_Compras.Rows > 1 Then
        Chk_Seleccionar.Visible = False
        If Grid_Compras.ColSel = 1 Then
            Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Compras, Chk_Seleccionar)
            Grid_Compras.TextMatrix(Grid_Compras.RowSel, 1) = ""
            If Grid_Compras.TextMatrix(Grid_Compras.RowSel, 1) = "SI" Then
                Chk_Seleccionar.Value = 1
            Else
                Chk_Seleccionar.Value = 0
            End If
        End If
    End If
End Sub
Private Sub Grid_Compras_LeaveCell()
    Chk_Seleccionar.Visible = False
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Calcula_Cantidades()
'DESCRIPCIÓN         : Realiza el calculo de retencion de impuestos
'PARÁMETROS          :
'CREO                : Julio Cruz
'FECHA_CREO          : 4-Oct-2010
'MODIFICO            :
'FECHA_MODIFICO      :
'CAUSA_MODIFICACIÓN  :
'*******************************************************************************
Public Sub Calcula_Cantidades()
Dim SubTotal As Double
Dim Monto_Nota As Double
Dim Anticipo As Double


    If Txt_Subtotal.Text <> "" Then
        If (Val(Txt_Total.Text) = 0 And Chk_Seleccionar.Value = 0) Or (Txt_Total.Text <> "" And Lst_Anticipos.SelCount > 0) Or (Txt_Flete.Text <> "") Or (Txt_Subtotal.Text <> "") Then
            'VALIDA SI EL TIPO ES DE FACTURA ES GASOLINA
            If Cmb_Tipo.Text = "Gasolina" Then
                SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
            Else
                SubTotal = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Flete.Text, ","))
            End If
            If Txt_Monto_Nota.Text <> "" Then Monto_Nota = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Monto_Nota.Text, ","))
            If Txt_Anticipo.Text <> "" Then Anticipo = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ","))
            If Cmb_Tipo.Text = "Sin IVA" Then
                Txt_IVA.Text = "0.00"
            Else
                ''If Cmb_Tipo.Text <> "Gasolina" Then
                    Txt_IVA.Text = Format(Val(SubTotal * PG_Retencion_IVA), "###,##0.00")
                ''End If
            End If
            If Cmb_Tipo.Text = "Honorarios" Or Cmb_Tipo.Text = "Arrendamiento" Then
                Txt_Retencion_IVA.Text = Format(SubTotal * Porcentaje_IVA, "###,##0.00")
                Txt_Retencion_ISR.Text = Format(SubTotal * PG_Retencion_ISR, "###,##0.00")
                Txt_Impuesto_Cedular.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) * PG_Impuesto_Cedular, "###,##0.00")
            End If
            If Cmb_Tipo.Text = "Fletes" Then
                Txt_Retencion_IVA.Text = 0
                Txt_Retencion_ISR.Text = 0
                Txt_Impuesto_Cedular.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) * PG_Retencion_Flete, "###,##0.00")
            End If
            If Cmb_Tipo.Text = "Factura" Or Cmb_Tipo.Text = "Sin IVA" Then
                Txt_Retencion_IVA.Text = 0
                Txt_Retencion_ISR.Text = 0
                Txt_Impuesto_Cedular.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Flete.Text, ",")) * PG_Retencion_Flete, "###,##0.00")
            End If
            If Cmb_Tipo.Text = "Gasolina" Then
                Txt_Total.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Flete.Text, ",")) + Val(SubTotal) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ",")), "###,##0.00") - Val(Txt_Retencion_IVA.Text) - Val(Txt_Retencion_ISR.Text) - Val(Txt_Impuesto_Cedular.Text)
            Else
                Txt_Total.Text = Format(Val(SubTotal) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ",")), "###,##0.00") - Val(Txt_Retencion_IVA.Text) - Val(Txt_Retencion_ISR.Text) - Val(Txt_Impuesto_Cedular.Text)
            End If
            Txt_Saldo.Text = Format(Val(Txt_Total.Text) - Anticipo - Monto_Nota, "###,##0.00")
            'formatea los text
            Txt_IVA.Text = Format(Txt_IVA.Text, "###,##0.00")
            Txt_Retencion_IVA.Text = Format(Txt_Retencion_IVA.Text, "###,##0.00")
            Txt_Retencion_ISR.Text = Format(Txt_Retencion_ISR.Text, "###,##0.00")
            Txt_Total.Text = Format(Txt_Total.Text, "###,##0.00")
            Txt_Saldo.Text = Format(Txt_Saldo.Text, "###,##0.00")
            Txt_Monto_Nota.Text = Format(Txt_Monto_Nota.Text, "###,##0.00")
        End If
    End If
End Sub



Private Sub Grid_Consulta_Facturas_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Facturas As rdoResultset

On Error GoTo Handler
    If Grid_Consulta_Facturas.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
        Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1) & "'"
        Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 8) & "'"
        Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Facturas
            If Not Rs_Consulta_Facturas.EOF Then
                If Rs_Consulta_Facturas.rdoColumns("Cancelada") = "N" Then
                    Btn_Eliminar.Caption = "Cancelar"
                    Btn_Modificar.Enabled = True
                Else
                    Btn_Eliminar.Caption = "Eliminar"
                    Btn_Modificar.Enabled = False
                End If
            End If
        End With
        Rs_Consulta_Facturas.Close
    End If
    Exit Sub
Handler:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Grid_Consulta_Facturas_DblClick()
    If Grid_Consulta_Facturas.Rows > 1 Then
        If Btn_Mas.Caption = "-" Then
            Btn_Mas_Click
        End If
        Call Btn_Seleccionar_Click
        Fra_Datos_Factura.Enabled = False
        Fra_Nota_Credito.Enabled = False
        Fra_Anticipos.Enabled = False
        Fra_Cantidades.Enabled = False
        Fra_Condiciones_Pago.Enabled = True
    End If
End Sub


Private Sub Lst_Anticipos_ItemCheck(Item As Integer)
Dim I As Integer
Dim Rs_Consulta_Anticipos_Proveedores As rdoResultset               'Manejo de Registro

    'Consulta los vales pendientes de facturar
    Mi_SQL = " SELECT Total FROM Adm_Proveedores_Anticipos "
    Mi_SQL = Mi_SQL & " WHERE No_Anticipo = '" & Mid(Lst_Anticipos.List(Item), 1, 10) & "'"
    Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Lst_Anticipos.ListCount > 0 And Not Rs_Consulta_Anticipos_Proveedores.EOF Then
        If Lst_Anticipos.Selected(Item) = True Then
            Txt_Anticipo.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ",")) + Rs_Consulta_Anticipos_Proveedores("Total")
            Call Txt_Anticipo_Change
        Else
            Txt_Anticipo.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Anticipo.Text, ",")) - Rs_Consulta_Anticipos_Proveedores("Total")
            Call Txt_Anticipo_Change
        End If
    End If
    ''Txt_Anticipo.Text = Format(Txt_Anticipo.Text, "###,##0.00")
End Sub


Private Sub Txt_Anticipo_Change()
    Txt_IVA.Text = ""
    Txt_Retencion_IVA.Text = ""
    Txt_Retencion_ISR.Text = ""
    Txt_Total.Text = ""
    Txt_Saldo.Text = ""
    Txt_Monto_Nota.Text = ""
    Call Calcula_Cantidades
    If Chk_Seleccionar.Value = 1 Then Call Chk_Seleccionar_Click
End Sub


Private Sub Txt_Flete_Change()
    Call Calcula_Cantidades
End Sub

Private Sub Txt_Flete_KeyPress(KeyAscii As Integer)
        Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Flete.Text, True)
End Sub


Private Sub Txt_Flete_LostFocus()
        Txt_Flete.Text = Format(Txt_Flete.Text, "#,###,##0.00")
End Sub


Private Sub Txt_Monto_Nota_Change()
    Calcula_Cantidades
End Sub


Private Sub Txt_Monto_Nota_GotFocus()
    If Cmb_Tipo.Text = "Gasolina" Then
        SendKeys "{Home}+{End}"
    End If
End Sub


Private Sub Txt_Monto_Nota_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Monto_Nota, True)
End Sub


Private Sub Txt_Subtotal_Change()
    Txt_IVA.Text = ""
    Txt_Retencion_IVA.Text = ""
    Txt_Retencion_ISR.Text = ""
    Txt_Total.Text = ""
    Txt_Saldo.Text = ""
    Txt_Monto_Nota.Text = ""
    Call Calcula_Cantidades
End Sub


Private Sub Txt_Subtotal_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Subtotal.Text, True)
End Sub


Private Sub Txt_Subtotal_LostFocus()
        Txt_Subtotal.Text = Format(Txt_Subtotal.Text, "#,###,##0.00")
End Sub
'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Imprime_Contra_Recibo()
'   DESCRIPCIÓN: Imprime el saldo de la factura
'   PARÁMETROS:1.- No_Factura: El numero de la factura
'              2.- Saldo: El total del importe realizado
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO          :Rafael Muñoz
'   FECHA_MODIFICO    :28-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Sub Imprime_Contra_Recibo()
Dim Rs_Consulta_Formatos As rdoResultset
Dim Rs_Consulta_Formatos_Generales As rdoResultset
Dim Rs_Consulta_Formatos_Detalles As rdoResultset
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset
Dim Longitud As Integer
Dim I As Integer
Dim Inicio As Integer
Dim Salto As Double
Dim Cont_Renglon As Double
Dim Total_Vale As Double

    Mi_SQL = "SELECT * FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE  Nombre='CONTRA RECIBO'"
    Set Rs_Consulta_Formatos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE  Nombre='CONTRA RECIBO'"
    Mi_SQL = Mi_SQL & " AND Tipo='General'"
    Set Rs_Consulta_Formatos_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE  Nombre='CONTRA RECIBO'"
    Mi_SQL = Mi_SQL & " AND Tipo='Detalle'"
    Set Rs_Consulta_Formatos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Mi_SQL = "SELECT * FROM Adm_Proveedores_Facturas"
    Mi_SQL = Mi_SQL & " WHERE Contra_Recibo=" & Txt_Contra_Recibo.Text
    Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Formatos.EOF Then
        With Rs_Consulta_Formatos
            'Comienza la impresion de la factura
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
        Total_Vale = 0
        While Not Rs_Consulta_Facturas_Proveedores.EOF
            Cont_Renglon = Cont_Renglon + Salto
            While Not Rs_Consulta_Formatos_Detalles.EOF
                Printer.CurrentX = Rs_Consulta_Formatos_Detalles.rdoColumns("X")
                Printer.CurrentY = Rs_Consulta_Formatos_Detalles.rdoColumns("Y") + Cont_Renglon
                Longitud = Rs_Consulta_Formatos_Detalles.rdoColumns("Longitud")
                If Rs_Consulta_Formatos_Detalles.rdoColumns("Campo") = "No_Factura" Then Printer.Print Mid(Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura"), 1, Longitud)
                If Rs_Consulta_Formatos_Detalles.rdoColumns("Campo") = "Fecha" Then Printer.Print Format(Rs_Consulta_Facturas_Proveedores.rdoColumns("Fecha"), "dd/MMM/yy")
                If Rs_Consulta_Formatos_Detalles.rdoColumns("Campo") = "Importe" Then
                    Printer.Print Format(Rs_Consulta_Facturas_Proveedores.rdoColumns("Saldo"), "###,###,###.00")
                    Total_Vale = Total_Vale + Rs_Consulta_Facturas_Proveedores.rdoColumns("Saldo")
                End If
                If Rs_Consulta_Formatos_Detalles.rdoColumns("Campo") = "Comentarios" Then Printer.Print Mid(Rs_Consulta_Facturas_Proveedores.rdoColumns("Comentarios"), 1, Longitud)
                Rs_Consulta_Formatos_Detalles.MoveNext
            Wend
            Rs_Consulta_Formatos_Detalles.MoveFirst
            Rs_Consulta_Facturas_Proveedores.MoveNext
        Wend
            With Rs_Consulta_Formatos
                'Comienza la impresion de la factura
                Printer.ScaleMode = vbCentimeters
                'Configura la fuente de la factura
                Printer.FontSize = .rdoColumns("Tamaño_Generales")
                Printer.Font = .rdoColumns("Letra_Generales")
                If .rdoColumns("Estilo_Generales") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
            End With
            'Imprime la fecha de la factura y la ciudad
            With Rs_Consulta_Formatos_Generales
                While Not .EOF
                    Printer.CurrentX = .rdoColumns("X")
                    Printer.CurrentY = .rdoColumns("Y")
                    Longitud = .rdoColumns("Longitud")
                    If .rdoColumns("Campo") = "Proveedor" Then Printer.Print Mid(Cmb_Proveedor.Text, 1, Longitud)
                    If .rdoColumns("Campo") = "Empresa" Then Printer.Print "Empresa:  " & UCase(Mid(Cmb_Empresa.Text, 2, 17))
                    If .rdoColumns("Campo") = "Usuario" Then Printer.Print "Usuario:  " & Nombre_Usuario
                    If .rdoColumns("Campo") = "Fecha_Factura" Then Printer.Print Format(DTP_Fecha_Recepcion.Value, "dd MMM yyyy")
                    If .rdoColumns("Campo") = "Fecha_Pago" Then Printer.Print Format(DTP_Fecha_Pago.Value, "dd MMM yyyy")
                    If .rdoColumns("Campo") = "Numero" Then Printer.Print Txt_Contra_Recibo.Text
                    If .rdoColumns("Campo") = "Total" Then Printer.Print Format(Total_Vale, "###,###,###.00")
                    .MoveNext
                Wend
            End With
         Printer.EndDoc
    End If
End Sub

'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN    : Consulta_Facturas()
'   DESCRIPCIÓN             : Realiza una consulta a la tabla de Facturas_Proveedores
'   PARÁMETROS              : Generales de Facturas_Proveedores y Cat_Proveedores
'   CREO                    : Julio Cruz
'   FECHA_CREO              : 5-Oct-2010
'   MODIFICO                :
'   FECHA_MODIFICO          :
'   CAUSA_MODIFICACIÓN      :
'*******************************************************************************
Private Sub Consulta_Facturas()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset            'Manejo de Registro
Dim Columnas As Integer                                         'Almacena el numero de Columna
Dim Renglon As Integer                                          'Almacena el numero de Renglon
Dim Rs_Consulta_Movimientos As rdoResultset                     'Manejo de Registro
Dim Rs_Consulta_Nota_Credito As rdoResultset                    'Manejo de Registro
Dim Rs_Consulta_Tipos_Pagos As rdoResultset                     'Manejo de Registro
Dim Columna As Integer, I As Integer

    'Consulta las facturas
    Txt_Totales.Text = ""
    Txt_Abonos.Text = ""
    Txt_Saldos.Text = ""
    Mi_SQL = "SELECT No_Factura,Fecha,Total, Saldo,Moneda,Cat_Proveedores.Nombre,Adm_Proveedores_Facturas.Tipo,Adm_Proveedores_Facturas.Proveedor_ID,Fecha_Recepcion"
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas,Cat_Proveedores "
    Mi_SQL = Mi_SQL & " WHERE Adm_Proveedores_Facturas.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    If Chk_Busqueda_No_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND No_Factura='" & Trim(Txt_Con_No_Factura.Text) & "'"
    If Chk_Busqueda_Proveedor.Value = 1 And Cmb_Con_Proveedor.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Proveedores_Facturas.Proveedor_ID='" & Format(Cmb_Con_Proveedor.ItemData(Cmb_Con_Proveedor.ListIndex), "00000") & "'"
    End If
    If Chk_Busqueda_Fecha_Recepcion.Value = 1 Then Mi_SQL = Mi_SQL & " AND Fecha_Recepcion >= " & Par_Fecha & Format(DTP_Fecha_Recepcion_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & " AND Fecha_Recepcion <= " & Par_Fecha & Format(DTP_Fecha_Recepcion_Final.Value, "MM/dd/yyyy") & Par_Fecha & ""
    If Chk_Busqueda_Fecha_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND Fecha >= " & Par_Fecha & Format(DTP_Fecha_Factura_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & " AND Fecha <= " & Par_Fecha & Format(DTP_Fecha_Factura_Final.Value, "MM/dd/yyyy") & Par_Fecha & ""
    If Chk_Busqueda_Estatus.Value = 1 And Cmb_Con_Estatus.Text = "Pagadas" Then Mi_SQL = Mi_SQL & " AND Pagada = 'S'"
    If Chk_Busqueda_Estatus.Value = 1 And Cmb_Con_Estatus.Text = "Sin Pagar" Then Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Recepcion"
    Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Facturas_Proveedores
        Grid_Consulta_Facturas.Rows = 0
        Grid_Consulta_Facturas.Cols = 9
        Columnas = 10
        Grid_Consulta_Facturas.AddItem "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "Nombre" & Chr(9) & "Tipo Factura" & Chr(9) & "SubTotal" & Chr(9) & "Nota Credito" & Chr(9) & "Total" & Chr(9) & "Moneda" & Chr(9) & "Proveedor_ID"
        Renglon = 1
        While Not .EOF
            Grid_Consulta_Facturas.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Trim(.rdoColumns("No_Factura")) _
                & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Tipo") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") _
                & Chr(9) & "" & Chr(9) & Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & .rdoColumns("Moneda") _
                & Chr(9) & .rdoColumns("Proveedor_ID")
            Txt_Totales.Text = Val(Txt_Totales.Text) + .rdoColumns("Total")
            Txt_Saldos.Text = Val(Txt_Saldos.Text) + .rdoColumns("Saldo")
            Grid_Consulta_Facturas.FixedRows = 1
            'CONSULTA LA NOTA DE CREDITO SI APLICA
            Mi_SQL = "SELECT SUM(Importe) AS Cantidad"
            Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Notas_Credito"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("Proveedor_ID") & "'"
            Mi_SQL = Mi_SQL & " AND No_Factura = '" & Trim(Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura")) & "'"
            Set Rs_Consulta_Nota_Credito = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Nota_Credito
                If Not .EOF And Not IsNull(.rdoColumns("Cantidad")) Then
                    Grid_Consulta_Facturas.TextMatrix(Renglon, 5) = Format(.rdoColumns("Cantidad"), "#,###,##0.00")
                    Txt_Abonos.Text = Val(Txt_Abonos.Text) + .rdoColumns("Cantidad")
                End If
            End With
            Rs_Consulta_Nota_Credito.Close
           'CONSULTA LOS ANTICIPOS Y PAGOS DE LA FACTURA
            Mi_SQL = "SELECT No_Movimiento,Referencia,No_Cheque,Fecha,Cantidad,Tipo,Banco,Forma_Pago"
            Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
            Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Tipo='E'"
            Mi_SQL = Mi_SQL & " AND Estatus='A'"
            Mi_SQL = Mi_SQL & " AND Proveedor_Cliente='" & Rs_Consulta_Facturas_Proveedores.rdoColumns("Proveedor_ID") & "'"
            Mi_SQL = Mi_SQL & " AND No_Factura='" & Trim(Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura")) & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Movimientos
                While Not .EOF
                    If Grid_Consulta_Facturas.Cols <= Columnas Then
                        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 5
                        Columnas = Grid_Consulta_Facturas.Cols
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 5) = "Tipo"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "Referencia"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 3) = "Banco"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 2) = "Fecha"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 1) = "Abono"
                        Grid_Consulta_Facturas.ColWidth(Columnas - 4) = 1000
                        Grid_Consulta_Facturas.ColWidth(Columnas - 3) = 1000
                        Grid_Consulta_Facturas.ColWidth(Columnas - 2) = 1100
                        Grid_Consulta_Facturas.ColWidth(Columnas - 1) = 1000
                    Else
                        Columnas = Columnas + 5
                    End If
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 5) = .rdoColumns("Forma_Pago")
                    If .rdoColumns("Referencia") <> "" Then
                        Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = .rdoColumns("Referencia")
                    Else
                        Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = .rdoColumns("No_Cheque")
                    End If
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 3) = .rdoColumns("Banco")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 2) = Format(.rdoColumns("Fecha"), "dd/MMM/yyyy")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 1) = Format(.rdoColumns("Cantidad"), "#,###,##0.00")
                    Txt_Abonos.Text = Val(Txt_Abonos.Text) + .rdoColumns("Cantidad")
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Movimientos.Close
            
''            Mi_SQL = "SELECT Adm_Corte_Caja_Chica.No_Corte,Adm_Corte_Caja_Chica.Referencia,Adm_Corte_Caja_Chica.No_Cheque,Adm_Corte_Caja_Chica.Fecha,Adm_Corte_Caja_Chica.Forma_Pago,Cat_Bancos.Nombre AS Banco,Detalles_Corte_Caja_Chica.Cantidad"
''            Mi_SQL = Mi_SQL & " FROM Adm_Corte_Caja_Chica,Detalles_Corte_Caja_Chica,Cat_Bancos"
''            Mi_SQL = Mi_SQL & " WHERE Adm_Corte_Caja_Chica.No_Corte=Detalles_Corte_Caja_Chica.No_Corte"
''            Mi_SQL = Mi_SQL & " AND Adm_Corte_Caja_Chica.Banco_ID=Cat_Bancos.Banco_ID"
''            Mi_SQL = Mi_SQL & " AND Adm_Corte_Caja_Chica.Estatus='A'"
''            Mi_SQL = Mi_SQL & " AND Detalles_Corte_Caja_Chica.Proveedor_ID='" & Rs_Consulta_Facturas_Proveedores.rdoColumns("Proveedor_ID") & "'"
''            Mi_SQL = Mi_SQL & " AND Detalles_Corte_Caja_Chica.Referencia='" & Trim(Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura")) & "'"
''            Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''            With Rs_Consulta_Movimientos
''                While Not .EOF
''                    If Grid_Consulta_Facturas.Cols <= Columnas Then
''                        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 5
''                        Columnas = Grid_Consulta_Facturas.Cols
''                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 5) = "Tipo"
''                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "Referencia"
''                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 3) = "Banco"
''                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 2) = "Fecha"
''                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 1) = "Abono"
''                        Grid_Consulta_Facturas.ColWidth(Columnas - 4) = 1000
''                        Grid_Consulta_Facturas.ColWidth(Columnas - 3) = 1000
''                        Grid_Consulta_Facturas.ColWidth(Columnas - 2) = 1100
''                        Grid_Consulta_Facturas.ColWidth(Columnas - 1) = 1000
''                    Else
''                        Columnas = Columnas + 5
''                    End If
''                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 5) = .rdoColumns("Forma_Pago")
''                    If Trim(.rdoColumns("Referencia")) <> "" Then
''                        Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = .rdoColumns("Referencia")
''                    Else
''                        Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = .rdoColumns("No_Cheque")
''                    End If
''                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 3) = .rdoColumns("Banco")
''                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 2) = Format(.rdoColumns("Fecha"), "dd/MMM/yyyy")
''                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 1) = Format(.rdoColumns("Cantidad"), "#,###,##.00")
''                    Txt_Abonos.Text = Val(Txt_Abonos.Text) + .rdoColumns("Cantidad")
''                    .MoveNext
''                Wend
''            End With
''            Rs_Consulta_Movimientos.Close
            
            Columnas = 10
            Renglon = Renglon + 1
            .MoveNext
        Wend
        'Cambia los saldos al final
        Columna = Grid_Consulta_Facturas.Cols
        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 1
        Grid_Consulta_Facturas.TextMatrix(0, Columna) = "Saldo"
        For I = 1 To Grid_Consulta_Facturas.Rows - 1
            Grid_Consulta_Facturas.TextMatrix(I, Columna) = Grid_Consulta_Facturas.TextMatrix(I, 6)
            Grid_Consulta_Facturas.TextMatrix(I, 6) = Format(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Facturas.TextMatrix(I, 4), ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Facturas.TextMatrix(I, 5), ",")), "#,###,##0.00")
        Next I
        Grid_Consulta_Facturas.ColWidth(0) = 1000
        Grid_Consulta_Facturas.ColAlignment(0) = 3
        Grid_Consulta_Facturas.ColWidth(1) = 1000
        Grid_Consulta_Facturas.ColAlignment(1) = 3
        Grid_Consulta_Facturas.ColWidth(2) = 2500
        Grid_Consulta_Facturas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Consulta_Facturas.ColWidth(3) = 1000
        Grid_Consulta_Facturas.ColWidth(4) = 1000
        Grid_Consulta_Facturas.ColWidth(5) = 1000
        Grid_Consulta_Facturas.ColWidth(6) = 1000
        Grid_Consulta_Facturas.ColWidth(7) = 1000
        Grid_Consulta_Facturas.ColWidth(8) = 0
        Txt_Abonos.Text = Format(Val(Txt_Totales.Text) - Val(Txt_Saldos.Text), "#,###,##0.00")
        Txt_Totales.Text = Format(Txt_Totales.Text, "#,###,##0.00")
        Txt_Saldos.Text = Format(Txt_Saldos.Text, "#,###,##0.00")
    End With
    Rs_Consulta_Facturas_Proveedores.Close
End Sub
'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Consulta_Facturas_Vales()
'   DESCRIPCIÓN: Realiza una consulta a la tabla de Facturas
'   PARÁMETROS:Generales de Facturas_Proveedores y Cat_Proveedores
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO          :Rafael Muñoz
'   FECHA_MODIFICO    :28-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Private Sub Consulta_Facturas_Vales()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset
Dim Columnas As Integer
Dim Renglon As Integer
Dim Rs_Consulta_Movimientos As rdoResultset
Dim Rs_Consulta_Notas_Credito_Proveedores As rdoResultset
Dim RS_Consulta_Vales As rdoResultset
Dim I As Integer

    'Consulta las facturas
    Txt_Totales.Text = ""
    Txt_Abonos.Text = ""
    Txt_Saldos.Text = ""
    Mi_SQL = "SELECT No_Factura, Fecha, Total, Saldo, Moneda, Nombre_Corto, Adm_Facturas_Proveedores.Proveedor_ID "
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores, Cat_Proveedores "
    Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Proveedores.Proveedor_ID = Cat_Proveedores.Proveedor_ID "
    If Chk_Busqueda_No_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND No_Factura = '" & Txt_Con_No_Factura.Text & "'"
    If Chk_Busqueda_Proveedor.Value = 1 Then Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Proveedor_ID = '" & Format(Cmb_Con_Proveedor.ItemData(Cmb_Con_Proveedor.ListIndex), "00000") & "'"
    If Chk_Busqueda_Fecha_Recepcion.Value = 1 Then Mi_SQL = Mi_SQL & " AND Fecha_Recepcion >= '" & Format(DTP_Fecha_Recepcion_Inicial.Value, "MM/dd/yyyy") & "' AND Fecha_Recepcion <= '" & Format(DTP_Fecha_Recepcion_Final.Value, "MM/dd/yyyy") & "'"
    If Chk_Busqueda_Fecha_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND Fecha >= '" & Format(DTP_Fecha_Factura_Inicial.Value, "MM/dd/yyyy") & "' AND Fecha <= '" & Format(DTP_Fecha_Factura_Final.Value, "MM/dd/yyyy") & "'"
    If Chk_Busqueda_Estatus.Value = 1 And Cmb_Con_Estatus.Text = "Pagadas" Then Mi_SQL = Mi_SQL & " AND Pagada = 'S'"
    If Chk_Busqueda_Estatus.Value = 1 And Cmb_Con_Estatus.Text = "Sin Pagar" Then Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha "
    Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Facturas_Proveedores
        Grid_Consulta_Facturas.Rows = 0
        Grid_Consulta_Facturas.Cols = 10
        Columnas = 10
        Grid_Consulta_Facturas.AddItem "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "No. Vale" & Chr(9) & _
        "Cliente" & Chr(9) & "Proveedor" & Chr(9) & "SubTotal" & Chr(9) & "Nota Credito" & Chr(9) & "Total" & Chr(9) & _
        "Moneda" & Chr(9) & "Prov_ID"
        Renglon = 1
        While Not .EOF
            Grid_Consulta_Facturas.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & .rdoColumns("No_Factura") _
            & Chr(9) & "" & Chr(9) & "" & Chr(9) & .rdoColumns("Nombre_Corto") & Chr(9) & Format(.rdoColumns("Total"), "######.00") _
            & Chr(9) & "" & Chr(9) & .rdoColumns("Saldo") & Chr(9) & .rdoColumns("Moneda") & Chr(9) & .rdoColumns("Proveedor_ID")
            Txt_Totales.Text = Val(Txt_Totales.Text) + .rdoColumns("Total")
            Txt_Saldos.Text = Val(Txt_Saldos.Text) + .rdoColumns("Saldo")
            Grid_Consulta_Facturas.FixedRows = 1
            
            'CONSULTA EL VALE Y EL CLIENTE SI SE DESEA
                Mi_SQL = "SELECT No_Vale, Cat_Clientes.Nombre_Corto FROM Tra_Vales, Cat_Clientes "
                Mi_SQL = Mi_SQL & " WHERE Tra_Vales.Cliente_ID= Cat_Clientes.Cliente_ID "
                Mi_SQL = Mi_SQL & " AND No_Factura_Pemex = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura") & "'"
                Set RS_Consulta_Vales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With RS_Consulta_Vales
                If Not .EOF Then
                    Grid_Consulta_Facturas.TextMatrix(Renglon, 2) = .rdoColumns("No_Vale")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, 3) = .rdoColumns("Nombre_Corto")
                End If
            End With
            RS_Consulta_Vales.Close
             
             'CONSULTA LA NOTA DE CREDITO SI APLICA
                Mi_SQL = "SELECT SUM(Importe) as Cantidad"
                Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito_Proveedores "
                Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("Proveedor_ID") & "'"
                Mi_SQL = Mi_SQL & " AND No_Factura = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura") & "'"
                Set Rs_Consulta_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Notas_Credito_Proveedores
                If Not .EOF And Not IsNull(.rdoColumns("Cantidad")) Then
                    Grid_Consulta_Facturas.TextMatrix(Renglon, 4) = Format(.rdoColumns("Cantidad"), "#########.00")
                    Txt_Abonos.Text = Val(Txt_Abonos.Text) + .rdoColumns("Cantidad")
                End If
            End With
            Rs_Consulta_Notas_Credito_Proveedores.Close
            
            'CONSULTA LOS ANTICIPOS Y PAGOS DE LA FACTURA
                Mi_SQL = "SELECT No_Movimiento, Referencia, Fecha, Cantidad, Tipo_Movimiento, Banco "
                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Tipo = 'E' "
                Mi_SQL = Mi_SQL & " AND No_Factura = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("No_Factura") & "'"
                Mi_SQL = Mi_SQL & " AND Estatus = 'A'"
                Mi_SQL = Mi_SQL & " AND Proveedor_Cliente = '" & Rs_Consulta_Facturas_Proveedores.rdoColumns("Proveedor_ID") & "'"
                Mi_SQL = Mi_SQL & " ORDER BY Fecha "
                Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Movimientos
                While Not .EOF
                    If Grid_Consulta_Facturas.Cols <= Columnas Then
                        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 5
                        Columnas = Grid_Consulta_Facturas.Cols
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 5) = "Tipo"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "Referencia"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 3) = "Banco"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 2) = "Fecha"
                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 1) = "Abono"
                        Grid_Consulta_Facturas.ColWidth(Columnas - 4) = 1000
                        Grid_Consulta_Facturas.ColWidth(Columnas - 3) = 1000
                        Grid_Consulta_Facturas.ColWidth(Columnas - 2) = 1100
                        Grid_Consulta_Facturas.ColWidth(Columnas - 1) = 1000
                    Else
                        Columnas = Columnas + 5
                    End If
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 5) = .rdoColumns("Tipo_Movimiento")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = .rdoColumns("Referencia")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 3) = .rdoColumns("Banco")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 2) = Format(.rdoColumns("Fecha"), "dd/MMM/yyyy")
                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 1) = Format(.rdoColumns("Cantidad"), "###,###,###.00")
                    Txt_Abonos.Text = Val(Txt_Abonos.Text) + .rdoColumns("Cantidad")
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Movimientos.Close
            Columnas = 10
            Renglon = Renglon + 1
            .MoveNext
        Wend
        'Cambia los saldos al final
        Columnas = Grid_Consulta_Facturas.Cols
        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 1
        Grid_Consulta_Facturas.TextMatrix(0, Columnas) = "Saldo"
        For I = 1 To Grid_Consulta_Facturas.Rows - 1
            Grid_Consulta_Facturas.TextMatrix(I, Columnas) = Format(Grid_Consulta_Facturas.TextMatrix(I, 7), "###,###,###.00")
            Grid_Consulta_Facturas.TextMatrix(I, 7) = Format(Val(Grid_Consulta_Facturas.TextMatrix(I, 5)) - Val(Grid_Consulta_Facturas.TextMatrix(I, 6)), "###,###,###.00")
        Next I
        Grid_Consulta_Facturas.ColWidth(0) = 1000
        Grid_Consulta_Facturas.ColAlignment(0) = 3
        Grid_Consulta_Facturas.ColWidth(1) = 1000
        Grid_Consulta_Facturas.ColAlignment(1) = 3
        Grid_Consulta_Facturas.ColWidth(2) = 1000
        Grid_Consulta_Facturas.ColAlignment(2) = 3
        Grid_Consulta_Facturas.ColWidth(3) = 1700
        Grid_Consulta_Facturas.ColWidth(4) = 1700
        Grid_Consulta_Facturas.ColWidth(5) = 1000
        Grid_Consulta_Facturas.ColWidth(6) = 1000
        Grid_Consulta_Facturas.ColWidth(7) = 1000
        Grid_Consulta_Facturas.ColWidth(8) = 1000
        Grid_Consulta_Facturas.ColWidth(9) = 0
        
        Txt_Abonos.Text = Format(Val(Txt_Totales.Text) - Val(Txt_Saldos.Text), "###,###,###.00")
        Txt_Totales.Text = Format(Txt_Totales.Text, "###,###,###.00")
        Txt_Saldos.Text = Format(Txt_Saldos.Text, "###,###,###.00")
        
    End With
    Rs_Consulta_Movimientos.Close
End Sub


