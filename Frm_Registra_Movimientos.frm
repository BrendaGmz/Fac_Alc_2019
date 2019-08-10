VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Registra_Movimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE MOVIMIENTOS"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   9630
   Begin VB.PictureBox Pic_Registro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   9825
      TabIndex        =   32
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Btn_Cancelar 
         Caption         =   "Cancelar"
         Height          =   555
         Left            =   5440
         Picture         =   "Frm_Registra_Movimientos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "B"
         Top             =   7215
         UseMaskColor    =   -1  'True
         Width           =   1350
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
         Height          =   1815
         Left            =   120
         TabIndex        =   61
         Top             =   5280
         Width           =   6375
         Begin VB.TextBox Txt_Comentarios 
            Height          =   1455
            Left            =   120
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   240
            Width           =   6135
         End
      End
      Begin VB.Frame Fra_Totales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Totales"
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
         Height          =   1815
         Left            =   6600
         TabIndex        =   60
         Top             =   5280
         Width           =   2895
         Begin VB.TextBox Txt_Porcentaje_Descuento 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Top             =   840
            Width           =   510
         End
         Begin VB.TextBox Txt_Importe_Descuento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Txt_Subtotal_0 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox Txt_Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Txt_IVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox Txt_Subtotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "%"
            Height          =   195
            Left            =   1635
            TabIndex        =   68
            Top             =   885
            Width           =   120
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Descuento"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   885
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Subtotal 0"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   585
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1485
            Width           =   360
         End
         Begin VB.Label Lbl_IVA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "IVA"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   1185
            Width           =   255
         End
         Begin VB.Label Lbl_Subtotal 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   285
            Width           =   585
         End
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   8100
         Picture         =   "Frm_Registra_Movimientos.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7215
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   120
         Picture         =   "Frm_Registra_Movimientos.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "A"
         Top             =   7215
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Consultar 
         Caption         =   "Consultar"
         Height          =   555
         Left            =   2780
         Picture         =   "Frm_Registra_Movimientos.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "C"
         Top             =   7215
         Width           =   1350
      End
      Begin VB.Frame Fra_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
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
         Left            =   120
         TabIndex        =   59
         Top             =   2565
         Width           =   9375
         Begin VB.CommandButton Btn_Eliminar 
            Caption         =   "Borrar Detalle"
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   2355
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Detalles 
            Height          =   2055
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Conceptos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conceptos"
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
         Height          =   800
         Left            =   120
         TabIndex        =   52
         Top             =   1755
         Width           =   9375
         Begin VB.TextBox Txt_Porcentaje_Impuesto 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5730
            TabIndex        =   15
            Top             =   420
            Width           =   570
         End
         Begin VB.CommandButton Btn_Agregar 
            Caption         =   "Agregar"
            Height          =   285
            Left            =   8520
            TabIndex        =   18
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox Txt_Importe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7415
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   420
            Width           =   1095
         End
         Begin VB.TextBox Txt_Precio_Unitario 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6310
            TabIndex        =   16
            Top             =   420
            Width           =   1095
         End
         Begin VB.ComboBox Cmb_Conceptos 
            Height          =   315
            Left            =   1865
            TabIndex        =   14
            Top             =   420
            Width           =   3855
         End
         Begin VB.TextBox Txt_Cantidad 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox Txt_Unidad 
            Height          =   285
            Left            =   880
            TabIndex        =   13
            Top             =   420
            Width           =   950
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unidad"
            Height          =   195
            Left            =   1140
            TabIndex        =   65
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Importe"
            Height          =   195
            Left            =   7680
            TabIndex        =   57
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "% Imp."
            Height          =   195
            Left            =   5750
            TabIndex        =   56
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "P. Unitario"
            Height          =   195
            Left            =   6525
            TabIndex        =   55
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Concepto"
            Height          =   195
            Left            =   3480
            TabIndex        =   54
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame Fra_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos Generales"
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
         Height          =   1635
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   9375
         Begin VB.TextBox Txt_Dias_Credito 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   920
            Width           =   1455
         End
         Begin VB.ComboBox Cmb_Forma_Pago 
            Height          =   315
            ItemData        =   "Frm_Registra_Movimientos.frx":0450
            Left            =   4440
            List            =   "Frm_Registra_Movimientos.frx":045A
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   920
            Width           =   1575
         End
         Begin VB.TextBox Txt_No_Factura 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   920
            Width           =   1575
         End
         Begin VB.ComboBox Cmb_Tipo 
            Height          =   315
            ItemData        =   "Frm_Registra_Movimientos.frx":0470
            Left            =   1320
            List            =   "Frm_Registra_Movimientos.frx":047A
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   565
            Width           =   1575
         End
         Begin VB.ComboBox Cmb_Clientes 
            Height          =   315
            ItemData        =   "Frm_Registra_Movimientos.frx":048F
            Left            =   4440
            List            =   "Frm_Registra_Movimientos.frx":0491
            TabIndex        =   5
            Top             =   560
            Width           =   4815
         End
         Begin VB.TextBox Txt_Estatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Txt_No_Movimiento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Txt_Clientes 
            Height          =   315
            Left            =   4440
            TabIndex        =   4
            Top             =   560
            Width           =   4695
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Emision 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   1275
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   89456643
            CurrentDate     =   41034
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Vencimiento 
            Height          =   315
            Left            =   7800
            TabIndex        =   10
            Top             =   1275
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   89456643
            CurrentDate     =   41034
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha 
            Height          =   315
            Left            =   7800
            TabIndex        =   2
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   89456643
            CurrentDate     =   41034
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   3300
            TabIndex        =   78
            Top             =   975
            Width           =   855
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Vencimiento"
            Height          =   195
            Left            =   6360
            TabIndex        =   77
            Top             =   1335
            Width           =   1365
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Días de Crédito"
            Height          =   195
            Left            =   6360
            TabIndex        =   76
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Factura"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   980
            Width           =   795
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Emisión"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   1335
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Movimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   620
            Width           =   1170
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre"
            Height          =   195
            Left            =   3300
            TabIndex        =   51
            Top             =   620
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   3300
            TabIndex        =   50
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Registro"
            Height          =   195
            Left            =   6360
            TabIndex        =   49
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Movimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   285
            Width           =   1110
         End
      End
   End
   Begin VB.PictureBox Pic_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   9825
      TabIndex        =   69
      Top             =   0
      Width           =   9855
      Begin VB.Frame Fra_Consulta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultado Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         TabIndex        =   72
         Top             =   1920
         Width           =   9375
         Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
            Height          =   5415
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   9551
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Filtros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filtros de Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   9375
         Begin VB.CheckBox Chk_Movimiento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Movimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   555
            Left            =   5160
            Picture         =   "Frm_Registra_Movimientos.frx":0493
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "C"
            Top             =   1080
            Width           =   1350
         End
         Begin VB.CommandButton Btn_Regresar 
            Caption         =   "Regresar"
            Height          =   555
            Left            =   7245
            Picture         =   "Frm_Registra_Movimientos.frx":05DD
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.TextBox Txt_Consulta_Movimiento 
            Height          =   315
            Left            =   1820
            TabIndex        =   39
            Top             =   660
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox Cmb_Consulta_Proveedor 
            Height          =   315
            Left            =   5160
            TabIndex        =   37
            Top             =   300
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.ComboBox Cmb_Consulta_Cliente 
            Height          =   315
            Left            =   5160
            TabIndex        =   35
            Top             =   300
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.ComboBox Cmb_Consulta_Estatus 
            Height          =   315
            ItemData        =   "Frm_Registra_Movimientos.frx":06DF
            Left            =   1820
            List            =   "Frm_Registra_Movimientos.frx":06E9
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1020
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox Chk_Fecha 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   3960
            TabIndex        =   40
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Chk_Cliente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   3960
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox Chk_Estatus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   975
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_inicio 
            Height          =   285
            Left            =   5160
            TabIndex        =   41
            Top             =   660
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   89456643
            CurrentDate     =   41034
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Final 
            Height          =   285
            Left            =   7440
            TabIndex        =   42
            Top             =   660
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   89456643
            CurrentDate     =   41034
         End
         Begin VB.CheckBox Chk_Proveedor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   3960
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox Cmb_Consulta_Tipo 
            Height          =   315
            ItemData        =   "Frm_Registra_Movimientos.frx":06FE
            Left            =   1820
            List            =   "Frm_Registra_Movimientos.frx":0708
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo de Movimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label Lbl_Al 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   7140
            TabIndex        =   71
            Top             =   705
            Visible         =   0   'False
            Width           =   135
         End
      End
   End
End
Attribute VB_Name = "Frm_Registra_Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Muestra_Movimientos()
Dim Rs_Consulta As rdoResultset
Dim Rs_Consulta_Detalles As rdoResultset
Dim Concepto_ID As String

    Grid_Detalles.Rows = 0
    Mi_SQL = "SELECT * FROM Ope_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(Grid_Consulta.TextMatrix(Grid_Consulta.RowSel, 0)) & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta.EOF Then
            With Rs_Consulta
                Btn_Regresar_Click
                Txt_No_Movimiento.Text = Trim(.rdoColumns("No_Movimiento"))
                Txt_Estatus.Text = Trim(.rdoColumns("Estatus"))
                If Txt_Estatus.Text = "ACTIVO" Then
                    Txt_Estatus.BackColor = &HC00000
                    Btn_Cancelar.Enabled = True
                Else
                    Txt_Estatus.BackColor = &HFF&
                    Btn_Cancelar.Enabled = False
                End If
                Dtp_Fecha.Value = .rdoColumns("Fecha")
                Dtp_Fecha_Emision.Value = .rdoColumns("Fecha_Emision")
                Dtp_Fecha_Vencimiento.Value = .rdoColumns("Fecha_Vencimiento")
                Txt_No_Factura.Text = Trim(.rdoColumns("Factura"))
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Tipo"), Cmb_Tipo)
                If Cmb_Tipo.Text = "INGRESO" Then
                    Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Cliente_ID"), Cmb_Clientes)
                Else
                    Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Proveedor_ID"), Cmb_Clientes)
                End If
                Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Forma_Pago")), Cmb_Forma_Pago)
                Txt_Dias_Credito.Text = Val(Txt_Dias_Credito.Text)
                Txt_Subtotal.Text = Format(.rdoColumns("Subtotal"), "###,##0.00")
                Txt_Subtotal_0.Text = Format(.rdoColumns("Subtotal_0"), "###,##0.00")
                If Val(.rdoColumns("Porcentaje_Descuento")) > 0 Then
                    Txt_Porcentaje_Descuento.Text = Format(.rdoColumns("Porcentaje_Descuento"), "###,##0.00")
                    Txt_Importe_Descuento.Text = Format(.rdoColumns("Descuento"), "###,##0.00")
                Else
                    Txt_Importe_Descuento.Text = "0.00"
                End If
                Txt_Total.Text = Format(.rdoColumns("Total"), "###,##0.00")
                Txt_Comentarios.Text = Trim(.rdoColumns("Comentarios"))
            End With
            
            Mi_SQL = "SELECT * FROM Ope_Movimientos_Detalles"
            Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(Grid_Consulta.TextMatrix(Grid_Consulta.RowSel, 0)) & "'"
            Set Rs_Consulta_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Detalles.EOF Then
                    If Grid_Detalles.Rows = 0 Then
                        Grid_Detalles.Cols = 7
                        Grid_Detalles.AddItem "Cantidad" & Chr(9) & "Unidad" & Chr(9) & "Concepto_ID" & Chr(9) & "Concepto" & Chr(9) _
                            & "% Imp." & Chr(9) & "P. Unitario" & Chr(9) & "Importe"
                    End If
                    With Rs_Consulta_Detalles
                        While Not .EOF
                            If Not IsNull(.rdoColumns("Concepto_ID")) Then
                                Concepto_ID = Trim(.rdoColumns("Concepto_ID"))
                            Else
                                Concepto_ID = ""
                            End If
                            Grid_Detalles.AddItem Format(.rdoColumns("Cantidad"), "#,##0.00") & Chr(9) _
                                & Trim(.rdoColumns("Unidad")) & Chr(9) & Concepto_ID & Chr(9) _
                                & Trim(.rdoColumns("Concepto")) & Chr(9) & Format(.rdoColumns("Tasa_IVA"), "##0.00") & Chr(9) _
                                & Format(.rdoColumns("Precio_Unitario"), "###,##0.00") & Chr(9) _
                                & Format(.rdoColumns("Importe"), "###,##0.00")
                            .MoveNext
                        Wend
                    End With
                    Formatea_Columnas_Grid
                End If
            Rs_Consulta_Detalles.Close
        End If
    Rs_Consulta.Close
End Sub


Private Sub Btn_Agregar_Click()
    If Val(Txt_Cantidad.Text) = 0 Then
        MsgBox "Favor de ingresar la cantidad", vbExclamation
        Txt_Cantidad.SetFocus
        Exit Sub
    End If
    If Cmb_Conceptos.Text = "" Then
        MsgBox "Favor de ingresar el concepto", vbExclamation
        Cmb_Conceptos.SetFocus
        Exit Sub
    End If
    If (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Porcentaje_Impuesto.Text, ",")) <> 0 And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Porcentaje_Impuesto.Text, ",")) <> 16) Then
        MsgBox "Ingrese una tasa de impuesto válida", vbExclamation
        Txt_Porcentaje_Impuesto.SetFocus
        Exit Sub
    End If
    If Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio_Unitario.Text, ",")) = 0 Then
        MsgBox "Favor de ingresar el precio unitario", vbExclamation
        Txt_Precio_Unitario.SetFocus
        Exit Sub
    End If
    If Grid_Detalles.Rows = 0 Then
        Grid_Detalles.Cols = 7
        Grid_Detalles.AddItem "Cantidad" & Chr(9) & "Unidad" & Chr(9) & "Concepto_ID" & Chr(9) & "Concepto" & Chr(9) _
            & "% Imp." & Chr(9) & "P. Unitario" & Chr(9) & "Importe"
    End If
    If Cmb_Conceptos.ListIndex > -1 Then
        Grid_Detalles.AddItem Format(Txt_Cantidad.Text, "#,##0.00") & Chr(9) & Trim(Txt_Unidad.Text) & Chr(9) _
            & Format(Cmb_Conceptos.ItemData(Cmb_Conceptos.ListIndex), "00000") & Chr(9) & Trim(Cmb_Conceptos.Text) & Chr(9) _
            & Format(Txt_Porcentaje_Impuesto.Text, "##0.00") & Chr(9) & Format(Txt_Precio_Unitario.Text, "###,##0.00") & Chr(9) _
            & Trim(Txt_Importe.Text)
    Else
        Grid_Detalles.AddItem Format(Txt_Cantidad.Text, "#,##0.00") & Chr(9) & Trim(Txt_Unidad.Text) & Chr(9) & "" & Chr(9) _
            & Trim(Cmb_Conceptos.Text) & Chr(9) _
            & Format(Txt_Porcentaje_Impuesto.Text, "##0.00") & Chr(9) & Format(Txt_Precio_Unitario.Text, "###,##0.00") & Chr(9) _
            & Trim(Txt_Importe.Text)
    End If
    Formatea_Columnas_Grid
    Calcula_Totales
    Txt_Cantidad.Text = ""
    Txt_Unidad.Text = ""
    Cmb_Conceptos.ListIndex = -1
    Cmb_Conceptos.Text = ""
    Txt_Precio_Unitario.Text = ""
    Txt_Importe.Text = ""
    Txt_Porcentaje_Impuesto.Text = Tasa_IVA
    Txt_Cantidad.SetFocus
End Sub

Private Sub Formatea_Columnas_Grid()
    If Grid_Detalles.Rows > 1 Then
        Grid_Detalles.FixedRows = 1
        Grid_Detalles.ColWidth(0) = 800     'Cantidad
        Grid_Detalles.ColWidth(1) = 950     'unidad
        Grid_Detalles.ColWidth(2) = 0       'Concepto_ID
        Grid_Detalles.ColWidth(3) = 4000    'Concepto
        Grid_Detalles.ColWidth(4) = 600     '% Imp
        Grid_Detalles.ColWidth(5) = 1200    'Precio unitario
        Grid_Detalles.ColWidth(6) = 1200    'importe
    Else
        Grid_Detalles.Rows = 0
    End If
    
    If Grid_Consulta.Rows > 1 Then
        Grid_Consulta.FixedRows = 1
        Grid_Consulta.ColWidth(0) = 1300    'No_Movimiento
        Grid_Consulta.ColAlignment(0) = 3
        Grid_Consulta.ColWidth(1) = 1050    'Tipo
        Grid_Consulta.ColAlignment(1) = 3
        Grid_Consulta.ColWidth(2) = 1200    'Fecha
        Grid_Consulta.ColAlignment(2) = 3
        Grid_Consulta.ColWidth(3) = 3950    'Cliente / Proveedor
        Grid_Consulta.ColAlignment(3) = 1
        Grid_Consulta.ColWidth(4) = 1200    'Estatus
        Grid_Consulta.ColAlignment(4) = 3
    Else
        Grid_Consulta.Rows = 0
    End If
End Sub

Private Sub Calcula_Totales()
Dim Subtotal As Double
Dim Subtotal_0 As Double
Dim Impuesto As Double
Dim Fila As Integer

    Subtotal = 0
    Subtotal_0 = 0
    Impuesto = 0
    For Fila = 1 To Grid_Detalles.Rows - 1
        If Val(Grid_Detalles.TextMatrix(Fila, 4)) = 16 Then
            Subtotal = Subtotal + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 6), ","))
        Else
            Subtotal_0 = Subtotal_0 + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 6), ","))
        End If
    Next
    Txt_Subtotal.Text = Format(Subtotal, "###,##0.00")
    Txt_Subtotal_0.Text = Format(Subtotal_0, "###,##0.00")
    Impuesto = Subtotal * (Tasa_IVA / 100)
    Txt_IVA.Text = Format(Impuesto, "###,##0.00")
    Txt_Total.Text = Format(Subtotal + Subtotal_0 + Impuesto, "###,##0.00")
    If Val(Txt_Porcentaje_Descuento.Text) > 0 Then Txt_Porcentaje_Descuento_Change
End Sub

Private Sub Btn_Buscar_Click()
    If Cmb_Consulta_Tipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de movimiento a consultar", vbExclamation
        Cmb_Consulta_Tipo.SetFocus
        Exit Sub
    End If
    If Chk_Movimiento.Value = 1 Then
        If Val(Txt_Consulta_Movimiento.Text) = 0 Then
            MsgBox "Ingrese el número de movimiento a consultar", vbExclamation
            Txt_Consulta_Movimiento.SetFocus
            Exit Sub
        End If
    End If
    If Chk_Estatus.Value = 1 Then
        If Cmb_Consulta_Estatus.ListIndex = -1 Then
            MsgBox "Seleccione el estatus a consultar", vbExclamation
            Cmb_Consulta_Estatus.SetFocus
            Exit Sub
        End If
    End If
    If Chk_Cliente.Value = 1 Then
        If Cmb_Consulta_Cliente.ListIndex = -1 Then
            MsgBox "Seleccione el cliente a consultar", vbExclamation
            Cmb_Consulta_Cliente.SetFocus
            Exit Sub
        End If
    End If
    If Chk_Proveedor.Value = 1 Then
        If Cmb_Consulta_Proveedor.ListIndex = -1 Then
            MsgBox "Seleccione el proveedor a consultar", vbExclamation
            Cmb_Consulta_Proveedor.SetFocus
            Exit Sub
        End If
    End If
    Consulta_Movimientos
End Sub

Private Sub Btn_Cancelar_Click()
Dim Rs_Modifica As rdoResultset
Dim Motivo As String

On Error GoTo errorHandler
    If Txt_No_Movimiento.Text <> "" Then
        If Txt_Estatus.Text = "ACTIVO" Then
            If MsgBox("¿Esta seguro de cancelar el movimiento?", vbCritical + vbYesNo) = vbYes Then
                Motivo = InputBox("Proporcione el motivo de cancelación")
                If Trim(Motivo) <> "" Then
                    MDIFrm_Apl_Principal.MousePointer = 11
                    Conexion_Base.BeginTrans
                        'Actualiza el encabezado de la factura
                        Mi_SQL = "SELECT No_Movimiento, Estatus, Fecha_Cancelacion, Motivo_Cancelacion, Usuario_Cancelo"
                        Mi_SQL = Mi_SQL & " FROM Ope_Movimientos"
                        Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(Txt_No_Movimiento.Text) & "'"
                        Set Rs_Modifica = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            With Rs_Modifica
                                .Edit
                                    .rdoColumns("Estatus") = "CANCELADO"
                                    .rdoColumns("Fecha_Cancelacion") = Now
                                    .rdoColumns("Motivo_Cancelacion") = Motivo_Cancelacion
                                    .rdoColumns("Usuario_Cancelo") = Nombre_Usuario
                                .Update
                            End With
                        Rs_Modifica.Close
                    Conexion_Base.CommitTrans
                    MDIFrm_Apl_Principal.MousePointer = 0
                    MsgBox "Movimiento cancelado", vbInformation
                    Txt_Estatus.BackColor = &HFF&
                    Txt_Estatus.Text = "CANCELADO"
                    Btn_Cancelar.Enabled = False
                Else
                    MsgBox "Debe de proporcionar el motivo de la cancelación", vbExclamation
                End If
            End If
        End If
    Else
        MsgBox "Seleccione un movimiento para cancelar", vbExclamation
    End If
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    'Obtiene el error
    If Err.Number = 7777 Then
        MsgBox Err.Description
    Else
        For Each Rdo_Error In rdoErrors
            MsgBox Rdo_Error.Description
        Next
    End If
End Sub

Private Sub Btn_Consultar_Click()
    Pic_Registro.Visible = False
    Pic_Consulta.Visible = True
    Cmb_Consulta_Tipo.SetFocus
End Sub

Private Sub Btn_Eliminar_Click()
    If Grid_Detalles.Rows > 1 Then
        If Grid_Detalles.Rows = 2 Then
            Grid_Detalles.Rows = 0
        Else
            Grid_Detalles.RemoveItem Grid_Detalles.RowSel
        End If
        Calcula_Totales
    End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Grid_Detalles.Rows = 0
        Fra_Generales.Enabled = True
        Fra_Conceptos.Enabled = True
        Fra_Detalles.Enabled = True
        Fra_Totales.Enabled = True
        Fra_Comentarios.Enabled = True
        Txt_Porcentaje_Impuesto.Text = Tasa_IVA
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Consultar.Enabled = False
        Btn_Cancelar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Txt_Estatus.Text = "ACTIVO"
        Txt_Estatus.BackColor = &HC00000
        Dtp_Fecha.Value = Now
        Dtp_Fecha_Emision.Value = Now
        Dtp_Fecha_Vencimiento.Value = Now
        Cmb_Tipo.SetFocus
    Else
        If Cmb_Tipo.ListIndex = -1 Then
            MsgBox "Seleccione el tipo de movimiento", vbExclamation
            Cmb_Clientes.SetFocus
            Exit Sub
        End If
        If Cmb_Clientes.ListIndex = -1 Then
            MsgBox "Seleccione el nombre del cliente", vbExclamation
            Cmb_Clientes.SetFocus
            Exit Sub
        End If
        If Grid_Detalles.Rows = 0 Then
            MsgBox "Agregue las partidas del movimiento", vbExclamation
            Txt_Cantidad.SetFocus
            Exit Sub
        End If
        Alta_Movimiento
    End If
End Sub

Private Sub Alta_Movimiento()
Dim Rs_Alta As rdoResultset
Dim Rs_Alta_Detalles As rdoResultset
Dim Contador As Integer

On Error GoTo errorHandler
    MDIFrm_Apl_Principal.MousePointer = 11
    Conexion_Base.BeginTrans
        Txt_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Ope_Movimientos", "No_Movimiento"), "0000000000")
        'Registra el encabezado de la factura
        Set Rs_Alta = Conectar_Ayudante.Recordset_Agregar("Ope_Movimientos")
            With Rs_Alta
                .AddNew
                    .rdoColumns("No_Movimiento") = Txt_No_Movimiento.Text
                    .rdoColumns("Tipo") = Trim(Cmb_Tipo.Text)
                    If Trim(Cmb_Tipo.Text) = "INGRESO" Then
                        .rdoColumns("Cliente_ID") = Format(Cmb_Clientes.ItemData(Cmb_Clientes.ListIndex), "00000")
                    Else
                        .rdoColumns("Proveedor_ID") = Format(Cmb_Clientes.ItemData(Cmb_Clientes.ListIndex), "00000")
                    End If
                    .rdoColumns("Forma_Pago") = Trim(Cmb_Forma_Pago.Text)
                    .rdoColumns("Dias_Credito") = Val(Txt_Dias_Credito.Text)
                    .rdoColumns("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Emision") = Format(Dtp_Fecha_Emision.Value, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Vencimiento") = Format(Dtp_Fecha_Vencimiento.Value, "MM/dd/yyyy")
                    .rdoColumns("Factura") = Trim(Txt_No_Factura.Text)
                    .rdoColumns("Tasa_Impuesto") = Tasa_IVA
                    .rdoColumns("Estatus") = Trim(Txt_Estatus.Text)
                    .rdoColumns("Porcentaje_Descuento") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Porcentaje_Descuento.Text, ","))
                    .rdoColumns("Descuento") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Importe_Descuento.Text, ","))
                    .rdoColumns("SubTotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
                    .rdoColumns("SubTotal_0") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ","))
                    .rdoColumns("Impuestos") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ","))
                    .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
                    .rdoColumns("Abono") = 0
                    .rdoColumns("Saldo") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
                    .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios.Text))
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
        Rs_Alta.Close
        
        Set Rs_Alta_Detalles = Conectar_Ayudante.Recordset_Agregar("Ope_Movimientos_Detalles")
        For Fila = 1 To Grid_Detalles.Rows - 1
            With Rs_Alta_Detalles
                .AddNew
                    .rdoColumns("No_Movimiento") = Txt_No_Movimiento.Text
                    .rdoColumns("Cantidad") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 0), ","))
                    .rdoColumns("Unidad") = Trim(Grid_Detalles.TextMatrix(Fila, 1))
                    If Trim(Grid_Detalles.TextMatrix(Fila, 2)) = "" Then
                        .rdoColumns("Concepto_ID") = ""
                    Else
                        .rdoColumns("Concepto_ID") = Trim(Grid_Detalles.TextMatrix(Fila, 2))
                    End If
                    .rdoColumns("Concepto") = Trim(Grid_Detalles.TextMatrix(Fila, 3))
                    .rdoColumns("Tasa_IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 4), ","))
                    .rdoColumns("Precio_Unitario") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 5), ","))
                    .rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalles.TextMatrix(Fila, 6), ","))
                .Update
            End With
        Next
        Rs_Alta_Detalles.Close
    Conexion_Base.CommitTrans
    'Regresa los controles inicial
    Fra_Generales.Enabled = False
    Fra_Detalles.Enabled = False
    Fra_Conceptos.Enabled = False
    Fra_Totales.Enabled = False
    Fra_Comentarios.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Btn_Consultar.Enabled = True
    Btn_Cancelar.Enabled = True
    Btn_Eliminar.Enabled = True
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox "Movimiento registrado satisfactoriamente", vbInformation
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Consulta.Visible = False
    Pic_Registro.Visible = True
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Grid_Detalles.Rows = 0
        Fra_Generales.Enabled = False
        Fra_Conceptos.Enabled = False
        Fra_Detalles.Enabled = False
        Fra_Comentarios.Enabled = False
        Fra_Totales.Enabled = False
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Nuevo.Enabled = True
        Btn_Cancelar.Enabled = True
        Btn_Consultar.Enabled = True
        Btn_Salir.Caption = "Salir"
    End If
End Sub

Private Sub Chk_Cliente_Click()
    Cmb_Consulta_Cliente.ListIndex = -1
    Cmb_Consulta_Cliente.Visible = False
    If Chk_Cliente.Value = 1 Then
        Cmb_Consulta_Cliente.Visible = True
        Cmb_Consulta_Cliente.SetFocus
    End If
End Sub

Private Sub Chk_Estatus_Click()
    Cmb_Consulta_Estatus.ListIndex = -1
    Cmb_Consulta_Estatus.Visible = False
    If Chk_Estatus.Value = 1 Then
        Cmb_Consulta_Estatus.Visible = True
        Cmb_Consulta_Estatus.SetFocus
    End If
End Sub

Private Sub Chk_Fecha_Click()
    Dtp_Fecha_inicio.Visible = False
    Dtp_Fecha_Final.Visible = False
    Lbl_Al.Visible = False
    If Chk_Fecha.Value = 1 Then
        Dtp_Fecha_inicio.Visible = True
        Dtp_Fecha_Final.Visible = True
        Lbl_Al.Visible = True
        Dtp_Fecha_inicio.Value = Now
        Dtp_Fecha_Final.Value = Now
        Dtp_Fecha_inicio.SetFocus
    End If
End Sub

Private Sub Chk_Movimiento_Click()
    Txt_Consulta_Movimiento.Text = ""
    Txt_Consulta_Movimiento.Visible = False
    If Chk_Movimiento.Value = 1 Then
        Txt_Consulta_Movimiento.Visible = True
        Txt_Consulta_Movimiento.SetFocus
    End If
End Sub

Private Sub Chk_Proveedor_Click()
    Cmb_Consulta_Proveedor.ListIndex = -1
    Cmb_Consulta_Proveedor.Visible = False
    If Chk_Proveedor.Value = 1 Then
        Cmb_Consulta_Proveedor.Visible = True
        Cmb_Consulta_Proveedor.SetFocus
    End If
End Sub

Private Sub Cmb_Clientes_Change()
    Cmb_Forma_Pago.ListIndex = -1
    Txt_Dias_Credito.Text = ""
End Sub

Private Sub Cmb_Clientes_Click()
Dim Rs_Consulta As rdoResultset

    If Cmb_Clientes.ListIndex > -1 Then
        If Cmb_Tipo.ListIndex = 0 Then
            Mi_SQL = "SELECT Cliente_ID, Forma_Pago, Dias_Credito FROM Cat_Clientes"
        Else
            Mi_SQL = "SELECT Proveedor_ID, Forma_Pago, Dias_Credito FROM Cat_Proveedores"
        End If
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Clientes.ItemData(Cmb_Clientes.ListIndex), "00000") & "'"
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta.EOF Then
                With Rs_Consulta
                    If Not IsNull(.rdoColumns("Forma_Pago")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Forma_Pago"), Cmb_Forma_Pago)
                        Txt_Dias_Credito.Text = Val(.rdoColumns("Dias_Credito"))
                        Dtp_Fecha_Vencimiento.Value = DateAdd("d", Val(Txt_Dias_Credito.Text), Dtp_Fecha_Vencimiento.Value)
                    Else
                        Cmb_Forma_Pago.ListIndex = -1
                        Txt_Dias_Credito.Text = ""
                    End If
                End With
            End If
        Rs_Consulta.Close
    End If
End Sub

Private Sub Cmb_Consulta_Tipo_Click()
    If Cmb_Consulta_Tipo.ListIndex > -1 Then
        Chk_Proveedor.Visible = False
        Chk_Proveedor.Value = 0
        Chk_Cliente.Visible = False
        Chk_Cliente.Value = 0
        If Cmb_Consulta_Tipo.ListIndex = 0 Then
            Chk_Cliente.Visible = True
        Else
            Chk_Proveedor.Visible = True
        End If
    End If
End Sub

Private Sub Cmb_Tipo_Click()
    If Cmb_Tipo.ListIndex > -1 Then
        Cmb_Clientes.Clear
        If Cmb_Tipo.ListIndex = 0 Then
            Lbl_Nombre.Caption = "Cliente"
            Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID, Nombre_Corto", "Cat_Clientes WHERE Estatus = 'ACTIVO'", Cmb_Clientes, 0, "Nombre_Corto")
            Call Conectar_Ayudante.Llena_Combo_Item("Concepto_ID, Descripcion", "Cat_Conceptos WHERE Tipo = 'INGRESO' AND Estatus = 'ACTIVO'", Cmb_Conceptos, 0, "Descripcion")
        Else
            Lbl_Nombre.Caption = "Proveedor"
            Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID, Nombre_Corto", "Cat_Proveedores WHERE Estatus = 'ACTIVO'", Cmb_Clientes, 0, "Nombre_Corto")
            Call Conectar_Ayudante.Llena_Combo_Item("Concepto_ID, Descripcion", "Cat_Conceptos WHERE Tipo = 'EGRESO' AND Estatus = 'ACTIVO'", Cmb_Conceptos, 0, "Descripcion")
        End If
    Else
        Lbl_Nombre.Caption = "Nombre"
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 8310
    Me.Width = 9720
    Me.Top = 300
    Me.Left = 3000
    Pic_Consulta.Visible = False
    Pic_Registro.Visible = True
    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID, Nombre_Corto", "Cat_Clientes WHERE Estatus = 'ACTIVO'", Cmb_Consulta_Cliente, 0, "Nombre_Corto")
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID, Nombre_Corto", "Cat_Proveedores WHERE Estatus = 'ACTIVO'", Cmb_Consulta_Proveedor, 0, "Nombre_Corto")
End Sub

Private Sub Grid_Consulta_DblClick()
    If Grid_Consulta.Rows > 1 Then
        Muestra_Movimientos
    End If
End Sub

Private Sub Grid_Consulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grid_Consulta_DblClick
    End If
End Sub

Private Sub Txt_Cantidad_Change()
    If Val(Conectar_Ayudante.Quitar_Caracter(Txt_Cantidad.Text, ",")) > 0 Then
        If Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio_Unitario.Text, ",")) > 0 Then
            Txt_Importe.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Cantidad.Text, ",")) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio_Unitario.Text, ",")), "###,##0.00")
        Else
            Txt_Importe.Text = ""
        End If
    Else
        Txt_Importe.Text = ""
    End If
End Sub

Private Sub Txt_Cantidad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cantidad.Text, True)
End Sub

Private Sub Txt_Porcentaje_Descuento_Change()
Dim Descuento As Double
Dim Impuesto As Double
Dim Total As Double

    If Val(Txt_Total.Text) > 0 Then
        If Val(Txt_Porcentaje_Descuento.Text) > 0 Then
            If Val(Txt_Porcentaje_Descuento.Text) < 100 Then
                Descuento = (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ","))) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Porcentaje_Descuento.Text, ",")) / 100
                If Val(Txt_Subtotal.Text) > 0 Then
                    Impuesto = (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ",")) - Val(Descuento)) * (1 + (Tasa_IVA / 100))
                    Impuesto = Impuesto - (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ",")) - Val(Descuento))
                Else
                    Impuesto = Format(0, "###,###0.00")
                End If
                Txt_Total.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ",")) - Val(Descuento) + Val(Impuesto)
                Txt_Importe_Descuento.Text = Format(Descuento, "###,###0.00")
                Txt_IVA.Text = Format(Impuesto, "###,###0.00")
                Txt_Total.Text = Format(Txt_Total.Text, "###,###0.00")
            Else
                MsgBox "El porcentaje de descuento no puede ser mayor que 100%", vbExclamation
                Txt_Porcentaje_Descuento.SetFocus
            End If
        Else
            Txt_Importe_Descuento.Text = Format(0, "###,##0.00")
            If Val(Txt_Subtotal.Text) > 0 Then
                Impuesto = (Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ","))) * (1 + (Tasa_IVA / 100))
                Impuesto = Impuesto - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ","))
            Else
                Impuesto = Format(0, "###,###0.00")
            End If
            Txt_Total.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal_0.Text, ",")) + Val(Impuesto)
            Txt_IVA.Text = Format(Impuesto, "###,###0.00")
            Txt_Total.Text = Format(Txt_Total.Text, "###,###0.00")
        End If
    Else
        If Val(Txt_Porcentaje_Descuento.Text) > 0 Then
            MsgBox "Ingrese las partidas a facturar", vbExclamation
            Txt_Porcentaje_Descuento.Text = ""
            Txt_Cantidad.SetFocus
        End If
    End If
End Sub

Private Sub Txt_Porcentaje_Descuento_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Porcentaje_Descuento.Text, True)
End Sub

Private Sub Txt_Porcentaje_Impuesto_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Porcentaje_Impuesto.Text, True)
End Sub

Private Sub Txt_Precio_Unitario_Change()
    If Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio_Unitario.Text, ",")) > 0 Then
        If Val(Conectar_Ayudante.Quitar_Caracter(Txt_Cantidad.Text, ",")) > 0 Then
            Txt_Importe.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Cantidad.Text, ",")) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Precio_Unitario.Text, ",")), "###,##0.00")
        Else
            Txt_Importe.Text = ""
        End If
    Else
        Txt_Importe.Text = ""
    End If
End Sub

Private Sub Txt_Precio_Unitario_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Precio_Unitario.Text, True)
End Sub

Private Sub Consulta_Movimientos()
Dim Rs_Consulta As rdoResultset

    Grid_Consulta.Rows = 0
    If Cmb_Consulta_Tipo.ListIndex = 0 Then
        Mi_SQL = "SELECT Ope_Movimientos.No_Movimiento, Tipo, Fecha, Ope_Movimientos.Estatus, Cat_Clientes.Nombre_Corto AS Nombre"
        Mi_SQL = Mi_SQL & " FROM Ope_Movimientos, Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Tipo = '" & Trim(Cmb_Consulta_Tipo.Text) & "'"
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID = Ope_Movimientos.Cliente_ID"
    Else
        Mi_SQL = "SELECT Ope_Movimientos.No_Movimiento, Tipo, Fecha, Ope_Movimientos.Estatus, Cat_Proveedores.Nombre_Corto AS Nombre"
        Mi_SQL = Mi_SQL & " FROM Ope_Movimientos, Cat_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Tipo = '" & Trim(Cmb_Consulta_Tipo.Text) & "'"
        Mi_SQL = Mi_SQL & " AND Cat_Proveedores.Proveedor_ID = Ope_Movimientos.Proveedor_ID"
    End If
    If Chk_Movimiento.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND No_Movimiento = '" & Format(Txt_Consulta_Movimiento.Text, "0000000000") & "'"
    End If
    If Chk_Estatus.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Estatus = '" & Trim(Cmb_Consulta_Estatus.Text) & "'"
    End If
    If Chk_Cliente.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cliente_ID = '" & Format(Cmb_Consulta_Cliente.ItemData(Cmb_Consulta_Cliente.ListIndex), "00000") & "'"
    End If
    If Chk_Proveedor.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Proveedor_ID = '" & Format(Cmb_Consulta_Proveedor.ItemData(Cmb_Consulta_Proveedor.ListIndex), "00000") & "'"
    End If
    If Chk_Fecha.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Fecha >=" & Par_Fecha & Format(Dtp_Fecha_inicio.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Fecha <=" & Par_Fecha & Format(Dtp_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Tipo, No_Movimiento"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta.EOF Then
            Grid_Consulta.Cols = 5
            Grid_Consulta.AddItem "No_Movimiento" & Chr(9) & "Tipo" & Chr(9) & "Fecha" & Chr(9) & "Cliente / Proveedor" & Chr(9) & "Estatus"
            With Rs_Consulta
                While Not .EOF
                    Grid_Consulta.AddItem Trim(.rdoColumns("No_Movimiento")) & Chr(9) & Trim(.rdoColumns("Tipo")) & Chr(9) & Format(.rdoColumns("Fecha"), "dd MMM yyyy") & Chr(9) & Trim(.rdoColumns("Nombre")) & Chr(9) & Trim(.rdoColumns("Estatus"))
                    .MoveNext
                Wend
            End With
            Formatea_Columnas_Grid
        End If
    Rs_Consulta.Close
End Sub
