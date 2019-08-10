VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Alm_Entradas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ENTRADAS ALMACEN"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   12450
   Begin VB.PictureBox Pic_Entradas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   15
      ScaleHeight     =   1920
      ScaleWidth      =   12405
      TabIndex        =   31
      Top             =   15
      Width           =   12405
      Begin VB.Frame Fra_Datos_Documento 
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
         Height          =   1785
         Left            =   7140
         TabIndex        =   38
         Top             =   135
         Width           =   5175
         Begin VB.TextBox Txt_Comentarios_Documento 
            Height          =   435
            Left            =   1605
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1275
            Width           =   3405
         End
         Begin VB.TextBox Txt_Total_Factura 
            Height          =   285
            Left            =   1605
            TabIndex        =   6
            Top             =   315
            Visible         =   0   'False
            Width           =   3405
         End
         Begin VB.TextBox Txt_Numero_Factura 
            Height          =   285
            Left            =   1605
            TabIndex        =   5
            Top             =   315
            Width           =   3390
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Documento 
            Height          =   315
            Left            =   1605
            TabIndex        =   7
            Top             =   615
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   16842755
            CurrentDate     =   40444
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Recepcion_Factura 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   930
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy"
            Format          =   16842755
            CurrentDate     =   39136
            MaxDate         =   402133
            MinDate         =   2
         End
         Begin VB.Label Lbl_Recepcion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Recibe"
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
            Left            =   15
            TabIndex        =   56
            Top             =   990
            Width           =   1200
         End
         Begin VB.Label Lbl_Numero_Entrda_Documento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Numero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   15
            TabIndex        =   41
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Lbl_Fecha_Documento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   15
            TabIndex        =   40
            Top             =   630
            Width           =   1560
         End
         Begin VB.Label Lbl_Comentarios 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Comentarios"
            Height          =   255
            Left            =   60
            TabIndex        =   39
            Top             =   1365
            Width           =   975
         End
         Begin VB.Label Lbl_Total 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   57
            Top             =   330
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame Fra_Generales_Entradas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos de Entrada"
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
         Height          =   1785
         Left            =   120
         TabIndex        =   32
         Top             =   135
         Width           =   6945
         Begin VB.ComboBox Cmb_Proveedor_Entradas 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   630
            Width           =   5610
         End
         Begin VB.ComboBox Cmb_Tipo_Entrada 
            Height          =   315
            ItemData        =   "Frm_Ope_Entradas.frx":0000
            Left            =   1200
            List            =   "Frm_Ope_Entradas.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   975
            Width           =   2685
         End
         Begin VB.TextBox Txt_Comentarios_Datos_Entrada 
            Height          =   405
            Left            =   1215
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1305
            Width           =   5610
         End
         Begin VB.TextBox Txt_Numero_Entrada 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   270
            Width           =   2685
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Cotizacion 
            Height          =   315
            Left            =   5325
            TabIndex        =   3
            Top             =   975
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   16842755
            CurrentDate     =   40444
         End
         Begin VB.Label Lbl_Tipo_Entrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo "
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
            Left            =   45
            TabIndex        =   37
            Top             =   1035
            Width           =   450
         End
         Begin VB.Label Lbl_Observaciones 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   45
            TabIndex        =   36
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label Lbl_Fecha 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4560
            TabIndex        =   35
            Top             =   990
            Width           =   540
         End
         Begin VB.Label Lbl_Proveedor_Entradas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   45
            TabIndex        =   34
            Top             =   652
            Width           =   1710
         End
         Begin VB.Label Lbl_Numero_Entrda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            TabIndex        =   33
            Top             =   293
            Width           =   1125
         End
      End
   End
   Begin VB.PictureBox Pic_Detalles_Entrada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   165
      ScaleHeight     =   5040
      ScaleWidth      =   12255
      TabIndex        =   55
      Top             =   1935
      Width           =   12255
      Begin VB.Frame Fra_Datos_Producto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del Producto"
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
         Height          =   1980
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   12195
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
            Height          =   420
            Left            =   10845
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Ope_Entradas.frx":0051
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1485
            Width           =   1260
         End
         Begin VB.TextBox Txt_Aplica_IVA 
            Height          =   285
            Left            =   2655
            TabIndex        =   78
            Top             =   225
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Txt_Cantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   195
            Width           =   1410
         End
         Begin VB.ComboBox Cmb_Descripcion 
            Height          =   315
            Left            =   2625
            TabIndex        =   11
            Top             =   195
            Width           =   9465
         End
         Begin VB.Frame Fra_Entrada 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Entrada"
            Height          =   945
            Left            =   6885
            TabIndex        =   69
            Top             =   495
            Width           =   5235
            Begin VB.CheckBox Chk_Caducidad 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Caducidad"
               Height          =   195
               Left            =   2655
               TabIndex        =   18
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox Txt_Numero_Lote 
               Height          =   315
               Left            =   1245
               MaxLength       =   20
               TabIndex        =   17
               Top             =   585
               Width           =   1365
            End
            Begin VB.ComboBox Cmb_Rack 
               Height          =   315
               ItemData        =   "Frm_Ope_Entradas.frx":3307
               Left            =   3255
               List            =   "Frm_Ope_Entradas.frx":3309
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   1635
               Visible         =   0   'False
               Width           =   1830
            End
            Begin VB.ComboBox Cmb_Nivel 
               Height          =   315
               ItemData        =   "Frm_Ope_Entradas.frx":330B
               Left            =   7245
               List            =   "Frm_Ope_Entradas.frx":330D
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   1635
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.ComboBox Cmb_Fila 
               Height          =   315
               ItemData        =   "Frm_Ope_Entradas.frx":330F
               Left            =   6030
               List            =   "Frm_Ope_Entradas.frx":3311
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   1635
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.TextBox Txt_Existencia 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   225
               Width           =   1320
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Caducidad 
               Height          =   315
               Left            =   3795
               TabIndex        =   19
               Top             =   540
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd MMM yyyy"
               Format          =   16842755
               CurrentDate     =   39694
               MaxDate         =   402133
               MinDate         =   2
            End
            Begin VB.Label Lbl_Lote 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "No. Lote"
               Height          =   195
               Left            =   135
               TabIndex        =   77
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lbl_Rack 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rack"
               Height          =   195
               Left            =   2370
               TabIndex        =   76
               Top             =   1695
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.Label Lbl_Nivel 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel"
               Height          =   195
               Left            =   6675
               TabIndex        =   75
               Top             =   1695
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label Lbl_Fila 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fondo"
               Height          =   195
               Left            =   5190
               TabIndex        =   74
               Top             =   1695
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Existencia"
               Height          =   195
               Index           =   35
               Left            =   135
               TabIndex        =   73
               Top             =   285
               Width           =   720
            End
         End
         Begin VB.Frame Fra_Costos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Costos"
            Height          =   945
            Left            =   135
            TabIndex        =   64
            Top             =   495
            Width           =   6720
            Begin VB.TextBox Txt_Costo_Sin_IVA 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1035
               TabIndex        =   12
               Top             =   180
               Width           =   2580
            End
            Begin VB.TextBox Txt_Importe_Costo_Con_Iva 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4230
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   548
               Width           =   2400
            End
            Begin VB.TextBox Txt_Costo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   540
               Width           =   2535
            End
            Begin VB.TextBox Txt_IVA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4230
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   180
               Width           =   2400
            End
            Begin VB.Label Lbl_Costo_Sin_IVA 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Costo Unit."
               Height          =   195
               Left            =   180
               TabIndex        =   68
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Lbl_Subtotal_Producto 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Importe"
               Height          =   195
               Left            =   3645
               TabIndex        =   67
               Top             =   615
               Width           =   525
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Costo"
               Height          =   195
               Index           =   27
               Left            =   180
               TabIndex        =   66
               Top             =   608
               Width           =   405
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "IVA"
               Height          =   195
               Index           =   24
               Left            =   3690
               TabIndex        =   65
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   22
            Left            =   150
            TabIndex        =   79
            Top             =   255
            Width           =   630
         End
      End
      Begin VB.Frame Fra_Detalles_Entrada 
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
         Height          =   2025
         Left            =   15
         TabIndex        =   58
         Top             =   1965
         Width           =   12165
         Begin VB.Frame Fra_Totales 
            BackColor       =   &H00FFFFFF&
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
            Height          =   1335
            Left            =   10080
            TabIndex        =   59
            Top             =   135
            Width           =   2025
            Begin VB.TextBox Txt_Total_IVA 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   675
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   570
               Width           =   1275
            End
            Begin VB.TextBox Txt_Subtotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   675
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   240
               Width           =   1275
            End
            Begin VB.TextBox Txt_Total 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   675
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   930
               Width           =   1275
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "I.V.A."
               Height          =   195
               Index           =   53
               Left            =   45
               TabIndex        =   62
               Top             =   630
               Width           =   390
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Subtotal"
               Height          =   195
               Index           =   39
               Left            =   45
               TabIndex        =   61
               Top             =   300
               Width           =   585
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Total"
               Height          =   195
               Index           =   59
               Left            =   45
               TabIndex        =   60
               Top             =   990
               Width           =   360
            End
         End
         Begin VB.CommandButton Btn_Elimina_detalle 
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
            Height          =   450
            Left            =   90
            Picture         =   "Frm_Ope_Entradas.frx":3313
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1500
            Width           =   1260
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Entradas 
            Height          =   1245
            Left            =   105
            TabIndex        =   21
            Top             =   225
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
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
         Left            =   45
         Picture         =   "Frm_Ope_Entradas.frx":65C5
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "A"
         Top             =   4200
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
         Left            =   10800
         Picture         =   "Frm_Ope_Entradas.frx":9AFC
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "A"
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
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
         Left            =   2790
         Picture         =   "Frm_Ope_Entradas.frx":D1FB
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "A"
         Top             =   4200
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
         Left            =   8205
         Picture         =   "Frm_Ope_Entradas.frx":10892
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "C"
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Imprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reimprimir"
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
         Left            =   5520
         Picture         =   "Frm_Ope_Entradas.frx":13E1E
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "A"
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
   End
   Begin VB.PictureBox Pic_Busqueda 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   120
      ScaleHeight     =   5085
      ScaleWidth      =   12390
      TabIndex        =   42
      Top             =   1950
      Width           =   12390
      Begin VB.Frame Fra_Busqueda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5040
         Left            =   60
         TabIndex        =   43
         Top             =   -15
         Width           =   12150
         Begin VB.CommandButton Btn_Consultar_Entrada 
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
            Left            =   9120
            Picture         =   "Frm_Ope_Entradas.frx":172E4
            Style           =   1  'Graphical
            TabIndex        =   54
            Tag             =   "C"
            Top             =   630
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.CommandButton Btn_Crystal 
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
            Left            =   10665
            Picture         =   "Frm_Ope_Entradas.frx":1A870
            Style           =   1  'Graphical
            TabIndex        =   53
            Tag             =   "A"
            Top             =   630
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.TextBox Txt_Busqueda_No_Entrada 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   48
            Top             =   225
            Width           =   2790
         End
         Begin VB.CheckBox Chk_No_Entrada 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Entrada"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   285
            Width           =   1230
         End
         Begin VB.CheckBox Chk_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   1050
            Width           =   870
         End
         Begin VB.ComboBox Cmb_Busqueda_Proveedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            TabIndex        =   45
            Top             =   607
            Width           =   7410
         End
         Begin VB.CheckBox Chk_Proveedor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   180
            TabIndex        =   44
            Top             =   667
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker Dtp_Busqueda_Fecha_Inicial 
            Height          =   315
            Left            =   1530
            TabIndex        =   49
            Top             =   990
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   16842755
            CurrentDate     =   37074
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Busqueda 
            Height          =   3495
            Left            =   90
            TabIndex        =   50
            Top             =   1380
            Width           =   11925
            _ExtentX        =   21034
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker Dtp_Busqueda_Fecha_Final 
            Height          =   315
            Left            =   6150
            TabIndex        =   51
            Top             =   990
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   16842755
            CurrentDate     =   37074
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Index           =   17
            Left            =   5250
            TabIndex        =   52
            Top             =   1035
            Width           =   135
         End
      End
   End
End
Attribute VB_Name = "Frm_Alm_Entradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Entrada_Almacen
'DESCRIPCION            : Da de alta la entrada a almacen
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 21-Sep-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Public Sub Entrada_Almacen()
Dim Rs_Ope_Entrada As rdoResultset
Dim Rs_Ope_Entrada_Detalles As rdoResultset
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Pedido As rdoResultset
Dim Rs_Alta_Tmp_Facturas_Proveedores As rdoResultset
Dim Mi_SQL As String
Dim Cont_Fila As Integer
Dim No_Control As String
Dim Utilidad_Producto As Double

On Error GoTo Handler
    Conexion_Base.BeginTrans
    
    'SE REGISTRA LA FACTURA EN EL SISTEMA
     Set Rs_Alta_Tmp_Facturas_Proveedores = Conectar_Ayudante.Recordset_Agregar("Tmp_Proveedores_Facturas")
     With Rs_Alta_Tmp_Facturas_Proveedores
         .AddNew
              No_Control = Format(Conectar_Ayudante.Maximo_Catalogo("Tmp_Proveedores_Facturas", "No_Control"), "0000000000")
             .rdoColumns("No_Control") = No_Control
             .rdoColumns("No_Factura") = Format(Txt_Numero_Factura.Text, "0000000000")
             .rdoColumns("Proveedor_ID") = Trim(Format(Cmb_Proveedor_Entradas.ItemData(Cmb_Proveedor_Entradas.ListIndex), "00000"))
             .rdoColumns("Fecha_Recepcion") = Format(Dtp_Fecha_Documento.Value, "MM/dd/yyyy")
             .rdoColumns("Subtotal") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
             .rdoColumns("IVA") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total_IVA.Text, ","))
             .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
             .rdoColumns("Flete") = 0
             .rdoColumns("Total_Factura") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total_Factura.Text, ","))
             .rdoColumns("Cancelada") = "NO"
             .rdoColumns("Facturar") = "NO"
             .rdoColumns("Aplicada") = "NO"
             .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Documento.Text))
             .rdoColumns("Usuario_Creo") = Nombre_Usuario
             .rdoColumns("Fecha_Creo") = Now
         .Update
     End With
     Rs_Alta_Tmp_Facturas_Proveedores.Close
    
    
    'SE DAN DE ALTA LOS DATOS GENERALES DE LA ENTRADA
    Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas")
    With Rs_Ope_Entrada
        .AddNew
            .rdoColumns("No_Control") = No_Control
            .rdoColumns("Entrada_ID") = Format(Txt_Numero_Entrada.Text, "0000000000")
            .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor_Entradas.ItemData(Cmb_Proveedor_Entradas.ListIndex), "00000")
            .rdoColumns("Fecha_Factura") = Format(Dtp_Fecha_Cotizacion.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Recepcion_Factura") = Format(Dtp_Fecha_Recepcion_Factura.Value, "MM/dd/yyyy")
            .rdoColumns("Tipo_Entrada") = Trim(Cmb_Tipo_Entrada.Text)
            .rdoColumns("Costo_Total") = (Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
            .rdoColumns("IVA") = (Conectar_Ayudante.Quitar_Caracter(Txt_Total_IVA.Text, ","))
            .rdoColumns("Total") = (Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
            If Trim(Cmb_Tipo_Entrada.Text) = "COMPRA" Then
                .rdoColumns("Estatus") = "RECEPCION"
            Else
                If Trim(Cmb_Tipo_Entrada.Text) = "PRODUCCIÓN" Then
                    .rdoColumns("Estatus") = "PRODUCCIÓN"
                End If
                If Trim(Cmb_Tipo_Entrada.Text) = "AJUSTE" Then
                    .rdoColumns("Estatus") = "AJUSTE"
                End If
                If Trim(Cmb_Tipo_Entrada.Text) = "TRASPASO" Then
                    .rdoColumns("Estatus") = "TRASPASO"
                End If
                If Trim(Cmb_Tipo_Entrada.Text) = "INVENTARIO INICIAL" Then
                    .rdoColumns("Estatus") = "INVENTARIO INICIAL"
                End If
            End If
            .rdoColumns("Observaciones") = UCase(Txt_Comentarios_Datos_Entrada.Text)
            .rdoColumns("Usuario_Creo") = Trim(Nombre_Usuario)
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Ope_Entrada.Close
    
    'SE DAN DE ALTA LOS DETALLES DE LA ENTRADA
    For Cont_Fila = 1 To Grid_Entradas.Rows - 1 Step 1
        Set Rs_Ope_Entrada_Detalles = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas_Detalles")
        With Rs_Ope_Entrada_Detalles
            .AddNew
                Rs_Ope_Entrada_Detalles.rdoColumns("Entrada_ID") = Format(Txt_Numero_Entrada.Text, "0000000000")
                Rs_Ope_Entrada_Detalles.rdoColumns("Producto_ID") = Grid_Entradas.TextMatrix(Cont_Fila, 0)
                Rs_Ope_Entrada_Detalles.rdoColumns("Descripcion") = Grid_Entradas.TextMatrix(Cont_Fila, 1)
                Rs_Ope_Entrada_Detalles.rdoColumns("Costo") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Cont_Fila, 3), ","))
                Rs_Ope_Entrada_Detalles.rdoColumns("Importe") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Cont_Fila, 4), ","))
                Rs_Ope_Entrada_Detalles.rdoColumns("Cantidad") = Val(Grid_Entradas.TextMatrix(Cont_Fila, 2))
                Rs_Ope_Entrada_Detalles.rdoColumns("No_Lote") = Grid_Entradas.TextMatrix(Cont_Fila, 5)
                If Grid_Entradas.TextMatrix(Cont_Fila, 6) <> "" Then Rs_Ope_Entrada_Detalles.rdoColumns("Fecha_Caducidad") = Format(Grid_Entradas.TextMatrix(Cont_Fila, 6), "MM/dd/yyyy")
                Rs_Ope_Entrada_Detalles.rdoColumns("Faltante") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Cont_Fila, 2), ","))
                Rs_Ope_Entrada_Detalles.rdoColumns("Impuesto") = Grid_Entradas.TextMatrix(Cont_Fila, 7)
                Rs_Ope_Entrada_Detalles.rdoColumns("IVA") = Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Cont_Fila, 8), ",")
                If Trim(Cmb_Tipo_Entrada.Text) = "COMPRA" Then
                    Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "RECEPCION"
                Else
                    If Trim(Cmb_Tipo_Entrada.Text) = "PRODUCCIÓN" Then
                        Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "PRODUCCIÓN"
                    End If
                    If Trim(Cmb_Tipo_Entrada.Text) = "AJUSTE" Then
                        Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "AJUSTE"
                    End If
                    If Trim(Cmb_Tipo_Entrada.Text) = "TRASPASO" Then
                        Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "TRASPASO"
                    End If
                    If Trim(Cmb_Tipo_Entrada.Text) = "INVENTARIO INICIAL" Then
                        .rdoColumns("Estatus") = "INVENTARIO INICIAL"
                    End If
                End If
            .Update
        End With
        Rs_Ope_Entrada_Detalles.Close
        'ACTUALIZA EXISTENCIA EN EL CATALOGO DE PRODUCTOS
        Mi_SQL = " SELECT * FROM Cat_Productos "
        Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Grid_Entradas.TextMatrix(Cont_Fila, 0) & "' "
        Set Rs_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Cat_Productos
            .Edit
                If Not IsNull(.rdoColumns("Existencia")) Then
                    .rdoColumns("Existencia") = Val(.rdoColumns("Existencia")) + Val(Grid_Entradas.TextMatrix(Cont_Fila, 2))
                Else
                    .rdoColumns("Existencia") = Val(Grid_Entradas.TextMatrix(Cont_Fila, 2))
                End If
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
        Rs_Cat_Productos.Close
    Next Cont_Fila
    Conexion_Base.CommitTrans
    MsgBox "Entrada dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Buscar.Enabled = True
    Fra_Generales_Entradas.Enabled = False
    Fra_Datos_Documento.Enabled = False
    Fra_Datos_Producto.Enabled = False
    Fra_Detalles_Entrada.Enabled = False
    If MsgBox("¿Desea enviarla a impresión?", vbQuestion + vbYesNo) = vbYes Then
        Imprime_Entrada
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    If Err.Number = 482 Then
        MsgBox "Impresion Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        MsgBox Err.Description
    End If
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Agregar_Click
'DESCRIPCION            : Agrega los productos al grid Grid_Entradas
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 20-Sep-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Agregar_Click()
Dim Cont_Filas As Integer
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
Dim Año As Integer
Dim Mes As Integer
Dim Dia As Integer
Dim Fecha_Valida As Date
Dim Agregar_Partida As Boolean
Dim Producto_Lote As String
Dim Fecha_Cadicidad As String
Dim Costo_Sin_IVA As String

On Error GoTo Handler

    If Cmb_Descripcion.ListIndex > -1 Then
        If Cmb_Proveedor_Entradas.ListIndex > -1 Then
            If Not Txt_Cantidad.Text = "" Then
                'VALIDA SI LA FECHA DE CADUCIDAD ES MENOR O IGUAL A UN AÑO TOMANDO EN CUENTA LA FECHA DEL DIA
                Año = DatePart("yyyy", Now)
                Mes = DatePart("m", Now)
                Dia = DatePart("d", Now)
                Fecha_Valida = DateSerial(Año + 1, Mes, Dia)
                If Chk_Caducidad.Value = 1 And Dtp_Fecha_Caducidad.Value <= Now Then
                    If MsgBox("La fecha de caducidad  es menor o igual a la fecha actual ¿Desea agregar el producto?", vbQuestion + vbYesNo) = vbYes Then
                        Agregar_Partida = True
                    Else
                        Agregar_Partida = False
                    End If
                Else
                    If Chk_Caducidad.Value = 1 And Dtp_Fecha_Caducidad.Value <= Fecha_Valida Then
                        If MsgBox("La fecha de caducidad es menor o igual a un año ¿Desea agregar el producto?", vbQuestion + vbYesNo) = vbYes Then
                            Agregar_Partida = True
                        Else
                            Agregar_Partida = False
                        End If
                    Else
                        Agregar_Partida = True
                    End If
                End If
                
                'SE VALIDA SE LA PARTIDA SE VA A AGREGAR
                If Agregar_Partida = True Then
                    If Grid_Entradas.Rows < 1 Then
                        'Llena el Grid
                        Grid_Entradas.Rows = 0
                        Grid_Entradas.Cols = 9
                        'Se agrega el encabezado
                        Grid_Entradas.AddItem "Producto ID" & Chr(9) & "Descripción" & Chr(9) & _
                        "Cantidad" & Chr(9) & "Costo" & Chr(9) & "Importe" & Chr(9) & "No. Lote" _
                        & Chr(9) & "Caducidad" & Chr(9) & "Impuesto" & Chr(9) & "IVA"
                    End If
                    For Cont_Filas = 0 To Grid_Entradas.Rows - 1 Step 1
                        If Format(Grid_Entradas.TextMatrix(Cont_Filas, 0), "00000") = Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") Then
                            MsgBox "El Producto ya esta en la lista", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                            Exit Sub
                        End If
                    Next
                    If Chk_Caducidad.Value = 1 Then
                        Fecha_Cadicidad = Dtp_Fecha_Caducidad.Value
                    Else
                        Fecha_Cadicidad = ""
                    End If
                    If Trim(Txt_Aplica_IVA.Text) = "SI" Then
                        'Se agregan los registros al grid
                        Costo_Sin_IVA = Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")
                        Grid_Entradas.AddItem Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & Cmb_Descripcion.Text & Chr(9) _
                        & Txt_Cantidad.Text & Chr(9) _
                        & Format(Costo_Sin_IVA, "#,##0.00") & Chr(9) & Format(Val(Costo_Sin_IVA) * Val(Txt_Cantidad.Text), "#,##0.00") & Chr(9) & Txt_Numero_Lote.Text & Chr(9) & Format(Fecha_Cadicidad, "MM/dd/yyyy") & Chr(9) & PG_Retencion_IVA & Chr(9) _
                        & Format((Val(Costo_Sin_IVA) * Val(Txt_Cantidad.Text)) * (Val(PG_Retencion_IVA)), "#,##0.00")
                    Else
                        'Se agregan los registros al grid
                        Costo_Sin_IVA = Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")
                        Grid_Entradas.AddItem Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & Cmb_Descripcion.Text & Chr(9) _
                        & Txt_Cantidad.Text & Chr(9) _
                        & Format(Costo_Sin_IVA, "#,##0.00") & Chr(9) & Format(Val(Costo_Sin_IVA) * Val(Txt_Cantidad.Text), "#,##0.00") & Chr(9) & Txt_Numero_Lote.Text & Chr(9) & Format(Fecha_Cadicidad, "MM/dd/yyyy") & Chr(9) & PG_Retencion_IVA & Chr(9) _
                        & ""
                    End If
                    'Grid Partidas
                    If Grid_Entradas.Rows > 1 Then
                        Grid_Entradas.FixedRows = 1
                        Grid_Entradas.FixedCols = 1
                        Grid_Entradas.ColWidth(0) = 1000       'Producto_ID
                        Grid_Entradas.ColAlignment(0) = 3
                        Grid_Entradas.FixedRows = 1
                        Grid_Entradas.ColWidth(1) = 4000    'Descripcion
                        Grid_Entradas.ColAlignment(1) = 1
                        Grid_Entradas.ColWidth(2) = 750    'Cantidad
                        Grid_Entradas.ColAlignment(2) = 1
                        Grid_Entradas.ColWidth(3) = 1000    'Costo
                        Grid_Entradas.ColAlignment(3) = 3
                        Grid_Entradas.ColWidth(4) = 1000    'Importe
                        Grid_Entradas.ColAlignment(4) = 3
                        Grid_Entradas.ColWidth(5) = 800    'Numero Lote
                        Grid_Entradas.ColAlignment(5) = 3
                        Grid_Entradas.ColWidth(6) = 1000    'Fecha Caducidad
                        Grid_Entradas.ColAlignment(6) = 3
                        Grid_Entradas.ColWidth(7) = 0    'Impuesto
                        Grid_Entradas.ColWidth(8) = 0    'IVA
                        'Pone el setfocus en la primera fila del Grid
                        With Grid_Entradas
                            .Col = 1
                            .Row = 1
                            .ColSel = .Cols - 1
                            .RowSel = 1
                            .TopRow = .Row
                            .SetFocus
                        End With
                    End If
                    Call Calcula_Totales
                    Cmb_Descripcion.Text = ""
                Else
                    Exit Sub
                End If
            Else
                Txt_Cantidad.SetFocus
            End If
        Else
            MsgBox "Seleccione un Proveedor", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
            Cmb_Proveedor_Entradas.SetFocus
        End If
    Else
        Cmb_Descripcion.SetFocus
    End If
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub Btn_Buscar_Click()
    Cmb_Proveedor_Entradas.ListIndex = -1
    Cmb_Descripcion_KeyPress (13)
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Pic_Busqueda.Visible = True
    Pic_Detalles_Entrada.Visible = False
    Fra_Busqueda.Visible = True
    Grid_Entradas.Rows = 0
    Cmb_Tipo_Entrada.ListIndex = -1
    Chk_No_Entrada.Value = 0
    Chk_No_Entrada_Click
    Chk_Proveedor.Value = 0
    Chk_Proveedor_Click
    Chk_Fechas.Value = 0
    Chk_Fechas_Click
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Cancelar_Click
'DESCRIPCION            : Cancela la Entrada a almacen y resta de la existencia
'                         del producto la cantidad cancelada.
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 30-Sep-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Cancelar_Click()
Dim Mi_SQL As String
Dim Rs_Cancela_Entrada As rdoResultset         'Manejo del registro
Dim Rs_Modifica_Cat_Productos As rdoResultset
Dim Fila As Integer
    
On Error GoTo Handler:
        If MsgBox("¿Está seguro de cancelar la entrada?", vbQuestion + vbYesNo) = vbYes Then
            Conexion_Base.BeginTrans
            'SE CONSULTA EN Alm_Entradas LA ENTRADA QUE SE VA ACANCELAR
            Mi_SQL = "SELECT * FROM Alm_Entradas"
            Mi_SQL = Mi_SQL & " WHERE Entrada_ID ='" & Trim(Txt_Numero_Entrada.Text) & "'"
            Set Rs_Cancela_Entrada = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            'SE ASIGNAN LO VALORES EN LA TABLA PARA CANCELAR
            If Not Rs_Cancela_Entrada.EOF Then
                With Rs_Cancela_Entrada
                    .Edit
                        .rdoColumns("Estatus") = "CANCELADA"
                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                        .rdoColumns("Fecha_Modifico") = Now
                    .Update
                End With
            End If
            Rs_Cancela_Entrada.Close
            'CONSULTA LOS DETALLES DE LA ENTRADA PARA CANCELARLA Y DESCONTARLOS DEL INVENTARIO
            Mi_SQL = "SELECT * FROM Alm_Entradas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Entrada_ID ='" & Trim(Txt_Numero_Entrada.Text) & "'"
            Set Rs_Cancela_Entrada = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            While Not Rs_Cancela_Entrada.EOF
                If Rs_Cancela_Entrada.rdoColumns("Estatus") = "RECEPCION" Then
                    'CONSULTA LOS DATOS DEL PRODUCTO: SU EXISTENCIA EN ALMACEN Y SE RESTA LA CANTIDAD COLICITADA DEL CLIENTE
                    Mi_SQL = "SELECT Producto_ID,Existencia FROM Cat_Productos"
                    Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Rs_Cancela_Entrada.rdoColumns("Producto_ID") & "'"
                    Set Rs_Modifica_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Cat_Productos.EOF Then
                        Rs_Modifica_Cat_Productos.Edit
                            Rs_Modifica_Cat_Productos.rdoColumns("Existencia") = Val(Rs_Modifica_Cat_Productos.rdoColumns("Existencia")) - Val(Rs_Cancela_Entrada.rdoColumns("Faltante"))
                        Rs_Modifica_Cat_Productos.Update
                    End If
                    Rs_Modifica_Cat_Productos.Close
                    Rs_Cancela_Entrada.Edit
                        Rs_Cancela_Entrada.rdoColumns("Estatus") = "RECHAZADO"
                        Rs_Cancela_Entrada.rdoColumns("Faltante") = 0
                    Rs_Cancela_Entrada.Update
                Else
                    MsgBox "Parte de la entrada ya ha sido facturada por lo que no se permite ya su cancelación", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    Conexion_Base.RollbackTrans
                    Exit Sub
                End If
                Rs_Cancela_Entrada.MoveNext
            Wend
            Rs_Cancela_Entrada.Close
            MsgBox "La entrada " & Txt_Numero_Entrada.Text & " ha sido cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            Btn_Cancelar.Enabled = False
            Conexion_Base.CommitTrans
        End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Consultar_Entrada_Click()
    If Chk_Proveedor.Value = 1 And Cmb_Busqueda_Proveedor.Text = "" Then
        MsgBox "Seleccione el Proveedor de la Lista", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        Cmb_Busqueda_Proveedor.SetFocus
        Exit Sub
    End If
    If Chk_No_Entrada.Value = 1 And Trim(Txt_Busqueda_No_Entrada.Text) = "" Then
        MsgBox "Escriba el No. de Entrada", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        Txt_Busqueda_No_Entrada.SetFocus
    End If
    Call Consulta_Entradas
End Sub

Private Sub Btn_Crystal_Click()
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Pic_Busqueda.Visible = False
    Pic_Detalles_Entrada.Visible = True
    Grid_Entradas.Rows = 0
    Btn_Buscar.Enabled = True
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Elimina_detalle_Click
'DESCRIPCION            : Elimina la partida seleccionada del grid Grid_Entrdas
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 20-Sep-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Elimina_detalle_Click()
    If Grid_Entradas.Rows > 1 Then
        If Grid_Entradas.Rows = 2 Then
            Grid_Entradas.Rows = 0
        Else
            Grid_Entradas.RemoveItem Grid_Entradas.RowSel
        End If
        Call Calcula_Totales
    End If
End Sub

Private Sub Btn_Imprimir_Click()

On Error GoTo Handler

    If Txt_Numero_Entrada.Text <> "" Then
        Imprime_Entrada
    Else
        MsgBox "Seleccione una entrada para reimprimir", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    Exit Sub
Handler:
    If Err.Number = 482 Then
        MsgBox "Impresion Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Dtp_Fecha_Caducidad.Value = Now
        Btn_Salir.Caption = "Cancelar"
        Btn_Nuevo.Caption = "Dar de alta"
        Btn_Cancelar.Enabled = False
        Btn_Buscar.Enabled = False
        Grid_Entradas.Rows = 0
        Cmb_Proveedor_Entradas.Text = ""
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Txt_Numero_Entrada.Text = Conectar_Ayudante.Maximo_Catalogo("Alm_Entradas", "Entrada_ID")
        Fra_Generales_Entradas.Enabled = True
        Fra_Datos_Documento.Enabled = True
        Fra_Datos_Producto.Enabled = True
        Fra_Detalles_Entrada.Enabled = True
        Dtp_Fecha_Cotizacion.Value = Now
        Dtp_Fecha_Documento.Value = Now
        Dtp_Fecha_Recepcion_Factura.Value = Now
        Cmb_Proveedor_Entradas.SetFocus
    Else
        If Trim(Cmb_Tipo_Entrada.Text) = "COMPRA" Then
            If Not Txt_Total_Factura.Text = "" Then
                If Not Txt_Numero_Factura.Text = "" Then
                    If Cmb_Proveedor_Entradas.ListIndex > -1 Then
                        If Not Cmb_Tipo_Entrada.Text = "" Then
                            If Grid_Entradas.Rows > 1 Then
                                Call Entrada_Almacen
                            Else
                                MsgBox "Debe haber por lo menos una partida agregada", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                            End If
                        Else
                            MsgBox "Seleccione el tipo de entrada", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                            Cmb_Tipo_Entrada.SetFocus
                        End If
                    Else
                        MsgBox "Seleccione un proveedor", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                        Cmb_Proveedor_Entradas.SetFocus
                    End If
                Else
                    MsgBox "Falta el numero de Factura", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    Txt_Numero_Factura.SetFocus
                End If
            Else
                MsgBox " Falta el total de la Factura ", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Txt_Total_Factura.SetFocus
            End If
        Else
            If Not Txt_Total_Factura.Text = "" Then
                If Cmb_Proveedor_Entradas.ListIndex > -1 Then
                    If Not Cmb_Tipo_Entrada.Text = "" Then
                        If Grid_Entradas.Rows > 1 Then
                            Call Entrada_Almacen
                        Else
                            MsgBox "Debe haber por lo menos una partida agregada", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                        End If
                    Else
                        MsgBox "Seleccione el tipo de entrada", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                        Cmb_Tipo_Entrada.SetFocus
                    End If
                Else
                    MsgBox "Seleccione un proveedor", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                    Cmb_Proveedor_Entradas.SetFocus
                End If
            Else
                MsgBox " Falta el total de la Factura ", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
                Txt_Total_Factura.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else 'Cancelar
        Cmb_Proveedor_Entradas.ListIndex = -1
        Cmb_Descripcion_KeyPress (13)
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Fra_Generales_Entradas.Enabled = False
        Fra_Datos_Documento.Enabled = False
        Fra_Datos_Producto.Enabled = False
        Fra_Detalles_Entrada.Enabled = False
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Caption = "Nuevo"
        Grid_Entradas.Rows = 0
        Btn_Buscar.Enabled = True
    End If
End Sub

Private Sub Cmb_Estatus_Entradas_Change()

End Sub

Private Sub Chk_Caducidad_Click()
    If Chk_Caducidad.Value = 1 Then
        Dtp_Fecha_Caducidad.Enabled = True
    Else
        Dtp_Fecha_Caducidad.Enabled = False
    End If
End Sub

Private Sub Chk_Fechas_Click()
    Dtp_Busqueda_Fecha_Inicial.Enabled = Chk_Fechas.Value
    Dtp_Busqueda_Fecha_Inicial.Value = Now
    Dtp_Busqueda_Fecha_Final.Enabled = Chk_Fechas.Value
    Dtp_Busqueda_Fecha_Final.Value = Now
End Sub

Private Sub Chk_No_Entrada_Click()
    Txt_Busqueda_No_Entrada.Enabled = Chk_No_Entrada.Value
    Txt_Busqueda_No_Entrada.Text = ""
End Sub

Private Sub Chk_Proveedor_Click()
    Cmb_Busqueda_Proveedor.Enabled = Chk_Proveedor.Value
    Cmb_Busqueda_Proveedor.Text = ""
    Cmb_Busqueda_Proveedor.ListIndex = -1
End Sub
''*******************************************************************************
'NOMBRE DE LA FUNCION   : Cmb_Orden_Compra_Click
'DESCRIPCION            : Consulta los pedios relacionados con la orden de compra
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 10-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Cmb_Orden_Compra_Click()
Dim Rs_Consulta As rdoResultset
Dim Rs_Consulta_Cat_Productos As rdoResultset
Dim Rs_Consulta_Cat_Impuestos As rdoResultset
Dim Rs_Consulta_Referencia As rdoResultset
Dim Mi_SQL As String
Dim Impuesto As Double

On Error GoTo Handler
    Grid_Entradas.Rows = 0
    Grid_Detalles_Orden_Compra.Rows = 0
    Grid_Detalles_Orden_Compra.Cols = 11
    'SE CONSULTA LA REFERNCIA EN LOS DATOS GENERALES DEL PRODUCTO
    Mi_SQL = " SELECT Referencia FROM Ope_Pedidos WHERE Pedido_ID='" & Cmb_Orden_Compra.Text & "' "
    Set Rs_Consulta_Referencia = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Referencia.EOF Then
        Txt_Referencia.Text = Rs_Consulta_Referencia!Referencia
    End If
    Rs_Consulta_Referencia.Close
    'CONSULTA LOS DETALLES DE LA ORDEN DE COMPRA
    Mi_SQL = " SELECT * FROM  Ope_Pedidos_Detalles "
    Mi_SQL = Mi_SQL & " WHERE Pedido_ID='" & Cmb_Orden_Compra.Text & "'"
    Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Format(Cmb_Proveedor_Entradas.ItemData(Cmb_Proveedor_Entradas.ListIndex), "00000") & "'"  '"
    Mi_SQL = Mi_SQL & " AND Estatus ='PENDIENTE'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
        'Se agrega el encabezado
        Grid_Detalles_Orden_Compra.AddItem "Producto ID" & Chr(9) & "Clave" & Chr(9) & "Descripción" & Chr(9) & _
        "Marca" & Chr(9) & "Cantidad" & Chr(9) & "Costo" & Chr(9) & "Importe" & Chr(9) & "No. Lote" _
        & Chr(9) & "Fecha Caducidad" & Chr(9) & "Impuesto" & Chr(9) & "IVA"
        While Not Rs_Consulta.EOF
            'Consulta si en el producto se aplica iva
            Mi_SQL = " SELECT * FROM  Cat_productos "
            Mi_SQL = Mi_SQL & " WHERE Clave ='" & Trim(Rs_Consulta!Clave) & "'"
            Mi_SQL = Mi_SQL & " OR Producto_ID ='" & Trim(Rs_Consulta!Producto_ID) & "'"
            Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                'consulta el impuesto del producto
                Mi_SQL = " SELECT * FROM  Cat_Impuestos "
                Mi_SQL = Mi_SQL & " WHERE Impuesto_ID ='" & Trim(Rs_Consulta_Cat_Productos!Impuesto_ID) & "'"
                Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Impuestos.EOF Then
                    Impuesto = Val(Rs_Consulta_Cat_Impuestos!Impuesto)
                Else
                    Impuesto = 0
                End If
                    Grid_Detalles_Orden_Compra.AddItem Rs_Consulta_Cat_Productos!Producto_ID & Chr(9) & Rs_Consulta!Clave & Chr(9) & Rs_Consulta!Descripcion & Chr(9) & _
                    Rs_Consulta!Marca & Chr(9) & Rs_Consulta!Cantidad & Chr(9) & Rs_Consulta!Costo & Chr(9) & Rs_Consulta!Importe & Chr(9) & "" _
                    & Chr(9) & "" & Chr(9) & Impuesto & Chr(9) & Rs_Consulta_Cat_Productos!Aplica_IVA
                Rs_Consulta_Cat_Impuestos.Close
            Rs_Consulta_Cat_Productos.Close
            Rs_Consulta.MoveNext
        Wend
        If Grid_Detalles_Orden_Compra.Rows > 1 Then
            Grid_Detalles_Orden_Compra.FixedRows = 1
            Grid_Detalles_Orden_Compra.FixedCols = 2
            Grid_Detalles_Orden_Compra.ColWidth(0) = 0       'Producto_ID
            Grid_Detalles_Orden_Compra.ColAlignment(0) = 3
            Grid_Detalles_Orden_Compra.FixedRows = 1
            Grid_Detalles_Orden_Compra.ColWidth(1) = 1200    'Clave
            Grid_Detalles_Orden_Compra.ColAlignment(1) = 1
            Grid_Detalles_Orden_Compra.ColWidth(2) = 5249    'Descripcion
            Grid_Detalles_Orden_Compra.ColAlignment(2) = 1
            Grid_Detalles_Orden_Compra.ColWidth(3) = 1490    'Marca
            Grid_Detalles_Orden_Compra.ColAlignment(3) = 1
            Grid_Detalles_Orden_Compra.ColWidth(4) = 1200    'Cantidad
            Grid_Detalles_Orden_Compra.ColAlignment(4) = 3
            Grid_Detalles_Orden_Compra.ColWidth(5) = 1200    'Costo
            Grid_Detalles_Orden_Compra.ColAlignment(5) = 3
            Grid_Detalles_Orden_Compra.ColWidth(6) = 1200    'Importe
            Grid_Detalles_Orden_Compra.ColAlignment(6) = 3
            Grid_Detalles_Orden_Compra.ColWidth(7) = 0    'Numero Lote
            Grid_Detalles_Orden_Compra.ColWidth(8) = 0    'Fecha Caducidad
            Grid_Detalles_Orden_Compra.ColWidth(9) = 0    'Impuesto
            Grid_Detalles_Orden_Compra.ColWidth(10) = 0    'IVA
            'Pone el setfocus en la primera fila del Grid
            With Grid_Detalles_Orden_Compra
                .Col = 1
                .Row = 1
                .ColSel = .Cols - 1
                .RowSel = 1
                .TopRow = .Row
                .SetFocus
            End With
        End If
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub







Private Sub Cmb_Productos_Change()

End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Proveedor_Entradas_KeyPress
'DESCRIPCIÓN          : consulta los datos del proveedor
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 11-Enero-2011
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Proveedor_Entradas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor_Entradas, 1, "Nombre")
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Proveedor_Entradas.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Form_Load()
    Me.Height = 7425
    Me.Width = 12660
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor_Entradas, 1, "Estatus ='ACTIVO' AND Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Busqueda_Proveedor, 1, "Estatus ='ACTIVO' AND Nombre")
    Dtp_Fecha_Caducidad.Value = Now
    Call Cmb_Descripcion_KeyPress(13)
End Sub
''*******************************************************************************
'NOMBRE DE LA FUNCION   : Calcula_Totales
'DESCRIPCION            : Calcula los totales de las partidas que se ingresan al grid: Grid_Entradas
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 24-Sep-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Public Sub Calcula_Totales()
Dim Fila As Integer
Dim Total_IVA As Double

    Txt_Subtotal.Text = ""
    Txt_Total_IVA = ""
    Txt_Total.Text = ""
    Total_Costo_Sin_IVA = 0
    For Fila = 1 To Grid_Entradas.Rows - 1 Step 1
        Txt_Subtotal.Text = Val(Txt_Subtotal.Text) + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Fila, 4), ","))    'importes
        Total_IVA = Val(Total_IVA) + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Entradas.TextMatrix(Fila, 8), ","))   'Fila IVA
    Next
    Txt_Subtotal.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ","))
    Txt_Total_IVA.Text = Val(Total_IVA)
    Txt_Total.Text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Subtotal.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total_IVA.Text, ","))
    Txt_Total_Factura = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ","))
    Txt_Subtotal.Text = Format(Txt_Subtotal.Text, "#,##0.00")
    Txt_Total_IVA.Text = Format(Txt_Total_IVA.Text, "#,##0.00")
    Txt_Total.Text = Format(Val(Txt_Total.Text), "#,##0.00")
    Txt_Total_Factura = Format(Txt_Total_Factura.Text, "#,##0.00")
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Consulta_Entradas
'DESCRIPCIÓN                : Consulta las entradas de la tabla Alm_Entradas
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30-Septiembre-2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Consulta_Entradas()
Dim Mi_SQL As String
Dim Rs_Consulta_Entradas As rdoResultset
Dim Rs_Consulta_Nombre As rdoResultset
Dim Nombre As String

    Mi_SQL = " SELECT distinct (Alm_Entradas.Entrada_ID),Alm_Entradas.Costo_Total,Alm_Entradas.Fecha_Factura,"
    Mi_SQL = Mi_SQL & " Alm_Entradas.Tipo_Entrada as Tipo_Recepcion,Tmp_Proveedores_Facturas.No_Factura,Tmp_Proveedores_Facturas.Proveedor_ID"
    Mi_SQL = Mi_SQL & " FROM Alm_Entradas,Tmp_Proveedores_Facturas "
    Mi_SQL = Mi_SQL & " WHERE Alm_Entradas.No_Control = Tmp_Proveedores_Facturas.No_Control"
    ''Mi_SQL = Mi_SQL & " AND Alm_Entradas.Tipo_Entrada='RECEPCION'"
    If Chk_No_Entrada.Value = 1 And Trim(Txt_Busqueda_No_Entrada.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Entrada_ID LIKE '%" & Trim(Txt_Busqueda_No_Entrada.Text) & "%'"
    End If
    If Chk_Proveedor.Value = 1 And Cmb_Busqueda_Proveedor.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Proveedor_ID='" & Format(Cmb_Busqueda_Proveedor.ItemData(Cmb_Busqueda_Proveedor.ListIndex), "00000") & "'"
    End If
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Fecha_Factura BETWEEN " & Par_Fecha & Format(Dtp_Busqueda_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & " AND " & Par_Fecha & Format(Dtp_Busqueda_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas.Entrada_ID"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'LLena el grid con los datos del resultado de la busqueda
    If Not Rs_Consulta_Entradas.EOF Then
        Grid_Busqueda.Rows = 0
        Grid_Busqueda.AddItem "No. Entrada" & Chr(9) & "Nombre" & Chr(9) & "Tipo Entrada" & Chr(9) & "No. Factura" & Chr(9) & "Costo" & Chr(9) & "Fecha"
        With Rs_Consulta_Entradas
            While Not .EOF
                'Valida si es cliente o proveedor para consultar su nombre
                If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                    Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                    Set Rs_Consulta_Nombre = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Nombre.EOF Then
                        Nombre = Rs_Consulta_Nombre.rdoColumns("Nombre")
                        Rs_Consulta_Nombre.Close
                    Else
                        Nombre = ""
                    End If
                Else
                    Mi_SQL = ""
                    Nombre = ""
                End If
                'Se agregan los datos al grind
                Grid_Busqueda.AddItem .rdoColumns("Entrada_ID") & Chr(9) & Nombre _
                    & Chr(9) & .rdoColumns("Tipo_Recepcion") & Chr(9) & .rdoColumns("No_Factura") _
                    & Chr(9) & Format(.rdoColumns("Costo_Total"), "#,###,##0.00") & Chr(9) & Format(.rdoColumns("Fecha_Factura"), "dd/MMM/yyyy")
                Grid_Busqueda.FixedRows = 1
                Rs_Consulta_Entradas.MoveNext
            Wend
        End With
        Grid_Busqueda.ColWidth(0) = Grid_Busqueda.Width * 0.09   'No. Entrada
        Grid_Busqueda.ColAlignment(0) = flexAlignCenterCenter
        Grid_Busqueda.ColWidth(1) = Grid_Busqueda.Width * 0.4    'Nombre
        Grid_Busqueda.ColAlignment(1) = flexAlignLeftCenter
        Grid_Busqueda.ColWidth(2) = Grid_Busqueda.Width * 0.15   'Tipo Recepcion
        Grid_Busqueda.ColAlignment(2) = flexAlignLeftCenter
        Grid_Busqueda.ColWidth(3) = Grid_Busqueda.Width * 0.13   'No. Factura
        Grid_Busqueda.ColAlignment(3) = flexAlignCenterCenter
        Grid_Busqueda.ColWidth(4) = Grid_Busqueda.Width * 0.1    'Costo
        Grid_Busqueda.ColWidth(5) = Grid_Busqueda.Width * 0.1    'Fecha
        Grid_Busqueda.ColAlignment(5) = flexAlignCenterCenter
    Else 'Si no encontro registros con esos criterios,limpia el grid y manda mensaje al usuario
        Grid_Busqueda.Rows = 0
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Cmb_Busqueda_Proveedor.ListIndex = -1
        Dtp_Busqueda_Fecha_Inicial.Value = Now
        Dtp_Busqueda_Fecha_Final.Value = Now
        If Chk_No_Entrada.Value = 1 Or Chk_Proveedor.Value = 1 Or Chk_Fechas.Value = 1 Then
            MsgBox "No existen entradas bajo los criterios de búsqueda establecidos", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        Else
            MsgBox "No hay entradas registradas", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
        'Deshabilita todos los checkbox
        Chk_Proveedor.Value = 0
        Chk_No_Entrada.Value = 0
        Chk_Fechas.Value = 0
    End If
    Rs_Consulta_Entradas.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Grid_Busqueda_DblClick
'DESCRIPCIÓN                : Muestra los detalles de la entrada
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 30-Septiembre-2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Grid_Busqueda_DblClick()
Dim Mi_SQL As String
Dim Rs_Consulta_Entradas As rdoResultset
Dim Rs_Consulta_Entradas_Detalles As rdoResultset
Dim Rs_Consulta_Nombre As rdoResultset
Dim Nombre As String
    
    If Grid_Busqueda.Rows > 1 Then
        'Consulta para obtener los datos de la entrada
        Mi_SQL = "SELECT Alm_Entradas.*,Tmp_Proveedores_Facturas.*"
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas,Tmp_Proveedores_Facturas"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas.No_Control = Tmp_Proveedores_Facturas.No_Control"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas.Entrada_ID = '" & Grid_Busqueda.TextMatrix(Grid_Busqueda.RowSel, 0) & "'"
        Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Entradas.EOF Then
            With Rs_Consulta_Entradas
                Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                Set Rs_Consulta_Nombre = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Nombre.EOF Then
                    Nombre = Rs_Consulta_Nombre.rdoColumns("Nombre")
                Else
                    Nombre = ""
                End If
                Rs_Consulta_Nombre.Close
                Fra_Datos_Documento.Enabled = True
                Txt_Numero_Entrada.Text = .rdoColumns("Entrada_ID")
                If Not IsNull(.rdoColumns("Fecha_Factura")) Then Dtp_Fecha_Cotizacion.Value = Format(.rdoColumns("Fecha_Factura"), "dd/MMM/yyyy")
                If Not IsNull(.rdoColumns("Fecha_Recepcion_Factura")) Then Dtp_Fecha_Recepcion_Factura.Value = Format(.rdoColumns("Fecha_Recepcion_Factura"), "dd/MMM/yyyy")
                Cmb_Tipo_Entrada.Text = .rdoColumns("Tipo_Entrada")
                Cmb_Proveedor_Entradas.Text = Nombre
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Comentarios_Datos_Entrada.Text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("No_Factura")) Then Txt_Numero_Factura.Text = .rdoColumns("No_Factura")
                If Not IsNull(.rdoColumns("Fecha_Factura")) Then Dtp_Fecha_Documento.Value = Format(.rdoColumns("Fecha_Recepcion"), "dd/MMMM/yyyy")
                Txt_Total_Factura.Text = Format(.rdoColumns("Total_Factura"), "#0.00")
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Comentarios_Documento.Text = .rdoColumns("Comentarios")
                Txt_Subtotal.Text = Format(.rdoColumns("Subtotal"), "#0.00")
                Txt_Total_IVA.Text = Format(.rdoColumns("IVA"), "#0.00")
                Txt_Total.Text = Format(.rdoColumns("Total"), "#0.00")
            End With
            
            'Llena el Grid
            Grid_Entradas.Rows = 0
            Grid_Entradas.Cols = 9
            'Se agrega el encabezado
            Grid_Entradas.AddItem "Producto ID" & Chr(9) & "Descripción" & Chr(9) & _
            "Cantidad" & Chr(9) & "Costo" & Chr(9) & "Importe" & Chr(9) & "No. Lote" _
            & Chr(9) & "Fecha Caducidad" & Chr(9) & "Impuesto" & Chr(9) & "IVA"
            'Consultar los productos de la entrada para cargarlos al Grid_Detalles
            Mi_SQL = "SELECT Alm_Entradas_Detalles.* "
            Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Detalles "
            Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Detalles.Entrada_ID ='" & Grid_Busqueda.TextMatrix(Grid_Busqueda.RowSel, 0) & "'"
            Set Rs_Consulta_Entradas_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Entradas_Detalles.EOF Then
                With Rs_Consulta_Entradas_Detalles
                    While Not .EOF
                        'Se agregan los registros al grid
                        Grid_Entradas.AddItem Format(.rdoColumns("Producto_ID"), "00000") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) _
                        & .rdoColumns("Cantidad") & Chr(9) _
                        & .rdoColumns("Costo") & Chr(9) & Val(.rdoColumns("Costo")) * Val(.rdoColumns("Cantidad")) & Chr(9) & .rdoColumns("No_Lote") & Chr(9) & Format(.rdoColumns("Fecha_Caducidad").Value, "MM/dd/yyyy") & Chr(9) & .rdoColumns("Impuesto") & Chr(9) & .rdoColumns("IVA")
                        Grid_Entradas.FixedRows = 1
                        .MoveNext
                    Wend
                End With
            End If
            Rs_Consulta_Entradas_Detalles.Close
            If Grid_Entradas.Rows > 1 Then
                Grid_Entradas.FixedRows = 1
                Grid_Entradas.FixedCols = 1
                Grid_Entradas.ColWidth(0) = 1000       'Producto_ID
                Grid_Entradas.ColAlignment(0) = 3
                Grid_Entradas.FixedRows = 1
                Grid_Entradas.ColWidth(1) = 4000    'Descripcion
                Grid_Entradas.ColAlignment(1) = 1
                Grid_Entradas.ColWidth(2) = 750    'Cantidad
                Grid_Entradas.ColAlignment(2) = 1
                Grid_Entradas.ColWidth(3) = 1000    'Costo
                Grid_Entradas.ColAlignment(3) = 3
                Grid_Entradas.ColWidth(4) = 1000    'Importe
                Grid_Entradas.ColAlignment(4) = 3
                Grid_Entradas.ColWidth(5) = 800    'Numero Lote
                Grid_Entradas.ColAlignment(5) = 3
                Grid_Entradas.ColWidth(6) = 1000    'Fecha Caducidad
                Grid_Entradas.ColAlignment(6) = 3
                Grid_Entradas.ColWidth(7) = 0    'Impuesto
                Grid_Entradas.ColWidth(8) = 0    'IVA
                'Pone el setfocus en la primera fila del Grid
                With Grid_Entradas
                    .Col = 1
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                End With
            End If
            Btn_Buscar.Enabled = True
            Btn_Consultar_Entrada.Enabled = True
            Pic_Busqueda.Visible = False
            Pic_Detalles_Entrada.Visible = True
            Btn_Imprimir.Enabled = True
            'Valida el estatus de la entrada para ver si permitiria cancela
            If Rs_Consulta_Entradas.rdoColumns("Estatus") = "RECEPCION" Then
                Btn_Cancelar.Enabled = True
            Else
                Btn_Cancelar.Enabled = False
            End If
        End If
        Rs_Consulta_Entradas.Close
        Call Calcula_Totales
    End If
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Imprime_Entrada
'DESCRIPCIÓN                : Imprime la entrada solicitada por el usuario
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Octubre-2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************
Private Sub Imprime_Entrada()
Dim I As Integer
        
    Printer.FontSize = 8
    Printer.Font = "COURIER NEW"
    Printer.FontSize = 10
    Printer.FontSize = 8
    Printer.Print
    Printer.Print "******************************************************************************************************************************"
    Printer.Print
    Printer.Print Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("***********************************************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("***********************************************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("***********************************************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("***********************************************************************************************************************"))
    Printer.Print Spc(100); Format(Now, "dd/MMM/yyyy")
    Printer.Print Conectar_Ayudante.Alinea_Derecha("VALE DE ENTRADA DE ALMACEN", Len("***********************************************************************************************************************"))
    Printer.Print
    Printer.Print " No. Entrada   :  "; Txt_Numero_Entrada.Text
    Printer.Print " Proveedor     :  "; Cmb_Proveedor_Entradas.Text
    Printer.Print " Fecha Entrada :  "; Format(Dtp_Fecha_Documento.Value, "dd/MMM/yyyy")
    Printer.Print " No. Factura   :  "; Txt_Numero_Factura.Text
    Printer.Print "******************************************************************************************************************************"
    Printer.Print
    Printer.Print " Cantidad      Descripcion                                             Caducidad    No. Lote        Costo Compra      Importe "
    Printer.Print "______________________________________________________________________________________________________________________________"
    Printer.Print
    For I = 1 To Grid_Entradas.Rows - 1
        Printer.Print Conectar_Ayudante.Alinea_Derecha(Grid_Entradas.TextMatrix(I, 2), 7); _
            Spc(8); Trim(Mid(Grid_Entradas.TextMatrix(I, 1), 1, 45)); _
            Spc(57 - Len(Trim(Mid(Grid_Entradas.TextMatrix(I, 1), 1, 45)))); Format(Grid_Entradas.TextMatrix(I, 6), "MM/dd/yyyy"); _
            Spc(5 - Len(Trim(Mid(Grid_Entradas.TextMatrix(I, 2), 1, 35)))); Trim(Grid_Entradas.TextMatrix(I, 5)); _
            Spc(15); "$ " & Format(Mid(Grid_Entradas.TextMatrix(I, 3), 1, 15), "#,##0.00"); _
            Spc(6 - Len(Trim(Mid(Grid_Entradas.TextMatrix(I, 5), 1, 15)))); "$ " & Conectar_Ayudante.Alinea_Derecha(Format(Grid_Entradas.TextMatrix(I, 4), "#,##0.00"), 8)
    Next I
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print Spc(4); Txt_Comentarios_Datos_Entrada.Text
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print Spc(80); "-------------------------------------"
    Printer.Print Spc(80); "SubTotal   : $"; Conectar_Ayudante.Alinea_Derecha(Format(Txt_Subtotal.Text, "#,###,##0.0000"), 13)
    Printer.Print Spc(80); "IVA        : $"; Conectar_Ayudante.Alinea_Derecha(Format(Txt_Total_IVA.Text, "#,###,##0.0000"), 13)
    Printer.Print Spc(80); "TOTAL      : $"; Conectar_Ayudante.Alinea_Derecha(Format(Txt_Total.Text, "#,###,##0.0000"), 13)
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------"
    Printer.EndDoc
    MsgBox "Entrada enviada a impresion", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Grid_Detalles_Orden_Compra_Click
'DESCRIPCIÓN                : Consulta los datos del producto segun la clave
'                           :
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 11-Nov-2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************'
Private Sub Grid_Detalles_Orden_Compra_Click()
Dim Rs_Consulta As rdoResultset
Dim Rs_Consulta_Marca As rdoResultset
Dim Mi_SQL As String
Dim Rs_Consulta_Impuesto As rdoResultset
Dim Costo_Sin_IVA As Double

On Error GoTo Handler

    'SE CONSULTAN LOS DATOS DEL PRODUCTO
    Mi_SQL = " SELECT * FROM Cat_Productos "
    Mi_SQL = Mi_SQL & " WHERE Clave='" & Trim(Grid_Detalles_Orden_Compra.TextMatrix(Grid_Detalles_Orden_Compra.RowSel, 1)) & "' "
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
        'IMPUESTO DEL PRODUCTO Y MARCA
        Txt_Clave.Text = Rs_Consulta.rdoColumns("Clave")
        Txt_Producto_ID = Rs_Consulta.rdoColumns("Producto_ID")
        Txt_Producto.Text = Rs_Consulta.rdoColumns("Descripcion")
        Costo_Sin_IVA = Val(Grid_Detalles_Orden_Compra.TextMatrix(Grid_Detalles_Orden_Compra.RowSel, 5))
        If Not IsNull(Rs_Consulta.rdoColumns("Aplica_IVA")) Then
            Txt_Aplica_IVA.Text = Trim(UCase(Rs_Consulta.rdoColumns("Aplica_IVA")))
        Else
            Txt_Aplica_IVA.Text = "NO"
        End If
        Txt_Cantidad.Text = Grid_Detalles_Orden_Compra.TextMatrix(Grid_Detalles_Orden_Compra.RowSel, 4)
        Txt_Marca.Text = Grid_Detalles_Orden_Compra.TextMatrix(Grid_Detalles_Orden_Compra.RowSel, 3)
        If Not IsNull(Rs_Consulta.rdoColumns("Impuesto_ID")) Then
            Set Rs_Consulta_Impuesto = Conectar_Ayudante.Recordset_Consultar("SELECT * FROM Cat_impuestos WHERE Impuesto_ID='" & Rs_Consulta.rdoColumns("Impuesto_ID") & "'")
            If Not Rs_Consulta_Impuesto.EOF Then
                Txt_Impuesto.Text = Rs_Consulta_Impuesto.rdoColumns("Impuesto")
            End If
        Else
            Txt_Impuesto.Text = 0
        End If
''        'SE CONSULTA LA MARCA
''        Mi_SQL = " SELECT Nombre FROM Cat_Marcas "
''        Mi_SQL = Mi_SQL & " WHERE Marca_ID='" & Format(Rs_Consulta.rdoColumns("Marca_ID"), "00000") & "' "
''        Set Rs_Consulta_Marca = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''        If Not Rs_Consulta_Marca.EOF Then
''            Txt_Marca.Text = Rs_Consulta_Marca!Nombre
''        End If
''        Rs_Consulta_Marca.Close
        
        'COSTOS
            If Trim(UCase(Rs_Consulta!Aplica_IVA)) = "SI" Then
                'Costo sin iva * Impuesto/100
                ''Txt_IVA.Text = Format(Val(Rs_Consulta.rdoColumns("Costo")) * (Val(Txt_Impuesto.Text) / 100), "##,###,###.00") 'IVA
                Txt_IVA.Text = Format(Val(Costo_Sin_IVA) * (Val(Txt_Impuesto.Text) / 100), "##,###,###.00") 'IVA
            Else
                Txt_IVA.Text = 0
            End If
            'Costo sin iva + IVA
            ''Txt_Costo.Text = Format(Val(Rs_Consulta.rdoColumns("Costo")) + Val(Txt_IVA.Text), "##,###,###.00") 'COSTO CON IVA
            Txt_Costo.Text = Format(Val(Costo_Sin_IVA) + Val(Txt_IVA.Text), "##,###,###.00") 'COSTO CON IVA
            'Cantidad * Costo con iva
            Txt_Importe_Costo_Con_Iva.Text = Format((Val(Txt_Cantidad.Text) * Val(Txt_Costo.Text)), "##,###,###.00") 'IMPORTE
        'ENTRADA
            If Not IsNull(Rs_Consulta.rdoColumns("Existencia")) Then
                Txt_Existencia.Text = Rs_Consulta.rdoColumns("Existencia") 'EXISTENCIA
            Else
                 Txt_Existencia.Text = 0
            End If
            ''Txt_Costo_Sin_IVA = Format(Rs_Consulta.rdoColumns("Costo"), "##,###,###.00") 'COSTO SIN IVA
            Txt_Costo_Sin_IVA = Format(Costo_Sin_IVA, "##,###,###.00") 'COSTO SIN IVA
    End If
    Rs_Consulta.Close
    Sst_Detalle_Orden_Compra.Tab = 1
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub



'*******************************************************************************
'NOMBRE DE LA FUNCIÓN       : Txt_Clave_KeyPress
'DESCRIPCIÓN                : Consulta los datos del producto segun la clave introducida
'                           : en la caja de texto clave
'PARÁMETROS                 :
'CREO                       : Julio Cruz
'FECHA_CREO                 : 22-Octubre-2010
'MODIFICO                   :
'FECHA_MODIFICO             :
'CAUSA_MODIFICACIÓN         :
'*******************************************************************************'
Private Sub Txt_Clave_KeyPress(KeyAscii As Integer)
Dim Rs_Consulta As rdoResultset
Dim Rs_Consulta_Marca As rdoResultset
Dim Mi_SQL As String
Dim Rs_Consulta_Impuesto As rdoResultset

    If Txt_Clave.Locked = False Then
        If KeyAscii = 13 Then
            'SE CONSULTAN LOS DATOS DEL PRODUCTO
            Mi_SQL = " SELECT * FROM Cat_Productos "
            Mi_SQL = Mi_SQL & " WHERE Clave='" & Trim(Txt_Clave.Text) & "' "
            Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta.EOF Then
                'CLAVE DEL PRODUCTO Y MARCA
                Txt_Clave.Text = Rs_Consulta.rdoColumns("Clave")
                If Not IsNull(Rs_Consulta.rdoColumns("Impuesto_ID")) Then
                    Set Rs_Consulta_Impuesto = Conectar_Ayudante.Recordset_Consultar("SELECT * FROM Cat_impuestos WHERE Impuesto_ID='" & Rs_Consulta.rdoColumns("Impuesto_ID") & "'")
                    If Not Rs_Consulta_Impuesto.EOF Then
                        Txt_Impuesto.Text = Rs_Consulta_Impuesto.rdoColumns("Impuesto")
                    End If
                Else
                    Txt_Impuesto.Text = 0
                End If
                'SE CONSULTA LA MARCA
                Mi_SQL = " SELECT Nombre FROM Cat_Marcas "
                Mi_SQL = Mi_SQL & " WHERE Marca_ID='" & Format(Rs_Consulta.rdoColumns("Marca_ID"), "00000") & "' "
                Set Rs_Consulta_Marca = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Marca.EOF Then
                    Txt_Marca.Text = Rs_Consulta_Marca!Nombre
                End If
                Rs_Consulta_Marca.Close
                'COSTOS
                    If Cmb_Aplica_IVA.Text = "SI" Then
                        'Costo sin iva * Impuesto/100
                        Txt_IVA.Text = Format(Val(Rs_Consulta.rdoColumns("Costo")) * (Val(Txt_Impuesto.Text) / 100), "##,###,###.00") 'IVA
                    Else
                        Txt_IVA.Text = 0
                    End If
                    'Costo sin iva + IVA
                    Txt_Costo.Text = Format(Val(Rs_Consulta.rdoColumns("Costo")) + Val(Txt_IVA.Text), "##,###,###.00") 'COSTO CON IVA
                    'Cantidad * Costo con iva
                    Txt_Importe_Costo_Con_Iva.Text = Format((Val(Txt_Cantidad.Text) * Val(Txt_Costo.Text)), "##,###,###.00") 'IMPORTE
                'ENTRADA
                    If Not IsNull(Rs_Consulta.rdoColumns("Existencia")) Then
                        Txt_Existencia.Text = Rs_Consulta.rdoColumns("Existencia") 'EXISTENCIA
                    Else
                         Txt_Existencia.Text = 0
                    End If
                    Txt_Costo_Sin_IVA = Format(Rs_Consulta.rdoColumns("Costo"), "##,###,###.00") 'COSTO SIN IVA
            End If
            Rs_Consulta.Close
        End If
    End If
End Sub


Private Sub Txt_Costo_Sin_IVA_Change()
    'COSTOS
    Txt_IVA.Text = Format(Val(Txt_Costo_Sin_IVA.Text) * (Val(PG_Retencion_IVA)), "#,##0.00")   'IVA
    Txt_Costo.Text = Format(Val(Txt_Costo_Sin_IVA.Text) + Val(Txt_IVA.Text), "#,##0.00")  'COSTO CON IVA
    Txt_Importe_Costo_Con_Iva.Text = Format((Val(Txt_Cantidad.Text) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Costo.Text, ","))), "#,##0.00")  'IMPORTE
    'ENTRADA
    Txt_Costo_Sin_IVA = Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")   'COSTO SIN IVA
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Descripcion_Click
'DESCRIPCIÓN          : consulta los datos del producto
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Descripcion_Click()
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset

On Error GoTo Handler
        Txt_Costo_Sin_IVA.Text = ""
        Txt_Existencia.Text = ""
        Txt_Aplica_IVA.Text = ""
        Mi_SQL = " SELECT * FROM Cat_Productos WHERE Producto_ID='" & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & "' "
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta.EOF Then
            If Not IsNull(Rs_Consulta!Precio_Venta) Then
                If Not IsNull(Rs_Consulta!Costo) Then Txt_Costo_Sin_IVA.Text = Rs_Consulta!Costo
                If Not IsNull(Rs_Consulta!Existencia) Then Txt_Existencia.Text = Rs_Consulta!Existencia
                If Not IsNull(Rs_Consulta!Aplica_IVA) Then Txt_Aplica_IVA.Text = Rs_Consulta!Aplica_IVA
            End If
        End If
        Rs_Consulta.Close
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Descripcion_KeyPress
'DESCRIPCIÓN          : Llena el combo con los productos
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Descripcion, 1, "Nombre")
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Txt_Cantidad_Change
'DESCRIPCIÓN          : Modifica, el importe
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Txt_Cantidad_Change()
    'COSTOS
    Txt_IVA.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")) * (Val(PG_Retencion_IVA)), "#,##0.00")    'IVA
    Txt_Costo.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Txt_IVA.Text, ",")), "#,##0.00")    'COSTO CON IVA
    Txt_Importe_Costo_Con_Iva.Text = Format((Val(Conectar_Ayudante.Quitar_Caracter(Txt_Cantidad.Text, ",")) * Val(Conectar_Ayudante.Quitar_Caracter(Txt_Costo.Text, ","))), "#,##0.00")   'IMPORTE
    'ENTRADA
    Txt_Costo_Sin_IVA = Conectar_Ayudante.Quitar_Caracter(Txt_Costo_Sin_IVA.Text, ",")   'COSTO SIN IVA
End Sub
