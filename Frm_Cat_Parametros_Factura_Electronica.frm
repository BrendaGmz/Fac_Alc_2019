VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Parametros_Factura_Electronica 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Factura Electrónica"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   Icon            =   "Frm_Cat_Parametros_Factura_Electronica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9555
   Begin MSComDlg.CommonDialog Cdl_Rutas 
      Left            =   30
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8100
      Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1185
   End
   Begin TabDlg.SSTab Tab_Parametros 
      Height          =   6765
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11933
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Factura Electrónica"
      TabPicture(0)   =   "Frm_Cat_Parametros_Factura_Electronica.frx":024C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Parametros_Factura"
      Tab(0).Control(1)=   "Btn_Modificar"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Serie y Folios Facturas"
      TabPicture(1)   =   "Frm_Cat_Parametros_Factura_Electronica.frx":0268
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Serie_Folios"
      Tab(1).Control(1)=   "Btn_Eliminar_Serie_Folios"
      Tab(1).Control(2)=   "Btn_Nuevo_Serie_Folios"
      Tab(1).Control(3)=   "Btn_Modificar_Serie_Folios"
      Tab(1).Control(4)=   "Fra_Parametros_Serie_Folios"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Serie y Folios Nota Crédito"
      TabPicture(2)   =   "Frm_Cat_Parametros_Factura_Electronica.frx":0284
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Btn_Notas_Credito_Modificar_Serie_Folios"
      Tab(2).Control(1)=   "Btn_Notas_Credito_Nuevo_Serie_Folios"
      Tab(2).Control(2)=   "Btn_Notas_Credito_Eliminar_Serie_Folios"
      Tab(2).Control(3)=   "Grid_Notas_Credito_Serie_Folios"
      Tab(2).Control(4)=   "Fra_Notas_Credito_Parametros_Serie_Folios"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Serie y Folios Pagos"
      TabPicture(3)   =   "Frm_Cat_Parametros_Factura_Electronica.frx":02A0
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Fra_Pagos_Parametros_Serie_Folios"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Grid_Pagos_Serie_Folios"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Btn_Pagos_Modificar_Serie_Folios"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Btn_Pagos_Nuevo_Serie_Folios"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Btn_Pagos_Eliminar_Serie_Folios"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.CommandButton Btn_Pagos_Eliminar_Serie_Folios 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   5550
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":02BC
         Style           =   1  'Graphical
         TabIndex        =   88
         Tag             =   "B"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Pagos_Nuevo_Serie_Folios 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   150
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0406
         Style           =   1  'Graphical
         TabIndex        =   87
         Tag             =   "A"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Pagos_Modificar_Serie_Folios 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   2850
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0550
         Style           =   1  'Graphical
         TabIndex        =   86
         Tag             =   "M"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Pagos_Serie_Folios 
         Height          =   4920
         Left            =   150
         TabIndex        =   85
         Top             =   1080
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   8678
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Frame Fra_Pagos_Parametros_Serie_Folios 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   6330
         Left            =   90
         TabIndex        =   75
         Top             =   360
         Width           =   9360
         Begin VB.TextBox Txt_Pagos_Parametro_ID 
            Height          =   330
            Left            =   8190
            MaxLength       =   2
            TabIndex        =   80
            Top             =   3315
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Cmb_Pagos_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Parametros_Factura_Electronica.frx":0652
            Left            =   7035
            List            =   "Frm_Cat_Parametros_Factura_Electronica.frx":0662
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Pagos_Folio_Final 
            Height          =   330
            Left            =   4865
            MaxLength       =   10
            TabIndex        =   78
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Pagos_Folio_Inicial 
            Height          =   330
            Left            =   2695
            MaxLength       =   10
            TabIndex        =   77
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Pagos_Serie 
            Height          =   330
            Left            =   525
            MaxLength       =   5
            TabIndex        =   76
            Top             =   375
            Width           =   1830
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
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
            Left            =   7530
            TabIndex        =   84
            Top             =   150
            Width           =   645
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Final"
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
            Left            =   5355
            TabIndex        =   83
            Top             =   150
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Inicial"
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
            Left            =   3120
            TabIndex        =   82
            Top             =   150
            Width           =   990
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Serie"
            Height          =   195
            Left            =   1185
            TabIndex        =   81
            Top             =   150
            Width           =   360
         End
      End
      Begin VB.CommandButton Btn_Notas_Credito_Modificar_Serie_Folios 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   -72150
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":068F
         Style           =   1  'Graphical
         TabIndex        =   64
         Tag             =   "M"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Notas_Credito_Nuevo_Serie_Folios 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   -74850
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0791
         Style           =   1  'Graphical
         TabIndex        =   63
         Tag             =   "A"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Notas_Credito_Eliminar_Serie_Folios 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   -69450
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":08DB
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "B"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Notas_Credito_Serie_Folios 
         Height          =   4920
         Left            =   -74850
         TabIndex        =   36
         Top             =   1080
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   8678
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Frame Fra_Notas_Credito_Parametros_Serie_Folios 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   6330
         Left            =   -74910
         TabIndex        =   52
         Top             =   360
         Width           =   9360
         Begin VB.TextBox Txt_Notas_Credito_Serie 
            Height          =   330
            Left            =   525
            MaxLength       =   5
            TabIndex        =   32
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Notas_Credito_Folio_Inicial 
            Height          =   330
            Left            =   2695
            MaxLength       =   10
            TabIndex        =   33
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Notas_Credito_Folio_Final 
            Height          =   330
            Left            =   4865
            MaxLength       =   10
            TabIndex        =   34
            Top             =   375
            Width           =   1830
         End
         Begin VB.ComboBox Cmb_Notas_Credito_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Parametros_Factura_Electronica.frx":0A25
            Left            =   7035
            List            =   "Frm_Cat_Parametros_Factura_Electronica.frx":0A35
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   375
            Width           =   1830
         End
         Begin VB.TextBox Txt_Notas_Credito_Parametro_ID 
            Height          =   330
            Left            =   8190
            MaxLength       =   2
            TabIndex        =   53
            Top             =   3315
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Serie"
            Height          =   195
            Left            =   1185
            TabIndex        =   57
            Top             =   150
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Inicial"
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
            Left            =   3120
            TabIndex        =   56
            Top             =   150
            Width           =   990
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Final"
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
            Left            =   5355
            TabIndex        =   55
            Top             =   150
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
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
            Left            =   7530
            TabIndex        =   54
            Top             =   150
            Width           =   645
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Serie_Folios 
         Height          =   4860
         Left            =   -74850
         TabIndex        =   27
         Top             =   1140
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   8573
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.CommandButton Btn_Eliminar_Serie_Folios 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   -69420
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0A62
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "B"
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Nuevo_Serie_Folios 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   -74850
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0BAC
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "A"
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Modificar_Serie_Folios 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   -72195
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0CF6
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "M"
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.Frame Fra_Parametros_Serie_Folios 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   6330
         Left            =   -74940
         TabIndex        =   46
         Top             =   360
         Width           =   9360
         Begin VB.TextBox Txt_Parametro_ID 
            Height          =   330
            Left            =   8190
            MaxLength       =   2
            TabIndex        =   51
            Top             =   3315
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Cmb_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Parametros_Factura_Electronica.frx":0DF8
            Left            =   7155
            List            =   "Frm_Cat_Parametros_Factura_Electronica.frx":0E08
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   375
            Width           =   1740
         End
         Begin VB.TextBox Txt_Folio_Final 
            Height          =   330
            Left            =   4905
            MaxLength       =   10
            TabIndex        =   25
            Top             =   375
            Width           =   1740
         End
         Begin VB.TextBox Txt_Folio_Inicial 
            Height          =   330
            Left            =   2655
            MaxLength       =   10
            TabIndex        =   24
            Top             =   375
            Width           =   1740
         End
         Begin VB.TextBox Txt_Serie 
            Height          =   330
            Left            =   405
            MaxLength       =   5
            TabIndex        =   23
            Top             =   375
            Width           =   1740
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
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
            Left            =   7650
            TabIndex        =   50
            Top             =   150
            Width           =   645
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Final"
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
            Left            =   5355
            TabIndex        =   49
            Top             =   150
            Width           =   885
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Folio Inicial"
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
            Left            =   3120
            TabIndex        =   48
            Top             =   150
            Width           =   990
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Serie"
            Height          =   195
            Left            =   1185
            TabIndex        =   47
            Top             =   150
            Width           =   360
         End
      End
      Begin VB.Frame Fra_Parametros_Factura 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   5655
         Left            =   -74910
         TabIndex        =   37
         Top             =   360
         Width           =   9300
         Begin VB.ComboBox Cmb_Regimen_Fiscal 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   3405
            Width           =   7185
         End
         Begin VB.TextBox Txt_Mensaje_Factura 
            Height          =   330
            Left            =   1980
            MaxLength       =   255
            TabIndex        =   16
            Top             =   3405
            Width           =   7185
         End
         Begin VB.CommandButton Btn_Seleccionar_Ruta_NC 
            Caption         =   "Seleccionar"
            Height          =   330
            Left            =   8145
            TabIndex        =   12
            Top             =   2640
            Width           =   1005
         End
         Begin VB.TextBox Txt_Ruta_Notas_Credito 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1980
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2639
            Width           =   6120
         End
         Begin VB.TextBox Txt_Porcentaje_Pagare_Interes 
            Height          =   330
            Left            =   8145
            MaxLength       =   4
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   3015
            Width           =   1020
         End
         Begin VB.TextBox Txt_Dias_Expira_Vigencia_Certificado 
            Height          =   330
            Left            =   1995
            MaxLength       =   4
            TabIndex        =   13
            Top             =   3021
            Width           =   1020
         End
         Begin VB.CommandButton Btn_Seleccionar_Ruta_Certificado 
            Caption         =   "Seleccionar"
            Height          =   330
            Left            =   8160
            TabIndex        =   72
            Top             =   240
            Width           =   1005
         End
         Begin VB.Frame Fra_Timbrado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Timbrado de facturas"
            Height          =   1815
            Left            =   120
            TabIndex        =   65
            Top             =   3795
            Width           =   9060
            Begin VB.ComboBox Cmb_Ambiente_Timbrado 
               Height          =   315
               ItemData        =   "Frm_Cat_Parametros_Factura_Electronica.frx":0E35
               Left            =   1875
               List            =   "Frm_Cat_Parametros_Factura_Electronica.frx":0E3F
               TabIndex        =   21
               Top             =   1440
               Width           =   2745
            End
            Begin VB.TextBox Txt_Version_Timbrado 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   1875
               MaxLength       =   255
               TabIndex        =   17
               Top             =   210
               Width           =   2745
            End
            Begin VB.TextBox Txt_ID_Sucursal_Timbrado 
               Height          =   330
               Left            =   6315
               MaxLength       =   255
               TabIndex        =   18
               Top             =   165
               Width           =   2625
            End
            Begin VB.TextBox Txt_Codigo_Usuario_Timbrado 
               Height          =   330
               Left            =   1875
               MaxLength       =   255
               TabIndex        =   19
               Top             =   620
               Width           =   7065
            End
            Begin VB.TextBox Txt_Codigo_Usuario_Proveedor_Timbrado 
               Height          =   330
               Left            =   1875
               MaxLength       =   50
               TabIndex        =   20
               Top             =   1030
               Width           =   7065
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Password Llave"
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
               Left            =   0
               TabIndex        =   71
               Top             =   -180
               Width           =   1350
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ambiente Timbrado"
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
               TabIndex        =   70
               Top             =   1500
               Width           =   1635
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Versión"
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
               Left            =   135
               TabIndex        =   69
               Top             =   285
               Width           =   630
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "ID Sucursal"
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
               Left            =   4815
               TabIndex        =   68
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Codigo del  Usuario"
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
               Left            =   135
               TabIndex        =   67
               Top             =   663
               Width           =   1680
            End
            Begin VB.Label Label32 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Codigo Usuario Proveedor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   135
               TabIndex        =   66
               Top             =   978
               Width           =   1515
            End
         End
         Begin VB.TextBox Txt_Dias_Expira_Folios 
            Height          =   330
            Left            =   4905
            MaxLength       =   4
            TabIndex        =   14
            Top             =   3000
            Width           =   1020
         End
         Begin VB.TextBox Txt_Ruta_Xmls 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2257
            Width           =   6120
         End
         Begin VB.CommandButton Btn_Seleccionar_Ruta_Xmls 
            Caption         =   "Seleccionar"
            Height          =   330
            Left            =   8160
            TabIndex        =   10
            Top             =   2243
            Width           =   1005
         End
         Begin VB.TextBox Txt_Ruta_Pdfs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1875
            Width           =   6120
         End
         Begin VB.CommandButton Btn_Seleccionar_Ruta_Pdfs 
            Caption         =   "Seleccionar"
            Height          =   330
            Left            =   8145
            TabIndex        =   8
            Top             =   1875
            Width           =   1005
         End
         Begin VB.TextBox Txt_Password_Llave 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   2010
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   930
            Width           =   2820
         End
         Begin VB.TextBox Txt_Ruta_Llave_Privada 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2010
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   2
            Top             =   585
            Width           =   6120
         End
         Begin VB.CommandButton Btn_Seleccionar_Ruta_Llave_Privada 
            Caption         =   "Seleccionar"
            Height          =   330
            Left            =   8145
            TabIndex        =   3
            Top             =   585
            Width           =   1005
         End
         Begin VB.TextBox Txt_Ruta_Certificado 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2010
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   6120
         End
         Begin VB.Frame Fra_Parametros_Vigencia_Certificado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vigencia Certificado"
            Height          =   675
            Left            =   75
            TabIndex        =   41
            Top             =   1185
            Width           =   9090
            Begin MSComCtl2.DTPicker Dtp_Vigencia_Desde 
               Height          =   330
               Left            =   1930
               TabIndex        =   5
               Top             =   195
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MMM/yyyy HH:mm:ss"
               Format          =   34996227
               CurrentDate     =   40054
            End
            Begin MSComCtl2.DTPicker Dtp_Vigencia_Hasta 
               Height          =   330
               Left            =   6060
               TabIndex        =   6
               Top             =   195
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MMM/yyyy HH:mm:ss"
               Format          =   34996227
               CurrentDate     =   40054
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vigencia Hasta"
               Height          =   195
               Left            =   4815
               TabIndex        =   43
               Top             =   300
               Width           =   1080
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vigencia Desde"
               Height          =   195
               Left            =   390
               TabIndex        =   42
               Top             =   300
               Width           =   1125
            End
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ruta Notas Crédito"
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
            Top             =   2715
            Width           =   1620
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aviso Expira Vigencia"
            Height          =   195
            Left            =   135
            TabIndex        =   61
            Top             =   3075
            Width           =   1530
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aviso Expira Folios"
            Height          =   195
            Left            =   3435
            TabIndex        =   60
            Top             =   3075
            Width           =   1320
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Regimen Fiscal"
            Height          =   195
            Left            =   135
            TabIndex        =   59
            Top             =   3480
            Width           =   1080
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ruta Xmls"
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
            Left            =   135
            TabIndex        =   45
            Top             =   2311
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ruta Pdfs"
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
            Left            =   135
            TabIndex        =   44
            Top             =   1943
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Password Llave"
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
            Left            =   135
            TabIndex        =   40
            Top             =   990
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ruta Llave Privada"
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
            Left            =   135
            TabIndex        =   39
            Top             =   645
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ruta Certificado"
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
            Left            =   135
            TabIndex        =   38
            Top             =   315
            Width           =   1395
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "% Interes Pagare"
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
            Left            =   6615
            TabIndex        =   58
            Top             =   3090
            Width           =   1455
         End
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   500
         Left            =   -74895
         Picture         =   "Frm_Cat_Parametros_Factura_Electronica.frx":0E57
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "M"
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
End
Attribute VB_Name = "Frm_Cat_Parametros_Factura_Electronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Alta_Notas_Credito_Series_Folios
'DESCRIPCIÓN: Da de alta las series y folios con sus datos de año y autorizacion
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 15/Julio/2010 12:40pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Notas_Credito_Series_Folios()
Dim Rs_Alta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Alta de los parametros
    Set Rs_Alta_Parametros = Conectar_Ayudante.Recordset_Agregar("Cat_Parametros_Factura_Electronica_Folios")
        With Rs_Alta_Parametros
            .AddNew
                .rdoColumns("Tipo") = "NOTA_CREDITO"
                .rdoColumns("Serie") = Trim(Txt_Notas_Credito_Serie.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Notas_Credito_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Notas_Credito_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Notas_Credito_Estatus.text
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
    Rs_Alta_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Notas_Credito_Grid_Series_Folios
    MsgBox "Parámetros capturados"
    Fra_Notas_Credito_Parametros_Serie_Folios.Enabled = False
    Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = True
    Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = True
    Btn_Notas_Credito_Nuevo_Serie_Folios.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Grid_Notas_Credito_Serie_Folios.Enabled = True
    Limpia_Notas_Credito_Controles_Serie_Folios
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Alta_Series_Folios
'DESCRIPCIÓN: Da de alta las series y folios con sus datos de año y autorizacion
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 1:25pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Series_Folios()
Dim Rs_Alta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    'Alta de los parametros
    Set Rs_Alta_Parametros = Conectar_Ayudante.Recordset_Agregar("Cat_Parametros_Factura_Electronica_Folios")
        With Rs_Alta_Parametros
            .AddNew
                .rdoColumns("Tipo") = "FACTURA"
                .rdoColumns("Serie") = Trim(Txt_Serie.text)
'                .rdoColumns("Año") = Val(Txt_Año.text)
'                .rdoColumns("No_Autorizacion") = Val(Txt_No_Autorizacion.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Estatus.text
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
    Rs_Alta_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Grid_Series_Folios
    MsgBox "Parámetros capturados"
    Fra_Parametros_Serie_Folios.Enabled = False
    Btn_Modificar_Serie_Folios.Enabled = True
    Btn_Eliminar_Serie_Folios.Enabled = True
    Btn_Nuevo_Serie_Folios.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Grid_Serie_Folios.Enabled = True
    Limpia_Controles_Serie_Folios
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub
Private Sub Alta_Pagos_Series_Folios()
Dim Rs_Alta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    'Alta de los parametros
    Set Rs_Alta_Parametros = Conectar_Ayudante.Recordset_Agregar("Cat_Parametros_Factura_Electronica_Folios")
        With Rs_Alta_Parametros
            .AddNew
                .rdoColumns("Tipo") = "PAGOS"
                .rdoColumns("Serie") = Trim(Txt_Pagos_Serie.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Pagos_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Pagos_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Pagos_Estatus.text
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
    Rs_Alta_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Pagos_Grid_Series_Folios
    MsgBox "Parámetros capturados"
    Fra_Pagos_Parametros_Serie_Folios.Enabled = False
    Btn_Pagos_Modificar_Serie_Folios.Enabled = True
    Btn_Pagos_Eliminar_Serie_Folios.Enabled = True
    Btn_Pagos_Nuevo_Serie_Folios.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Grid_Pagos_Serie_Folios.Enabled = True
    Limpia_Pagos_Controles_Serie_Folios
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Limpia_Controles_Serie_Folios
'DESCRIPCIÓN: Limpia los controles de los parametros de folios
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 1:45pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Limpia_Controles_Serie_Folios()
    Txt_Parametro_ID.text = ""
    Txt_Serie.text = ""
    Txt_Folio_Inicial.text = ""
    Txt_Folio_Final.text = ""
    Cmb_Estatus.ListIndex = -1
End Sub

Private Sub Limpia_Pagos_Controles_Serie_Folios()
    Txt_Pagos_Parametro_ID.text = ""
    Txt_Pagos_Serie.text = ""
    Txt_Pagos_Folio_Inicial.text = ""
    Txt_Pagos_Folio_Final.text = ""
    Cmb_Pagos_Estatus.ListIndex = -1
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Limpia_Notas_Credito_Controles_Serie_Folios
'DESCRIPCIÓN: Limpia los controles de los parametros de folios
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 15/Julio/2010 12:20pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Limpia_Notas_Credito_Controles_Serie_Folios()
    Txt_Notas_Credito_Parametro_ID.text = ""
    Txt_Notas_Credito_Serie.text = ""
    Txt_Notas_Credito_Folio_Inicial.text = ""
    Txt_Notas_Credito_Folio_Final.text = ""
    Cmb_Notas_Credito_Estatus.ListIndex = -1
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Modifica_Notas_Credito_Series_Folios
'DESCRIPCIÓN: Modifica las series y folios con sus datos de año y autorizacion
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 15/Julio/2010 12:40pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Notas_Credito_Series_Folios()
Dim Rs_Modifica_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Actualiza de los parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Notas_Credito_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'NOTA_CREDITO'"
    Set Rs_Modifica_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Modifica_Parametros
            .Edit
                .rdoColumns("Serie") = Trim(Txt_Notas_Credito_Serie.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Notas_Credito_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Notas_Credito_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Notas_Credito_Estatus.text
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    Rs_Modifica_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Notas_Credito_Grid_Series_Folios
    MsgBox "Folios actualizados"
    Fra_Notas_Credito_Parametros_Serie_Folios.Enabled = False
    Btn_Notas_Credito_Nuevo_Serie_Folios.Enabled = True
    Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = True
    Btn_Notas_Credito_Modificar_Serie_Folios.Caption = "Modificar"
    Grid_Notas_Credito_Serie_Folios.Enabled = True
    Btn_Salir.Caption = "Salir"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

Private Sub Modifica_Pagos_Series_Folios()
Dim Rs_Modifica_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Actualiza de los parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Pagos_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'PAGOS'"
    Set Rs_Modifica_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Modifica_Parametros
            .Edit
                .rdoColumns("Serie") = Trim(Txt_Pagos_Serie.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Pagos_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Pagos_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Pagos_Estatus.text
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    Rs_Modifica_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Pagos_Grid_Series_Folios
    MsgBox "Folios actualizados"
    Fra_Pagos_Parametros_Serie_Folios.Enabled = False
    Btn_Pagos_Nuevo_Serie_Folios.Enabled = True
    Btn_Pagos_Eliminar_Serie_Folios.Enabled = True
    Btn_Pagos_Modificar_Serie_Folios.Caption = "Modificar"
    Grid_Pagos_Serie_Folios.Enabled = True
    Btn_Salir.Caption = "Salir"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Modifica_Series_Folios
'DESCRIPCIÓN: Modifica las series y folios con sus datos de año y autorizacion
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 1:40pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Series_Folios()
Dim Rs_Modifica_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Actualiza de los parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'FACTURA'"
    Set Rs_Modifica_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Modifica_Parametros
            .Edit
                .rdoColumns("Serie") = Trim(Txt_Serie.text)
                .rdoColumns("Año") = Val(Txt_Año.text)
                .rdoColumns("No_Autorizacion") = Val(Txt_No_Autorizacion.text)
                .rdoColumns("Folio_Inicial") = Val(Txt_Folio_Inicial.text)
                .rdoColumns("Folio_Final") = Val(Txt_Folio_Final.text)
                .rdoColumns("Estatus") = Cmb_Estatus.text
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    Rs_Modifica_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Grid_Series_Folios
    MsgBox "Parámetros actualizados"
    Fra_Parametros_Serie_Folios.Enabled = False
    Btn_Nuevo_Serie_Folios.Enabled = True
    Btn_Eliminar_Serie_Folios.Enabled = True
    Btn_Modificar_Serie_Folios.Caption = "Modificar"
    Grid_Serie_Folios.Enabled = True
    Btn_Salir.Caption = "Salir"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Parametros
'DESCRIPCIÓN: Consulta la tabla de los parametros de factura electronica
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 1:00pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Parametros()
Dim Rs_Consulta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Consulta la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Parametros
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Ruta_Certificado")) Then Txt_Ruta_Certificado.text = .rdoColumns("Ruta_Certificado")
            If Not IsNull(.rdoColumns("Ruta_Llave_Privada")) Then Txt_Ruta_Llave_Privada.text = .rdoColumns("Ruta_Llave_Privada")
            If Not IsNull(.rdoColumns("Password_Llave_Privada")) Then Txt_Password_Llave.text = .rdoColumns("Password_Llave_Privada")
            If Not IsNull(.rdoColumns("Vigencia_Certificado_Desde")) Then Dtp_Vigencia_Desde.Value = Format(.rdoColumns("Vigencia_Certificado_Desde"), "yyyy/MM/dd HH:mm:ss")
            If Not IsNull(.rdoColumns("Vigencia_Certificado_Hasta")) Then Dtp_Vigencia_Hasta.Value = Format(.rdoColumns("Vigencia_Certificado_Hasta"), "yyyy/MM/dd HH:mm:ss")
            If Not IsNull(.rdoColumns("Dias_Aviso_Expira_Vigencia")) Then Txt_Dias_Expira_Vigencia_Certificado.text = .rdoColumns("Dias_Aviso_Expira_Vigencia")
            If Not IsNull(.rdoColumns("Dias_Aviso_Termina_Folios")) Then Txt_Dias_Expira_Folios.text = .rdoColumns("Dias_Aviso_Termina_Folios")
            If Not IsNull(.rdoColumns("Ruta_Pdfs")) Then Txt_Ruta_Pdfs.text = .rdoColumns("Ruta_Pdfs")
            If Not IsNull(.rdoColumns("Ruta_Xmls")) Then Txt_Ruta_Xmls.text = .rdoColumns("Ruta_Xmls")
            If Not IsNull(.rdoColumns("Ruta_NC")) Then Txt_Ruta_Notas_Credito.text = .rdoColumns("Ruta_NC")
            If Not IsNull(.rdoColumns("Porcentaje_Interes_Pagare")) Then Txt_Porcentaje_Pagare_Interes.text = .rdoColumns("Porcentaje_Interes_Pagare")
            If Not IsNull(.rdoColumns("Mensaje_Factura")) Then Txt_Mensaje_Factura.text = .rdoColumns("Mensaje_Factura")
            If Not IsNull(.rdoColumns("Version_Timbrado")) Then Txt_Version_Timbrado.text = Trim(.rdoColumns("Version_Timbrado"))
            If Not IsNull(.rdoColumns("Codigo_Usuario")) Then Txt_Codigo_Usuario_Timbrado.text = Trim(.rdoColumns("Codigo_Usuario"))
            If Not IsNull(.rdoColumns("Codigo_Usuario_Proveedor")) Then Txt_Codigo_Usuario_Proveedor_Timbrado.text = Trim(.rdoColumns("Codigo_Usuario_Proveedor"))
            If Not IsNull(.rdoColumns("ID_Sucursal_Timbrado")) Then Txt_ID_Sucursal_Timbrado.text = Trim(.rdoColumns("ID_Sucursal_Timbrado"))
            If Not IsNull(.rdoColumns("Ambiente_Timbrado")) Then
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Ambiente_Timbrado"), Cmb_Ambiente_Timbrado)
            Else
                Cmb_Ambiente_Timbrado.ListIndex = -1
            End If
        End If
    End With
    Rs_Consulta_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Grid_Series_Folios
'DESCRIPCIÓN: Consulta la tabla de los parametros de folios de factura electronica
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 4:40pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Grid_Series_Folios()
Dim Rs_Consulta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Encabezado del grid
    Grid_Serie_Folios.Cols = 5
    Grid_Serie_Folios.Rows = 0
    Grid_Serie_Folios.AddItem "ID" & Chr(9) & "Serie" & Chr(9) & "Folio Inicial" & Chr(9) & "Folio Final" & Chr(9) & "Estatus"
    
    'Consulta la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios WHERE Tipo = 'FACTURA' ORDER BY Folio_Inicial"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametros
            If Not .EOF Then
                While Not .EOF
                    Grid_Serie_Folios.AddItem .rdoColumns("Parametro_Folio_ID") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Folio_Inicial") & Chr(9) & .rdoColumns("Folio_Final") & Chr(9) & .rdoColumns("Estatus")
                    .MoveNext
                Wend
            End If
        End With
    Rs_Consulta_Parametros.Close
    
    If Grid_Serie_Folios.Rows > 1 Then
        Grid_Serie_Folios.FixedRows = 1
        Grid_Serie_Folios.ColWidth(0) = 0
        Grid_Serie_Folios.ColWidth(1) = 2000
        Grid_Serie_Folios.ColWidth(2) = 2000
        Grid_Serie_Folios.ColWidth(3) = 2000
        Grid_Serie_Folios.ColWidth(4) = 2000
'        Grid_Serie_Folios.ColWidth(5) = 1800
'        Grid_Serie_Folios.ColWidth(6) = 1400
    End If
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

Private Sub Consulta_Pagos_Grid_Series_Folios()
Dim Rs_Consulta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Encabezado del grid
    Grid_Pagos_Serie_Folios.Cols = 5
    Grid_Pagos_Serie_Folios.Rows = 0
    Grid_Pagos_Serie_Folios.AddItem "ID" & Chr(9) & "Serie" & Chr(9) & "Folio Inicial" & Chr(9) & "Folio Final" & Chr(9) & "Estatus"
    
    'Consulta la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios WHERE Tipo = 'PAGOS' ORDER BY Folio_Inicial"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametros
            If Not .EOF Then
                While Not .EOF
                    Grid_Pagos_Serie_Folios.AddItem .rdoColumns("Parametro_Folio_ID") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Folio_Inicial") & Chr(9) & .rdoColumns("Folio_Final") & Chr(9) & .rdoColumns("Estatus")
                    .MoveNext
                Wend
            End If
        End With
    Rs_Consulta_Parametros.Close
    
    If Grid_Pagos_Serie_Folios.Rows > 1 Then
        Grid_Pagos_Serie_Folios.FixedRows = 1
        Grid_Pagos_Serie_Folios.ColWidth(0) = 0
        Grid_Pagos_Serie_Folios.ColWidth(1) = 2000
        Grid_Pagos_Serie_Folios.ColWidth(2) = 2000
        Grid_Pagos_Serie_Folios.ColWidth(3) = 2000
        Grid_Pagos_Serie_Folios.ColWidth(4) = 2000
    End If
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Notas_Credito_Grid_Series_Folios
'DESCRIPCIÓN: Consulta la tabla de los parametros de folios de notas de credito electronica
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 15/Julio/2010 12:50pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Notas_Credito_Grid_Series_Folios()
Dim Rs_Consulta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Encabezado del grid
    Grid_Notas_Credito_Serie_Folios.Cols = 5
    Grid_Notas_Credito_Serie_Folios.Rows = 0
    Grid_Notas_Credito_Serie_Folios.AddItem "ID" & Chr(9) & "Serie" & Chr(9) & "Folio Inicial" & Chr(9) & "Folio Final" & Chr(9) & "Estatus"
    'Consulta la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios WHERE Tipo = 'NOTA_CREDITO' ORDER BY Folio_Inicial"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametros
            If Not .EOF Then
                While Not .EOF
                    Grid_Notas_Credito_Serie_Folios.AddItem .rdoColumns("Parametro_Folio_ID") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Folio_Inicial") & Chr(9) & .rdoColumns("Folio_Final") & Chr(9) & .rdoColumns("Estatus")
                    .MoveNext
                Wend
            End If
        End With
    Rs_Consulta_Parametros.Close
    
    If Grid_Notas_Credito_Serie_Folios.Rows > 1 Then
        Grid_Notas_Credito_Serie_Folios.FixedRows = 1
        Grid_Notas_Credito_Serie_Folios.ColWidth(0) = 0
        Grid_Notas_Credito_Serie_Folios.ColWidth(1) = 2000
        Grid_Notas_Credito_Serie_Folios.ColWidth(2) = 2000
        Grid_Notas_Credito_Serie_Folios.ColWidth(3) = 2000
        Grid_Notas_Credito_Serie_Folios.ColWidth(4) = 2000
    End If
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Modifica_Parametros
'DESCRIPCIÓN: Actualiza la tabla de los parametros de factura electronica
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 01/Sept/2010
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Parametros()
Dim Rs_Modifica_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Consulta la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Modifica_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Modifica_Parametros
            If Not .EOF Then 'Si existe lo actualiza
                .Edit
                    .rdoColumns("Ruta_Certificado") = Trim(Txt_Ruta_Certificado.text)
                    .rdoColumns("Ruta_Llave_Privada") = Trim(Txt_Ruta_Llave_Privada.text)
                    .rdoColumns("Password_Llave_Privada") = Trim(Txt_Password_Llave.text)
                    .rdoColumns("Vigencia_Certificado_Desde") = Format(Dtp_Vigencia_Desde.Value, "yyyy/MM/dd HH:mm:ss")
                    .rdoColumns("Vigencia_Certificado_Hasta") = Format(Dtp_Vigencia_Hasta.Value, "yyyy/MM/dd HH:mm:ss")
                    .rdoColumns("Dias_Aviso_Expira_Vigencia") = Val(Txt_Dias_Expira_Vigencia_Certificado.text)
                    .rdoColumns("Dias_Aviso_Termina_Folios") = Val(Txt_Dias_Expira_Folios.text)
                    .rdoColumns("Ruta_Pdfs") = Trim(Txt_Ruta_Pdfs.text)
                    .rdoColumns("Ruta_Xmls") = Trim(Txt_Ruta_Xmls.text)
                    .rdoColumns("Ruta_NC") = Trim(Txt_Ruta_Notas_Credito.text)
'                    .rdoColumns("Ruta_SAT") = Trim(Txt_Ruta_Archivos_SAT.text)
'                    .rdoColumns("Porcentaje_Flete") = Val(Txt_Porcentaje_Flete.text)
'                    .rdoColumns("Porcentaje_IVA") = Txt_Porcentaje_IVA.text
'                    .rdoColumns("Porcentaje_Retencion_ISR") = Txt_Porcentaje_Retencion.text
'                    .rdoColumns("Porcentaje_Retencion_IVA") = Txt_Retencion_IVA.text
'                    .rdoColumns("Porcentaje_Retencion_Cedular") = Txt_Retencion_Cedular.text
                    .rdoColumns("Porcentaje_Interes_Pagare") = Val(Txt_Porcentaje_Pagare_Interes.text)
'                    .rdoColumns("Email") = Txt_Email.text
'                    .rdoColumns("Email_Destino_SAT") = Txt_Email_Destino_SAT.text
'                    .rdoColumns("Ruta_BD_Externa") = Txt_Ruta_BD_Externa.text
'                    .rdoColumns("Password_BD_Externa") = Txt_Password_BD_Externa.text
                    .rdoColumns("Mensaje_Factura") = Trim(Txt_Mensaje_Factura.text)
                    .rdoColumns("Version_Timbrado") = Trim(Txt_Version_Timbrado.text)
                    .rdoColumns("Codigo_Usuario") = Trim(Txt_Codigo_Usuario_Timbrado.text)
                    .rdoColumns("Codigo_Usuario_Proveedor") = Trim(Txt_Codigo_Usuario_Proveedor_Timbrado.text)
                    .rdoColumns("ID_Sucursal_Timbrado") = Trim(Txt_ID_Sucursal_Timbrado.text)
                    .rdoColumns("Ambiente_Timbrado") = Trim(Cmb_Ambiente_Timbrado.text)
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            Else 'Si no existe crea el registro
                .AddNew
                    .rdoColumns("Parametro_ID") = Val(Conectar_Ayudante.Maximo_Catalogo("Cat_Parametros_Factura_Electronica", "Parametro_ID"))
                    .rdoColumns("Ruta_Certificado") = Trim(Txt_Ruta_Certificado.text)
                    .rdoColumns("Ruta_Llave_Privada") = Trim(Txt_Ruta_Llave_Privada.text)
                    .rdoColumns("Password_Llave_Privada") = Trim(Txt_Password_Llave.text)
                    .rdoColumns("Vigencia_Certificado_Desde") = Format(Dtp_Vigencia_Desde.Value, "yyyy/MM/dd HH:mm:ss")
                    .rdoColumns("Vigencia_Certificado_Hasta") = Format(Dtp_Vigencia_Hasta.Value, "yyyy/MM/dd HH:mm:ss")
                    .rdoColumns("Dias_Aviso_Expira_Vigencia") = Val(Txt_Dias_Expira_Vigencia_Certificado.text)
                    .rdoColumns("Dias_Aviso_Termina_Folios") = Val(Txt_Dias_Expira_Folios.text)
                    .rdoColumns("Ruta_Pdfs") = Trim(Txt_Ruta_Pdfs.text)
                    .rdoColumns("Ruta_Xmls") = Trim(Txt_Ruta_Xmls.text)
                    .rdoColumns("Ruta_NC") = Trim(Txt_Ruta_Notas_Credito.text)
'                    .rdoColumns("Ruta_SAE") = Trim(Txt_Ruta_SAE.text)
'                    .rdoColumns("Porcentaje_Flete") = Val(Txt_Porcentaje_Flete.text)
'                    .rdoColumns("Porcentaje_IVA") = Txt_Porcentaje_IVA.text
'                    .rdoColumns("Porcentaje_Retencion_ISR") = Txt_Porcentaje_Retencion.text
'                    .rdoColumns("Porcentaje_Retencion_IVA") = Txt_Retencion_IVA.text
'                    .rdoColumns("Porcentaje_Retencion_Cedular") = Txt_Retencion_Cedular.text
                    .rdoColumns("Porcentaje_Interes_Pagare") = Val(Txt_Porcentaje_Pagare_Interes.text)
'                    .rdoColumns("Email") = Txt_Email.text
'                    .rdoColumns("Email_Destino_SAT") = Txt_Email_Destino_SAT.text
                    .rdoColumns("Mensaje_Factura") = Trim(Txt_Mensaje_Factura.text)
'                    .rdoColumns("Ruta_BD_Externa") = Txt_Ruta_BD_Externa.text
'                    .rdoColumns("Password_BD_Externa") = Txt_Password_BD_Externa.text
                    .rdoColumns("Version_Timbrado") = Trim(Txt_Version_Timbrado.text)
                    .rdoColumns("Codigo_Usuario") = Trim(Txt_Codigo_Usuario_Timbrado.text)
                    .rdoColumns("Codigo_Usuario_Proveedor") = Trim(Txt_Codigo_Usuario_Proveedor_Timbrado.text)
                    .rdoColumns("ID_Sucursal_Timbrado") = Trim(Txt_ID_Sucursal_Timbrado.text)
                    .rdoColumns("Ambiente_Timbrado") = Trim(Cmb_Ambiente_Timbrado.text)
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End If
        End With
    Rs_Modifica_Parametros.Close
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox "Parámetros actualizados"
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Fra_Parametros_Factura.Enabled = False
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Elimina_Series_Folios
'DESCRIPCIÓN: Elimina el registro de la tabla de los parametros de folios
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 1:50pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Elimina_Series_Folios()
Dim Rs_Elimina_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Elimina el registro de la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'FACTURA'"
    Set Rs_Elimina_Parametros = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
        With Rs_Elimina_Parametros
            If Not .EOF Then
                .Delete
            End If
        End With
    Rs_Elimina_Parametros.Close
    Limpia_Controles_Serie_Folios
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Grid_Series_Folios
    MsgBox "Parámetros eliminados"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Elimina_Notas_Credito_Series_Folios
'DESCRIPCIÓN: Elimina el registro de la tabla de los parametros de folios
'PARÁMETROS :
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 15/Julio/2010 12:50pm
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Elimina_Notas_Credito_Series_Folios()
Dim Rs_Elimina_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Elimina el registro de la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Notas_Credito_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'NOTA_CREDITO'"
    Set Rs_Elimina_Parametros = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
        With Rs_Elimina_Parametros
            If Not .EOF Then
                .Delete
            End If
        End With
    Rs_Elimina_Parametros.Close
    Limpia_Notas_Credito_Controles_Serie_Folios
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Notas_Credito_Grid_Series_Folios
    MsgBox "Rango de folios eliminados"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

Private Sub Elimina_Pagos_Series_Folios()
Dim Rs_Elimina_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Elimina el registro de la tabla de parametros
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Parametro_Folio_ID = " & Val(Txt_Pagos_Parametro_ID.text)
    Mi_SQL = Mi_SQL & " AND Tipo = 'PAGOS'"
    Set Rs_Elimina_Parametros = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
        With Rs_Elimina_Parametros
            If Not .EOF Then
                .Delete
            End If
        End With
    Rs_Elimina_Parametros.Close
    Limpia_Pagos_Controles_Serie_Folios
    MDIFrm_Apl_Principal.MousePointer = 0
    Consulta_Pagos_Grid_Series_Folios
    MsgBox "Rango de folios eliminados"
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Folios
'DESCRIPCIÓN: Realiza la validacion de las series y folios para no duplicarlos
'PARÁMETROS : Valida_Folios, regresa falso o verdadero
'             Tipo_Folios, pasa el tipo de folio si es FACTURA O NOTA CREDITO
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 4:50pm
'MODIFICO          : Ismael Prieto Sánchez
'FECHA_MODIFICO    : 15/Julio/2010 12:20pm
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Valida_Folios(Tipo_Folio As String) As Boolean
Dim Rs_Valida_Folios As rdoResultset  'Variable para el manejo de la tabla
Dim Valida As Boolean                   'Variable para regresar si valido o no

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'If Val(Txt_Folio_Inicial.Text) <> Val(Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 4)) Or Val(Txt_Folio_Final.Text) <> Val(Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 5)) Then
        'Realiza la validacion de los folios
        Mi_SQL = "SELECT TOP 1 Folio_Final FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Folio_Final >= " & Val(Txt_Folio_Inicial.text)
        Mi_SQL = Mi_SQL & " AND Folio_Final <= " & Val(Txt_Folio_Final.text)
        Mi_SQL = Mi_SQL & " AND Serie = '" & Trim(Txt_Serie.text) & "'"
        Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Folio_Final DESC"
        Set Rs_Valida_Folios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Valida_Folios
                If Not .EOF Then
                    Valida = False
                Else
                    Valida = True
                End If
            End With
        Rs_Valida_Folios.Close
'    Else
'        Valida = True
'    End If
    
    'Regresa el valor
    Valida_Folios = Valida
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Function
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Folios_Predeterminados
'DESCRIPCIÓN: Realiza la validacion de los predeterminados
'PARÁMETROS : Valida_Folios_Predeterminados, regresa falso o verdadero
'             Tipo_Folio, tipo de folio asignado FACTURA O NOTA CREDITO
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 5:00pm
'MODIFICO          : Ismael Prieto Sánchez
'FECHA_MODIFICO    : 15/Julio/2010 12:40pm
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Valida_Folios_Predeterminados(Tipo_Folio As String) As Boolean
Dim Rs_Valida_Folios As rdoResultset  'Variable para el manejo de la tabla
Dim Valida As Boolean                   'Variable para regresar si valido o no

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Realiza la validacion de los folios
    If Cmb_Estatus.text = "ACTIVO" Then
        Mi_SQL = "SELECT COUNT(*) AS Predeterminado FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO' AND Serie = '" & Trim(Txt_Notas_Credito_Serie.text) & "'"
        Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
        Set Rs_Valida_Folios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Valida_Folios
                If Not .EOF Then
                    If Val(.rdoColumns("Predeterminado")) = 0 Then
                        Valida = True
                    Else
                        Valida = False
                    End If
                Else
                    Valida = True
                End If
            End With
        Rs_Valida_Folios.Close
    Else
        Valida = True
    End If
    
    'Regresa el valor
    Valida_Folios_Predeterminados = Valida
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Function
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Folios_Consecutivo_Predeterminados
'DESCRIPCIÓN: Realiza la validacion del consecutivo de los folios de los predeterminados
'PARÁMETROS : Valida_Folios_Consecutivo_Predeterminados, regresa falso o verdadero
'             Tipo_Folio, tipo de folio asignado FACTURA O NOTA CREDITO
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 01/Sep/2009 7:30pm
'MODIFICO          : Ismael Prieto Sánchez
'FECHA_MODIFICO    : 15/Julio/2010 12:30pm
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Valida_Folios_Consecutivo_Predeterminados(Tipo_Folio As String) As Boolean
Dim Rs_Valida_Folios As rdoResultset  'Variable para el manejo de la tabla
Dim Valida As Boolean                   'Variable para regresar si valido o no

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Realiza la validacion de los folios
    If Cmb_Estatus.text = "ACTIVO" Then
        Mi_SQL = "SELECT COUNT(*) AS Predeterminado FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE (Estatus = 'ACTIVO' OR Estatus = 'PENDIENTE') AND Serie = '" & Trim(Txt_Serie.text) & "'"
        Mi_SQL = Mi_SQL & " AND Folio_Inicial < " & Val(Txt_Folio_Inicial.text)
        Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
        Set Rs_Valida_Folios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Valida_Folios
                If Not .EOF Then
                    If Val(.rdoColumns("Predeterminado")) = 0 Then
                        Valida = True
                    Else
                        Valida = False
                    End If
                Else
                    Valida = True
                End If
            End With
        Rs_Valida_Folios.Close
    Else
        Valida = True
    End If
    
    'Regresa el valor
    Valida_Folios_Consecutivo_Predeterminados = Valida
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Function
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Function

Private Sub Btn_Eliminar_Serie_Folios_Click()
Dim Respuesta As Integer    'Almacena la respuesta del usuario

    If Txt_Parametro_ID.text <> "" Then
        Respuesta = MsgBox("¿Esta seguro de eliminar el folio?", vbYesNo + vbCritical)
        If Respuesta = vbYes Then
            Elimina_Series_Folios
        End If
    Else
        MsgBox ("Seleccione una serie/folio a eliminar")
    End If
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
        Btn_Modificar.Caption = "Actualizar"
        Btn_Salir.Caption = "Cancelar"
        Fra_Parametros_Factura.Enabled = True
    Else
        'Valida los controles
        If Txt_Ruta_Certificado.text = "" Then
            MsgBox "Seleccione el certificado digital", vbExclamation
            Exit Sub
        End If
        If Txt_Ruta_Llave_Privada.text = "" Then
            MsgBox "Seleccione la llave privada", vbExclamation
            Exit Sub
        End If
        If Txt_Ruta_Pdfs.text = "" Then
            MsgBox "Seleccione la ruta de almacenamiento de los Pdfs", vbExclamation
            Exit Sub
        End If
        If Txt_Ruta_Xmls.text = "" Then
            MsgBox "Seleccione la ruta de almacenamiento de los Xmls", vbExclamation
            Exit Sub
        End If
        If Txt_Ruta_Notas_Credito.text = "" Then
            MsgBox "Seleccione la ruta de almacenamiento de los archivos de notas de crédito", vbExclamation
            Exit Sub
        End If
        If Txt_Password_Llave.text = "" Then
            MsgBox "Proporcione el password del certificado digital", vbExclamation
            Txt_Password_Llave.SetFocus
            Exit Sub
        End If
        If Txt_Porcentaje_Pagare_Interes.text = "" Then
            MsgBox "Proporcione el porcentaje de interes del pagare", vbExclamation
            Txt_Porcentaje_Pagare_Interes.SetFocus
            Exit Sub
        End If
        'Valida los parametros del timbrado
        If Txt_Codigo_Usuario_Timbrado.text = "" Then
            MsgBox "Ingrese el código de usuario", vbExclamation
            Txt_Codigo_Usuario_Timbrado.SetFocus
            Exit Sub
        End If
        If Txt_Codigo_Usuario_Proveedor_Timbrado.text = "" Then
            MsgBox "Ingrese el código de usuario ante el proveedor", vbExclamation
            Txt_Codigo_Usuario_Proveedor_Timbrado.SetFocus
            Exit Sub
        End If
        If Txt_ID_Sucursal_Timbrado.text = "" Then
            MsgBox "Ingrese el ID de la sucursal", vbExclamation
            Txt_ID_Sucursal_Timbrado.SetFocus
            Exit Sub
        End If
        If Cmb_Ambiente_Timbrado.ListIndex = -1 Then
            MsgBox "Seleccione el ambiente para la generación de timbrado", vbExclamation
            Cmb_Ambiente_Timbrado.SetFocus
            Exit Sub
        End If
        Modifica_Parametros
    End If
End Sub

Private Sub Btn_Modificar_Serie_Folios_Click()
    If Txt_Parametro_ID.text <> "" Then
       If Btn_Modificar_Serie_Folios.Caption = "Modificar" Then
            Fra_Parametros_Serie_Folios.Enabled = True
            Btn_Nuevo_Serie_Folios.Enabled = False
            Btn_Eliminar_Serie_Folios.Enabled = False
            Btn_Modificar_Serie_Folios.Caption = "Actualizar"
            Btn_Salir.Caption = "Cancelar"
            Grid_Serie_Folios.Enabled = False
        Else
            If Txt_Parametro_ID.text <> "" Then
'                If Val(Txt_Año.text) <= 0 Then
'                    MsgBox ("Proporcione el año de los folios")
'                    Txt_Año.SetFocus
'                    Exit Sub
'                End If
'                If Val(Txt_No_Autorizacion.text) <= 0 Then
'                    MsgBox ("Proporcione el no. de autorización de los folios")
'                    Txt_No_Autorizacion.SetFocus
'                    Exit Sub
'                End If
                If Val(Txt_Folio_Inicial.text) <= 0 Then
                    MsgBox ("Proporcione el no. inicial de folio")
                    Txt_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Folio_Final.text) <= 0 Then
                    MsgBox ("Proporcione el no. final de folio")
                    Txt_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Folio_Inicial.text) > Val(Txt_Folio_Final.text) Then
                    MsgBox ("El folio inicial no puede ser mayor al folio final.")
                    Txt_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Folio_Final.text) < Val(Txt_Folio_Inicial.text) Then
                    MsgBox ("El folio final no puede ser menor al folio inicial.")
                    Txt_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Cmb_Estatus.ListIndex < 0 Then
                    MsgBox ("Seleccione el estatus de los folios.")
                    Cmb_Estatus.SetFocus
                    Exit Sub
                End If
'                'Valida folios que no se repitan
'                If Valida_Folios("FACTURA") = True Then
                    'Valida que solo un rango este predeterminado
                    If Valida_Folios_Predeterminados("FACTURA") = True Then
                        'Valida el consecutivo predeterminado de folios
                        If Valida_Folios_Consecutivo_Predeterminados("FACTURA") = True Then
                            Modifica_Series_Folios
                        Else
                            MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                        End If
                    Else
                        MsgBox "Solo puede existir un rango predeterminado de folios, favor de verificarlo."
                    End If
'                Else
'                    MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
'                End If
            Else
                MsgBox ("Seleccione una serie/folio a modificar")
            End If
        End If
    Else
        MsgBox ("Seleccione una serie/folio a modificar")
    End If
End Sub

Private Sub Btn_Notas_Credito_Eliminar_Serie_Folios_Click()
Dim Respuesta As Integer    'Almacena la respuesta del usuario

    If Txt_Notas_Credito_Parametro_ID.text <> "" Then
        Respuesta = MsgBox("¿Esta seguro de eliminar el folio?", vbYesNo + vbCritical)
        If Respuesta = vbYes Then
            Elimina_Notas_Credito_Series_Folios
        End If
    Else
        MsgBox ("Seleccione una serie/folio a eliminar")
    End If
End Sub

Private Sub Btn_Notas_Credito_Modificar_Serie_Folios_Click()
    If Txt_Notas_Credito_Parametro_ID.text <> "" Then
       If Btn_Notas_Credito_Modificar_Serie_Folios.Caption = "Modificar" Then
            Fra_Notas_Credito_Parametros_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Nuevo_Serie_Folios.Enabled = False
            Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = False
            Btn_Notas_Credito_Modificar_Serie_Folios.Caption = "Actualizar"
            Btn_Salir.Caption = "Cancelar"
            Grid_Notas_Credito_Serie_Folios.Enabled = False
        Else
            If Txt_Notas_Credito_Parametro_ID.text <> "" Then
''                If Val(Txt_Notas_Credito_Año.text) <= 0 Then
''                    MsgBox ("Proporcione el año de los folios")
''                    Txt_Notas_Credito_Año.SetFocus
''                    Exit Sub
''                End If
''                If Val(Txt_Notas_Credito_No_Autorizacion.text) <= 0 Then
''                    MsgBox ("Proporcione el no. de autorización de los folios")
''                    Txt_Notas_Credito_No_Autorizacion.SetFocus
''                    Exit Sub
''                End If
                If Val(Txt_Notas_Credito_Folio_Inicial.text) <= 0 Then
                    MsgBox ("Proporcione el no. inicial de folio")
                    Txt_Notas_Credito_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Notas_Credito_Folio_Final.text) <= 0 Then
                    MsgBox ("Proporcione el no. final de folio")
                    Txt_Notas_Credito_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Notas_Credito_Folio_Inicial.text) > Val(Txt_Notas_Credito_Folio_Final.text) Then
                    MsgBox ("El folio inicial no puede ser mayor al folio final.")
                    Txt_Notas_Credito_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Notas_Credito_Folio_Final.text) < Val(Txt_Notas_Credito_Folio_Inicial.text) Then
                    MsgBox ("El folio final no puede ser menor al folio inicial.")
                    Txt_Notas_Credito_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Cmb_Notas_Credito_Estatus.ListIndex < 0 Then
                    MsgBox ("Seleccione el estatus de los folios.")
                    Cmb_Notas_Credito_Estatus.SetFocus
                    Exit Sub
                End If
''                'Valida folios que no se repitan
''                If Valida_Folios("NOTA CREDITO") = True Then
                    'Valida que solo un rango este predeterminado
                    If Valida_Folios_Predeterminados("NOTA_CREDITO") = True Then
                        'Valida el consecutivo predeterminado de folios
                        If Valida_Folios_Consecutivo_Predeterminados("NOTA_CREDITO") = True Then
                            Modifica_Notas_Credito_Series_Folios
                        Else
                            MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                        End If
                    Else
                        MsgBox "Solo puede existir un rango predeterminado de folios, favor de verificarlo."
                    End If
''                Else
''                    MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
''                End If
            Else
                MsgBox ("Seleccione una serie/folio a modificar")
            End If
        End If
    Else
        MsgBox ("Seleccione una serie/folio a modificar")
    End If
End Sub

Private Sub Btn_Notas_Credito_Nuevo_Serie_Folios_Click()
    If Btn_Notas_Credito_Nuevo_Serie_Folios.Caption = "Nuevo" Then
        Fra_Notas_Credito_Parametros_Serie_Folios.Enabled = True
        Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = False
        Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = False
        Btn_Notas_Credito_Nuevo_Serie_Folios.Caption = "Capturar"
        Btn_Salir.Caption = "Cancelar"
        Grid_Notas_Credito_Serie_Folios.Enabled = False
        Limpia_Notas_Credito_Controles_Serie_Folios
        Txt_Notas_Credito_Serie.SetFocus
    Else
''        If Val(Txt_Notas_Credito_Año.text) <= 0 Then
''            MsgBox ("Proporcione el año de los folios")
''            Txt_Notas_Credito_Año.SetFocus
''            Exit Sub
''        End If
''        If Val(Txt_Notas_Credito_No_Autorizacion.text) <= 0 Then
''            MsgBox ("Proporcione el no. de autorización de los folios")
''            Txt_Notas_Credito_No_Autorizacion.SetFocus
''            Exit Sub
''        End If
        If Val(Txt_Notas_Credito_Folio_Inicial.text) <= 0 Then
            MsgBox ("Proporcione el no. inicial de folio")
            Txt_Notas_Credito_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Notas_Credito_Folio_Final.text) <= 0 Then
            MsgBox ("Proporcione el no. final de folio")
            Txt_Notas_Credito_Folio_Final.SetFocus
            Exit Sub
        End If
        If Val(Txt_Notas_Credito_Folio_Inicial.text) > Val(Txt_Notas_Credito_Folio_Final.text) Then
            MsgBox ("El folio inicial no puede ser mayor al folio final.")
            Txt_Notas_Credito_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Notas_Credito_Folio_Final.text) < Val(Txt_Notas_Credito_Folio_Inicial.text) Then
            MsgBox ("El folio final no puede ser menor al folio inicial.")
            Txt_Notas_Credito_Folio_Final.SetFocus
            Exit Sub
        End If
        If Cmb_Notas_Credito_Estatus.ListIndex < 0 Then
            MsgBox ("Seleccione el estatus de los folios.")
            Cmb_Notas_Credito_Estatus.SetFocus
            Exit Sub
        End If
        'Valida folios que no se repitan
        If Valida_Folios("NOTA_CREDITO") = True Then
            'Valida que solo un rango este predeterminado
            If Valida_Folios_Predeterminados("NOTA_CREDITO") = True Then
                'Valida el consecutivo predeterminado de folios
                If Valida_Folios_Consecutivo_Predeterminados("NOTA_CREDITO") = True Then
                    Alta_Notas_Credito_Series_Folios
                Else
                    MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                End If
            Else
                MsgBox "Solo puede existir un rango activo de folios, favor de verificarlo."
            End If
        Else
            MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
        End If
    End If
End Sub

Private Sub Btn_Nuevo_Serie_Folios_Click()
    If Btn_Nuevo_Serie_Folios.Caption = "Nuevo" Then
        Fra_Parametros_Serie_Folios.Enabled = True
        Btn_Modificar_Serie_Folios.Enabled = False
        Btn_Eliminar_Serie_Folios.Enabled = False
        Btn_Nuevo_Serie_Folios.Caption = "Capturar"
        Btn_Salir.Caption = "Cancelar"
        Grid_Serie_Folios.Enabled = False
        Limpia_Controles_Serie_Folios
        Txt_Serie.SetFocus
    Else
'        If Val(Txt_Año.text) <= 0 Then
'            MsgBox ("Proporcione el año de los folios")
'            Txt_Año.SetFocus
'            Exit Sub
'        End If
'        If Val(Txt_No_Autorizacion.text) <= 0 Then
'            MsgBox ("Proporcione el no. de autorización de los folios")
'            Txt_No_Autorizacion.SetFocus
'            Exit Sub
'        End If
        If Val(Txt_Folio_Inicial.text) <= 0 Then
            MsgBox ("Proporcione el no. inicial de folio")
            Txt_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Folio_Final.text) <= 0 Then
            MsgBox ("Proporcione el no. final de folio")
            Txt_Folio_Final.SetFocus
            Exit Sub
        End If
        If Val(Txt_Folio_Inicial.text) > Val(Txt_Folio_Final.text) Then
            MsgBox ("El folio inicial no puede ser mayor al folio final")
            Txt_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Folio_Final.text) < Val(Txt_Folio_Inicial.text) Then
            MsgBox ("El folio final no puede ser menor al folio inicial")
            Txt_Folio_Final.SetFocus
            Exit Sub
        End If
        If Cmb_Estatus.ListIndex < 0 Then
            MsgBox ("Seleccione el estatus de los folios")
            Cmb_Estatus.SetFocus
            Exit Sub
        End If
        'Valida folios que no se repitan
        If Valida_Folios("FACTURA") = True Then
            'Valida que solo un rango este predeterminado
            If Valida_Folios_Predeterminados("FACTURA") = True Then
                'Valida el consecutivo predeterminado de folios
                If Valida_Folios_Consecutivo_Predeterminados("FACTURA") = True Then
                    Alta_Series_Folios
                Else
                    MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                End If
            Else
                MsgBox "Solo puede existir un rango activo de folios, favor de verificarlo."
            End If
        Else
            MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
        End If
    End If
End Sub

Private Sub Btn_Pagos_Eliminar_Serie_Folios_Click()
Dim Respuesta As Integer    'Almacena la respuesta del usuario

    If Txt_Pagos_Parametro_ID.text <> "" Then
        Respuesta = MsgBox("¿Esta seguro de eliminar el folio?", vbYesNo + vbCritical)
        If Respuesta = vbYes Then
            Elimina_Notas_Credito_Series_Folios
        End If
    Else
        MsgBox ("Seleccione una serie/folio a eliminar")
    End If
End Sub

Private Sub Btn_Pagos_Modificar_Serie_Folios_Click()
    If Txt_Pagos_Parametro_ID.text <> "" Then
       If Btn_Pagos_Modificar_Serie_Folios.Caption = "Modificar" Then
            Fra_Pagos_Parametros_Serie_Folios.Enabled = True
            Btn_Pagos_Nuevo_Serie_Folios.Enabled = False
            Btn_Pagos_Eliminar_Serie_Folios.Enabled = False
            Btn_Pagos_Modificar_Serie_Folios.Caption = "Actualizar"
            Btn_Salir.Caption = "Cancelar"
            Grid_Pagos_Serie_Folios.Enabled = False
        Else
            If Txt_Pagos_Parametro_ID.text <> "" Then
''                If Val(Txt_Pagos_Año.text) <= 0 Then
''                    MsgBox ("Proporcione el año de los folios")
''                    Txt_Pagos_Año.SetFocus
''                    Exit Sub
''                End If
''                If Val(Txt_Pagos_No_Autorizacion.text) <= 0 Then
''                    MsgBox ("Proporcione el no. de autorización de los folios")
''                    Txt_Pagos_No_Autorizacion.SetFocus
''                    Exit Sub
''                End If
                If Val(Txt_Pagos_Folio_Inicial.text) <= 0 Then
                    MsgBox ("Proporcione el no. inicial de folio")
                    Txt_Pagos_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Pagos_Folio_Final.text) <= 0 Then
                    MsgBox ("Proporcione el no. final de folio")
                    Txt_Pagos_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Pagos_Folio_Inicial.text) > Val(Txt_Pagos_Folio_Final.text) Then
                    MsgBox ("El folio inicial no puede ser mayor al folio final.")
                    Txt_Pagos_Folio_Inicial.SetFocus
                    Exit Sub
                End If
                If Val(Txt_Pagos_Folio_Final.text) < Val(Txt_Pagos_Folio_Inicial.text) Then
                    MsgBox ("El folio final no puede ser menor al folio inicial.")
                    Txt_Pagos_Folio_Final.SetFocus
                    Exit Sub
                End If
                If Cmb_Pagos_Estatus.ListIndex < 0 Then
                    MsgBox ("Seleccione el estatus de los folios.")
                    Cmb_Pagos_Estatus.SetFocus
                    Exit Sub
                End If
''                'Valida folios que no se repitan
''                If Valida_Folios("NOTA CREDITO") = True Then
                    'Valida que solo un rango este predeterminado
                    If Valida_Folios_Predeterminados("PAGOS") = True Then
                        'Valida el consecutivo predeterminado de folios
                        If Valida_Folios_Consecutivo_Predeterminados("PAGOS") = True Then
                            Modifica_Pagos_Series_Folios
                        Else
                            MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                        End If
                    Else
                        MsgBox "Solo puede existir un rango predeterminado de folios, favor de verificarlo."
                    End If
''                Else
''                    MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
''                End If
            Else
                MsgBox ("Seleccione una serie/folio a modificar")
            End If
        End If
    Else
        MsgBox ("Seleccione una serie/folio a modificar")
    End If
End Sub

Private Sub Btn_Pagos_Nuevo_Serie_Folios_Click()
    If Btn_Pagos_Nuevo_Serie_Folios.Caption = "Nuevo" Then
        Fra_Pagos_Parametros_Serie_Folios.Enabled = True
        Btn_Pagos_Modificar_Serie_Folios.Enabled = False
        Btn_Pagos_Eliminar_Serie_Folios.Enabled = False
        Btn_Pagos_Nuevo_Serie_Folios.Caption = "Capturar"
        Btn_Salir.Caption = "Cancelar"
        Grid_Pagos_Serie_Folios.Enabled = False
        Limpia_Pagos_Controles_Serie_Folios
        Txt_Pagos_Serie.SetFocus
    Else
        If Val(Txt_Pagos_Folio_Inicial.text) <= 0 Then
            MsgBox ("Proporcione el no. inicial de folio")
            Txt_Pagos_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Pagos_Folio_Final.text) <= 0 Then
            MsgBox ("Proporcione el no. final de folio")
            Txt_Pagos_Folio_Final.SetFocus
            Exit Sub
        End If
        If Val(Txt_Pagos_Folio_Inicial.text) > Val(Txt_Pagos_Folio_Final.text) Then
            MsgBox ("El folio inicial no puede ser mayor al folio final")
            Txt_Pagos_Folio_Inicial.SetFocus
            Exit Sub
        End If
        If Val(Txt_Pagos_Folio_Final.text) < Val(Txt_Pagos_Folio_Inicial.text) Then
            MsgBox ("El folio final no puede ser menor al folio inicial")
            Txt_Pagos_Folio_Final.SetFocus
            Exit Sub
        End If
        If Cmb_Pagos_Estatus.ListIndex < 0 Then
            MsgBox ("Seleccione el estatus de los folios")
            Cmb_Pagos_Estatus.SetFocus
            Exit Sub
        End If
        'Valida folios que no se repitan
        If Valida_Folios("PAGOS") = True Then
            'Valida que solo un rango este predeterminado
            If Valida_Folios_Predeterminados("PAGOS") = True Then
                'Valida el consecutivo predeterminado de folios
                If Valida_Folios_Consecutivo_Predeterminados("PAGOS") = True Then
                    Alta_Pagos_Series_Folios
                Else
                    MsgBox "No se puede predeterminar un folio mayor, si todavia hay folios anteriores disponibles."
                End If
            Else
                MsgBox "Solo puede existir un rango activo de folios, favor de verificarlo."
            End If
        Else
            MsgBox "Los folios que esta dando de alta ya estan registrados, favor de verificarlo."
        End If
    End If
End Sub


Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        If Tab_Parametros.Tab = 0 Then
            Btn_Modificar.Caption = "Modificar"
            Btn_Salir.Caption = "Salir"
            Fra_Parametros_Factura.Enabled = False
        ElseIf Tab_Parametros.Tab = 1 Then
            Fra_Parametros_Serie_Folios.Enabled = False
            Grid_Serie_Folios.Enabled = True
            Btn_Nuevo_Serie_Folios.Enabled = True
            Btn_Modificar_Serie_Folios.Enabled = True
            Btn_Eliminar_Serie_Folios.Enabled = True
            Btn_Nuevo_Serie_Folios.Caption = "Nuevo"
            Btn_Modificar_Serie_Folios.Caption = "Modificar"
            Btn_Salir.Caption = "Salir"
        ElseIf Tab_Parametros.Tab = 2 Then
            Fra_Notas_Credito_Parametros_Serie_Folios.Enabled = False
            Grid_Notas_Credito_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Nuevo_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Nuevo_Serie_Folios.Caption = "Nuevo"
            Btn_Notas_Credito_Modificar_Serie_Folios.Caption = "Modificar"
            Btn_Salir.Caption = "Salir"
        End If
    End If
End Sub

Private Sub Btn_Seleccionar_Ruta_Certificado_Click()
On Error GoTo ErrHandler
    'Set CancelError is True
    Cdl_Rutas.CancelError = True
    'Titulo de la ventana
    Cdl_Rutas.DialogTitle = "Seleccione el certificado"
    'Set flags
    Cdl_Rutas.Flags = cdlOFNHideReadOnly
    'Set filters
    Cdl_Rutas.Filter = "Certificado |*.cer|"
    'Specify default filter
    Cdl_Rutas.FilterIndex = 2
    'Display the Open dialog box
    Cdl_Rutas.ShowOpen
    'Display name of selected file
    Txt_Ruta_Certificado.text = Cdl_Rutas.fileName
    If Trim(Txt_Ruta_Certificado.text) <> "" Then
        Dtp_Vigencia_Desde.Value = CFD_Consulta_Vigencia_Desde(Txt_Ruta_Certificado.text)
        Dtp_Vigencia_Hasta.Value = CFD_Consulta_Vigencia_Hasta(Txt_Ruta_Certificado.text)
    End If
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Btn_Seleccionar_Ruta_Llave_Privada_Click()
On Error GoTo ErrHandler
    'Set CancelError is True
    Cdl_Rutas.CancelError = True
    'Titulo de la ventana
    Cdl_Rutas.DialogTitle = "Seleccione la llave privada"
    'Set flags
    Cdl_Rutas.Flags = cdlOFNHideReadOnly
    'Set filters
    Cdl_Rutas.Filter = "Llave privada |*.key|"
    'Specify default filter
    Cdl_Rutas.FilterIndex = 2
    'Display the Open dialog box
    Cdl_Rutas.ShowOpen
    'Display name of selected file
    Txt_Ruta_Llave_Privada.text = Cdl_Rutas.fileName
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Btn_Seleccionar_Ruta_NC_Click()
Dim Ruta As String
    Ruta = Selecciona_Ruta_Directorio(Me, "Indique la carpeta donde se almacenará los archivos de notas de crédito")
    If Ruta <> "" Then
        Txt_Ruta_Notas_Credito.text = Ruta
    End If
End Sub

Private Sub Btn_Seleccionar_Ruta_Pdfs_Click()
Dim Ruta As String
    Ruta = Selecciona_Ruta_Directorio(Me, "Indique la carpeta donde se almacenará los archivos PDF")
    If Ruta <> "" Then
        Txt_Ruta_Pdfs.text = Ruta
    End If
End Sub

Private Sub Btn_Seleccionar_Ruta_Xmls_Click()
Dim Ruta As String
    Ruta = Selecciona_Ruta_Directorio(Me, "Indique la carpeta donde se almacenará los archivos XML")
    If Ruta <> "" Then
        Txt_Ruta_Xmls.text = Ruta
    End If
End Sub

Private Sub Cmb_Estatus_Click()
    If Btn_Nuevo_Serie_Folios.Caption = "Capturar" Then
        If Cmb_Estatus.text = "TERMINADO" Or Cmb_Estatus.text = "CANCELADO" Then
            MsgBox "No se puede asignar este estatus porque esta haciendo una nueva captura."
            Cmb_Estatus.ListIndex = -1
        End If
    End If
End Sub

Private Sub Cmb_Notas_Credito_Estatus_Click()
    If Btn_Notas_Credito_Nuevo_Serie_Folios.Caption = "Capturar" Then
        If Cmb_Notas_Credito_Estatus.text = "TERMINADO" Or Cmb_Notas_Credito_Estatus.text = "CANCELADO" Then
            MsgBox "No se puede asignar este estatus porque esta haciendo una nueva captura."
            Cmb_Notas_Credito_Estatus.ListIndex = -1
        End If
    End If
End Sub

Private Sub Cmb_Pagos_Estatus_Click()
    If Btn_Pagos_Nuevo_Serie_Folios.Caption = "Capturar" Then
        If Cmb_Pagos_Estatus.text = "TERMINADO" Or Cmb_Pagos_Estatus.text = "CANCELADO" Then
            MsgBox "No se puede asignar este estatus porque esta haciendo una nueva captura."
            Cmb_Pagos_Estatus.ListIndex = -1
        End If
    End If
End Sub


Private Sub Cmb_Regimen_Fiscal_Click()
    Txt_Mensaje_Factura.text = Cmb_Regimen_Fiscal.text
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Tab_Parametros.Tab = 0
    Consulta_Parametros
    Consulta_Grid_Series_Folios
    Consulta_Notas_Credito_Grid_Series_Folios
    Consulta_Pagos_Grid_Series_Folios
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Codigo_Regimen_Fiscal+' '+Descripcion", "Cat_Regimen_Fiscal", Cmb_Regimen_Fiscal, 1, "Descripcion")
End Sub

Private Sub Grid_Notas_Credito_Serie_Folios_Click()
    If Grid_Notas_Credito_Serie_Folios.Rows > 1 Then
        Txt_Notas_Credito_Parametro_ID.text = Grid_Notas_Credito_Serie_Folios.TextMatrix(Grid_Notas_Credito_Serie_Folios.RowSel, 0)
        Txt_Notas_Credito_Serie.text = Grid_Notas_Credito_Serie_Folios.TextMatrix(Grid_Notas_Credito_Serie_Folios.RowSel, 1)
        Txt_Notas_Credito_Folio_Inicial.text = Grid_Notas_Credito_Serie_Folios.TextMatrix(Grid_Notas_Credito_Serie_Folios.RowSel, 2)
        Txt_Notas_Credito_Folio_Final.text = Grid_Notas_Credito_Serie_Folios.TextMatrix(Grid_Notas_Credito_Serie_Folios.RowSel, 3)
        Call Conectar_Ayudante.Asigna_Item_Combo(Grid_Notas_Credito_Serie_Folios.TextMatrix(Grid_Notas_Credito_Serie_Folios.RowSel, 4), Cmb_Notas_Credito_Estatus)
        If Cmb_Notas_Credito_Estatus.text = "ACTIVO" Or Cmb_Notas_Credito_Estatus.text = "PENDIENTE" Then
            Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = True
            Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = True
        Else
            Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = False
            Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = False
        End If
    End If
End Sub

Private Sub Grid_Pagos_Serie_Folios_Click()
    If Grid_Pagos_Serie_Folios.Rows > 1 Then
        Txt_Pagos_Parametro_ID.text = Grid_Pagos_Serie_Folios.TextMatrix(Grid_Pagos_Serie_Folios.RowSel, 0)
        Txt_Pagos_Serie.text = Grid_Pagos_Serie_Folios.TextMatrix(Grid_Pagos_Serie_Folios.RowSel, 1)
        Txt_Pagos_Folio_Inicial.text = Grid_Pagos_Serie_Folios.TextMatrix(Grid_Pagos_Serie_Folios.RowSel, 2)
        Txt_Pagos_Folio_Final.text = Grid_Pagos_Serie_Folios.TextMatrix(Grid_Pagos_Serie_Folios.RowSel, 3)
        Call Conectar_Ayudante.Asigna_Item_Combo(Grid_Pagos_Serie_Folios.TextMatrix(Grid_Pagos_Serie_Folios.RowSel, 4), Cmb_Pagos_Estatus)
        If Cmb_Pagos_Estatus.text = "ACTIVO" Or Cmb_Pagos_Estatus.text = "PENDIENTE" Then
            Btn_Pagos_Modificar_Serie_Folios.Enabled = True
            Btn_Pagos_Eliminar_Serie_Folios.Enabled = True
        Else
            Btn_Pagos_Modificar_Serie_Folios.Enabled = False
            Btn_Pagos_Eliminar_Serie_Folios.Enabled = False
        End If
    End If
End Sub


Private Sub Grid_Serie_Folios_Click()
    If Grid_Serie_Folios.Rows > 1 Then
        Txt_Parametro_ID.text = Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 0)
        Txt_Serie.text = Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 1)
        Txt_Folio_Inicial.text = Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 2)
        Txt_Folio_Final.text = Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 3)
        Call Conectar_Ayudante.Asigna_Item_Combo(Grid_Serie_Folios.TextMatrix(Grid_Serie_Folios.RowSel, 4), Cmb_Estatus)
        If Cmb_Estatus.text = "ACTIVO" Or Cmb_Estatus.text = "PENDIENTE" Then
            Btn_Modificar_Serie_Folios.Enabled = True
            Btn_Eliminar_Serie_Folios.Enabled = True
        Else
            Btn_Modificar_Serie_Folios.Enabled = False
            Btn_Eliminar_Serie_Folios.Enabled = False
        End If
    End If
End Sub

Private Sub Txt_Dias_Expira_Folios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Dias_Expira_Folios.text, False)
End Sub

Private Sub Txt_Dias_Expira_Vigencia_Certificado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Dias_Expira_Vigencia_Certificado.text, False)
End Sub

Private Sub Txt_Folio_Final_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Folio_Final.text, False)
End Sub

Private Sub Txt_Folio_Inicial_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Folio_Inicial.text, False)
End Sub

Private Sub Txt_Notas_Credito_Folio_Final_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Notas_Credito_Folio_Final.text, False)
End Sub

Private Sub Txt_Notas_Credito_Folio_Inicial_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Notas_Credito_Folio_Inicial.text, False)
End Sub

Private Sub Txt_Notas_Credito_Serie_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Letras(KeyAscii)
End Sub

Private Sub Txt_Pagos_Folio_Final_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pagos_Folio_Final.text, False)

End Sub

Private Sub Txt_Pagos_Folio_Inicial_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pagos_Folio_Inicial.text, False)

End Sub

Private Sub Txt_Pagos_Serie_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Letras(KeyAscii)
End Sub


Private Sub Txt_Porcentaje_Pagare_Interes_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Porcentaje_Pagare_Interes.text, False)
End Sub

Private Sub Txt_Serie_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Letras(KeyAscii)
End Sub


