VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Pagos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pagos"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7470
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
      Left            =   5970
      Picture         =   "Frm_Adm_Pagos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "A"
      Top             =   7170
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
      Left            =   195
      Picture         =   "Frm_Adm_Pagos.frx":36FF
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "A"
      Top             =   7170
      UseMaskColor    =   -1  'True
      Width           =   1350
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
      Height          =   615
      Left            =   150
      TabIndex        =   23
      Top             =   3435
      Width           =   7215
      Begin VB.TextBox Txt_Monto_Nota 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5250
         MaxLength       =   10
         TabIndex        =   3
         Top             =   225
         Width           =   1770
      End
      Begin VB.TextBox Txt_Nota_Credito 
         Height          =   285
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   2
         Top             =   225
         Width           =   2370
      End
      Begin VB.Label Lbl_Monto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   195
         Index           =   8
         Left            =   4275
         TabIndex        =   25
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Lbl_Nota_Credito 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Nota Credito"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   24
         Top             =   300
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab STb_Pagos 
      Height          =   2490
      Left            =   135
      TabIndex        =   14
      Top             =   900
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4392
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Pendientes de Pago"
      TabPicture(0)   =   "Frm_Adm_Pagos.frx":6C36
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Facturas_Pendientes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "A Pagar"
      TabPicture(1)   =   "Frm_Adm_Pagos.frx":6C52
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Facturas_Pagar"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Facturas_Pagar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   -74925
         TabIndex        =   17
         Top             =   375
         Width           =   7065
         Begin VB.CommandButton Btn_Consulta_Pagos_Facturas 
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
            Height          =   420
            Index           =   1
            Left            =   1455
            Picture         =   "Frm_Adm_Pagos.frx":6C6E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1560
            Width           =   1470
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
            Left            =   120
            Picture         =   "Frm_Adm_Pagos.frx":6F03
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1560
            Width           =   1260
         End
         Begin VB.TextBox Txt_Total 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5325
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1725
            Width           =   1620
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Facturas_Pagar 
            Height          =   1260
            Left            =   150
            TabIndex        =   18
            Top             =   270
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   2223
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
         Begin VB.Label Lbl_Total_Pagar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar:"
            Height          =   195
            Index           =   0
            Left            =   4050
            TabIndex        =   20
            Top             =   1800
            Width           =   1005
         End
      End
      Begin VB.Frame Fra_Facturas_Pendientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   75
         TabIndex        =   15
         Top             =   375
         Width           =   7065
         Begin VB.CommandButton Btn_Consulta_Pagos_Facturas 
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
            Height          =   420
            Index           =   0
            Left            =   1440
            Picture         =   "Frm_Adm_Pagos.frx":A1B5
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1560
            Width           =   1470
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
            Left            =   135
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Adm_Pagos.frx":A44A
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1590
            Width           =   1230
         End
         Begin VB.TextBox Txt_Saldo_Total 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   5325
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1770
            Width           =   1620
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Facturas_Pendientes 
            Height          =   1260
            Left            =   135
            TabIndex        =   16
            Top             =   240
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   2223
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Total"
            Height          =   195
            Left            =   4050
            TabIndex        =   28
            Top             =   1800
            Width           =   810
         End
      End
   End
   Begin VB.Frame Fra_Proveedor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proveedor"
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
      Left            =   135
      TabIndex        =   11
      Top             =   300
      Width           =   7215
      Begin VB.ComboBox Cmb_Proveedor 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   180
         Width           =   6015
      End
      Begin VB.Label Lbl_Nombre_Proveedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Fra_Datos_Pago 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos del Pago"
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
      Left            =   150
      TabIndex        =   1
      Top             =   4065
      Width           =   7215
      Begin VB.TextBox Txt_Dias 
         Height          =   330
         Left            =   5325
         MaxLength       =   20
         TabIndex        =   48
         Top             =   1642
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.ComboBox Cmb_Tipo_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Pagos.frx":D700
         Left            =   1650
         List            =   "Frm_Adm_Pagos.frx":D70A
         TabIndex        =   46
         Top             =   1650
         Width           =   2415
      End
      Begin VB.TextBox Txt_Comentarios 
         Height          =   270
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   39
         Top             =   2370
         Width           =   5415
      End
      Begin VB.TextBox Txt_Saldo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         Top             =   1995
         Width           =   1740
      End
      Begin VB.TextBox Txt_Total_Pesos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5355
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   525
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox Txt_Tipo_Cambio 
         Height          =   330
         Left            =   5325
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.TextBox Txt_Moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   3075
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Top             =   900
         Width           =   990
      End
      Begin VB.TextBox Txt_Concepto 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2700
         Width           =   5415
      End
      Begin VB.ComboBox Cmb_Banco 
         Height          =   315
         Left            =   1650
         TabIndex        =   8
         Top             =   150
         Width           =   5415
      End
      Begin VB.ComboBox Cmb_Forma_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Pagos.frx":D720
         Left            =   1650
         List            =   "Frm_Adm_Pagos.frx":D72D
         TabIndex        =   6
         Top             =   1275
         Width           =   2415
      End
      Begin VB.TextBox Txt_Pago 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   4
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   330
         Left            =   5325
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1260
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   315
         Left            =   1650
         TabIndex        =   9
         Top             =   525
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58261507
         CurrentDate     =   38038
      End
      Begin VB.Label Lbl_Dias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         Height          =   195
         Left            =   4185
         TabIndex        =   49
         Top             =   1710
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Lbl_Forma_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Pago"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   47
         Top             =   1710
         Width           =   1350
      End
      Begin VB.Label Lbl_Comentarios 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   38
         Top             =   2415
         Width           =   870
      End
      Begin VB.Label Lbl_Referencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   240
         Left            =   4185
         TabIndex        =   37
         Top             =   1305
         Width           =   840
      End
      Begin VB.Label Lbl_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pago                   $"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   36
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Lbl_Forma_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   35
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Lbl_Fecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   34
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Lbl_Banco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco "
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   33
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Lbl_Concepto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   32
         Top             =   2745
         Width           =   690
      End
      Begin VB.Label Lbl_Tipo_Cambio 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Cambio   $"
         Height          =   240
         Left            =   4200
         TabIndex        =   31
         Top             =   945
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Lbl_Total_Pesos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total en Pesos"
         Height          =   195
         Left            =   4200
         TabIndex        =   30
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Lbl_Saldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   2040
         Width           =   405
      End
   End
   Begin VB.Label Lbl_Pagos_Proveedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGOS A PROVEEDORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   13
      Top             =   0
      Width           =   3810
   End
End
Attribute VB_Name = "Frm_Adm_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cuenta As String
Public Formato As String
Public Leyenda As String

Private Sub Btn_Agregar_Click()
    If Grid_Facturas_Pendientes.RowSel > 0 Then
        If Txt_Moneda.Text = Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5) Or Txt_Moneda = "" Then
            If Grid_Facturas_Pagar.Rows = 0 Then
                Grid_Facturas_Pagar.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Moneda"
            End If
            Grid_Facturas_Pagar.AddItem Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0) & Chr(9) & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 1) _
            & Chr(9) & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 2) & Chr(9) & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 3) _
            & Chr(9) & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4) & Chr(9) & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5)
            Txt_Total.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4), ",")), "###,###,###.00")
            Txt_Saldo_Total.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Saldo_Total.Text, ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4), ",")), "###,###,###.00")
            Txt_Pago.Text = Txt_Total.Text
            If (Txt_Concepto.Text) = "" Then Txt_Concepto.Text = "" '"PAGO DE FACTURAS: "
            Txt_Concepto.Text = Txt_Concepto.Text '& "  " & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0)
            Txt_Moneda.Text = Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5)
            Grid_Facturas_Pagar.FixedRows = 1
            'remueve del otro grid la factura
            If Grid_Facturas_Pendientes.Rows = 2 Then
                Grid_Facturas_Pendientes.Rows = 0
            Else
                Grid_Facturas_Pendientes.RemoveItem (Grid_Facturas_Pendientes.RowSel)
            End If
        Else
            MsgBox "No se puede mezclar facturas de diferentes monedas", vbOKOnly + vbInformation, "ADMINISTRACIÓN"
        End If
    End If
End Sub

'Private Sub Btn_Capturar_Click()
'    If Btn_Capturar.Caption = "Nuevo" Then
'        STb_Pagos.Tab = 0
'        STb_Pagos.Enabled = True
'        Fra_Facturas_Pagar.Enabled = True
'        Fra_Facturas_Pendientes.Enabled = True
'        Fra_Proveedor.Enabled = True
'        Fra_Nota_Credito.Enabled = True
'        Fra_Datos_Pago.Enabled = True
'        Btn_Capturar.Caption = "Capturar"
'        Btn_Salir.Caption = "Cancelar"
'        Cmb_Proveedor.SetFocus
'    Else
'        Call Alta_Pago
'    End If
'End Sub

Private Sub Btn_Consulta_Pagos_Facturas_Click(Index As Integer)
    Load Frm_Adm_Movimientos_Consulta
    If STb_Pagos.Tab = 0 And Grid_Facturas_Pendientes.RowSel > 0 Then Frm_Adm_Movimientos_Consulta.Consulta_Movimientos (Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0))
    If STb_Pagos.Tab = 1 And Grid_Facturas_Pagar.RowSel > 0 Then Frm_Adm_Movimientos_Consulta.Consulta_Movimientos (Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 0))
End Sub

Private Sub Btn_Eliminar_Click()
Dim I As Integer                      'Variable que funciona como contador de las facturas a eliminar

    If Grid_Facturas_Pagar.RowSel > 0 Then
        If Grid_Facturas_Pendientes.Rows = 0 Then
            Grid_Facturas_Pendientes.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo"
        End If
        Grid_Facturas_Pendientes.AddItem Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 0) & Chr(9) & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 1) _
        & Chr(9) & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 2) & Chr(9) & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 3) _
        & Chr(9) & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 4) & Chr(9) & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 5)
        Txt_Total.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 4), ",")), "###,###,###.00")
        Txt_Saldo_Total.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Saldo_Total.Text, ",")) + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 4), ",")), "###,###,###.00")
        Txt_Pago.Text = Txt_Total.Text
        'remueve del otro grid la factura
        If Grid_Facturas_Pagar.Rows = 2 Then
            Grid_Facturas_Pagar.Rows = 0
            Txt_Moneda = ""
        Else
            Grid_Facturas_Pagar.RemoveItem (Grid_Facturas_Pagar.RowSel)
        End If
        Grid_Facturas_Pendientes.FixedRows = 1
        Txt_Concepto.Text = ""
        For I = 1 To Grid_Facturas_Pagar.Rows - 1
            Txt_Concepto.Text = Txt_Concepto.Text & "  " & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 0)
        Next I
        Txt_Concepto.Text = "PAGO DE FACTURAS: " & Txt_Concepto.Text
    End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Cmb_Banco.ListIndex = -1
        Txt_Total_Pesos.Text = ""
        DTP_Fecha.Value = Now
        Txt_Tipo_Cambio.Text = ""
        Cmb_Forma_Pago.ListIndex = -1
        Txt_Referencia.Text = ""
        Cmb_Tipo_Pago.ListIndex = -1
        Txt_Dias.Text = ""
        Txt_Comentarios.Text = ""
        Txt_Concepto.Text = ""
        Txt_Nota_Credito.Text = ""
        Txt_Monto_Nota.Text = ""
        Txt_Pago.Text = ""
        Txt_Saldo.Text = ""
        STb_Pagos.Tab = 0
        STb_Pagos.Enabled = True
        Fra_Facturas_Pagar.Enabled = True
        Fra_Facturas_Pendientes.Enabled = True
        Fra_Proveedor.Enabled = True
        Fra_Nota_Credito.Enabled = True
        Fra_Datos_Pago.Enabled = True
        Btn_Nuevo.Caption = "Capturar"
        Btn_Salir.Caption = "Cancelar"
        Cmb_Proveedor.SetFocus
    Else
        Call Alta_Pago
    End If
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Cmb_Proveedor.ListIndex = -1
        Cmb_Banco.ListIndex = -1
        Grid_Facturas_Pagar.Rows = 0
        Grid_Facturas_Pendientes.Rows = 0
        STb_Pagos.Tab = 0
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Caption = "Nuevo"
        Fra_Facturas_Pagar.Enabled = False
        Fra_Facturas_Pendientes.Enabled = False
        Fra_Datos_Pago.Enabled = False
        Fra_Nota_Credito.Enabled = False
        Fra_Proveedor.Enabled = False
        STb_Pagos.Enabled = False

    End If
End Sub

Private Sub Cmb_Banco_Click()
    Dim Rs_Consulta_Cat_Bancos As rdoResultset  '#  consulta el banco seleccionado
    
    If Cmb_Banco.ListIndex = -1 Then Exit Sub
    Mi_SQL = "SELECT Formato FROM Cat_Bancos " & _
    "WHERE Banco_ID = '" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "' "
    Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Bancos
        If Not .EOF Then
            Formato = .rdoColumns("Formato")
        Else
            Formato = ""
        End If
        .Close
    End With
    Set Rs_Consulta_Cat_Bancos = Nothing
End Sub

Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Forma_Pago_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Bancos As rdoResultset

    If Cmb_Forma_Pago.Text = "Cheque" Then
        Lbl_Referencia.Caption = "No. Cheque"
        If Cmb_Banco.ListIndex > -1 Then
            Mi_SQL = " SELECT * FROM Cat_Bancos  WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "'  "
            Set Rs_Consulta_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Bancos.EOF Then
                If Not IsNull(Rs_Consulta_Bancos!Numero_Inial_Cheque) Then
                    If Val(Rs_Consulta_Bancos!Numero_Inial_Cheque) > Val(Rs_Consulta_Bancos!Consecutivo_Cheque) Then
                       Txt_Referencia.Text = Val(Rs_Consulta_Bancos!Numero_Inial_Cheque)
                    Else
                       Txt_Referencia.Text = Val(Rs_Consulta_Bancos!Consecutivo_Cheque) + 1
                    End If
                Else
                    Txt_Referencia.Text = 1
                End If
            End If
            Rs_Consulta_Bancos.Close
        End If
    Else
        If Cmb_Forma_Pago.Text = "Efectivo" Then
            Lbl_Referencia.Caption = "Recibe"
        Else
            Lbl_Referencia.Caption = "Referencia"
        End If
    End If
End Sub

Private Sub Cmb_Proveedor_Click()
    Dim Rs_Consulta_Cat_Proveedores As rdoResultset     '#  consulta el proveedor
    Dim Rs_Consulta_Adm_Facturas_Proveedores As rdoResultset     '#  consulta las facturas del proveedor
    Dim Importe As Double
    
    If Cmb_Proveedor.ListIndex = -1 Then Exit Sub
    'Consulta datos del proveedor
    Mi_SQL = "SELECT Dias_Credito, Forma_Pago,Tipo_Pago,Dias_Credito " & _
    "FROM Cat_Proveedores WHERE Proveedor_ID = '" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "' "
    Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Proveedores
        '#  si encontro el dato
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Forma_Pago")) Then Cmb_Forma_Pago.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Forma_Pago"), Cmb_Forma_Pago)
            If Not IsNull(.rdoColumns("Tipo_Pago")) Then
                Cmb_Tipo_Pago.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Tipo_Pago"), Cmb_Tipo_Pago)
            Else
                Cmb_Tipo_Pago.ListIndex = -1
            End If
            If Cmb_Tipo_Pago.ListIndex = 1 Then Txt_Dias.Text = .rdoColumns("Dias_Credito")
        End If
        .Close
    End With
    Set Rs_Consulta_Cat_Proveedores = Nothing
    'Consulta los vales pendientes de facturar
    Mi_SQL = "SELECT No_Factura, Fecha, Total, Abono, Saldo, Moneda FROM Adm_Proveedores_Facturas " & _
    "WHERE Proveedor_ID = '" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "' AND Cancelada = 'N' AND Pagada = 'N' " & _
    "ORDER BY No_Factura "
    Set Rs_Consulta_Adm_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Grid_Facturas_Pendientes.Rows = 0
    Grid_Facturas_Pagar.Rows = 0
    With Rs_Consulta_Adm_Facturas_Proveedores
        Txt_Total.Text = 0
        Txt_Moneda = ""
        Txt_Saldo_Total.Text = 0
        Grid_Facturas_Pendientes.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Moneda"
        While Not .EOF
            Grid_Facturas_Pendientes.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
            & Chr(9) & Format(.rdoColumns("Total"), "###,###.00") & Chr(9) & Format(.rdoColumns("Abono"), "###,###.00") _
            & Chr(9) & Format(.rdoColumns("Saldo"), "###,###.00") & Chr(9) & .rdoColumns("Moneda")
            Txt_Saldo_Total.Text = Txt_Saldo_Total.Text + .rdoColumns("Saldo")
            .MoveNext
            Grid_Facturas_Pendientes.FixedRows = 1
        Wend
        Txt_Saldo_Total.Text = Format(Txt_Saldo_Total.Text, "###,###,###.00")
        Grid_Facturas_Pendientes.ColWidth(0) = 1000
        Grid_Facturas_Pendientes.ColAlignment(0) = 3
        Grid_Facturas_Pendientes.ColWidth(1) = 1000
        Grid_Facturas_Pendientes.ColAlignment(1) = 3
        Grid_Facturas_Pendientes.ColWidth(2) = 1200
        Grid_Facturas_Pendientes.ColWidth(3) = 1200
        Grid_Facturas_Pendientes.ColWidth(4) = 1200
        Grid_Facturas_Pendientes.ColWidth(5) = 800
        Grid_Facturas_Pagar.ColWidth(0) = 1000
        Grid_Facturas_Pagar.ColAlignment(0) = 3
        Grid_Facturas_Pagar.ColWidth(1) = 1000
        Grid_Facturas_Pagar.ColAlignment(1) = 3
        Grid_Facturas_Pagar.ColWidth(2) = 1200
        Grid_Facturas_Pagar.ColWidth(3) = 1200
        Grid_Facturas_Pagar.ColWidth(4) = 1200
        Grid_Facturas_Pagar.ColWidth(5) = 800
        STb_Pagos.Tab = 0
        .Close
    End With
    Set Rs_Consulta_Adm_Facturas_Proveedores = Nothing
End Sub

Private Sub Cmb_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Consulta_Proveedores
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub


Private Sub Cmb_Tipo_Pago_Click()
    If Cmb_Tipo_Pago.ListIndex = 1 Then
        Txt_Dias.Visible = True
        Lbl_Dias.Visible = True
    Else
        Txt_Dias.Visible = False
        Lbl_Dias.Visible = False
    End If
End Sub


Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 1000
    Me.Height = 8400
    Me.Width = 7590
    DTP_Fecha.Value = Now
    Cmb_Forma_Pago.ListIndex = 0
    Call Cmb_Proveedor_KeyPress(13)
End Sub



'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN        : Consulta_Proveedores()
'   DESCRIPCIÓN                 : Realiza una consulta a la tabla de Cat_Proveedores para llenar el combo de proveedores
'   PARÁMETROS                  : Generales de Proveedores
'   CREO                        : Julio Cruz
'   FECHA_CREO                  : 7-Oct-2010
'   MODIFICO                    :
'   FECHA_MODIFICO              :
'   CAUSA_MODIFICACIÓN          :
'*******************************************************************************
Public Sub Consulta_Proveedores()
   Cmb_Proveedor.Clear
   Call Conectar_Ayudante.Llena_Combo_Item(" DISTINCT CP.Proveedor_ID,CP.Nombre", "Cat_Proveedores CP,Adm_Proveedores_Facturas AFP ", Cmb_Proveedor, 1, " CP.Proveedor_ID=AFP.Proveedor_ID AND AFP.Saldo>0 AND Cancelada = 'N' AND Pagada = 'N' AND  CP.Nombre")  ' GROUP BY CP.Proveedor_ID,CP.Nombre")
End Sub

Private Sub Grid_Facturas_Pagar_DblClick()
    Call Btn_Eliminar_Click
End Sub

Private Sub Grid_Facturas_Pendientes_DblClick()
    Call Btn_Agregar_Click
End Sub



Private Sub Txt_Moneda_Change()
If UCase(Trim(Txt_Moneda.Text)) = UCase("Dolares") Then
    Txt_Tipo_Cambio.Visible = True
    Lbl_Tipo_Cambio.Visible = True
    Txt_Total_Pesos.Visible = True
    Lbl_Total_Pesos.Visible = True
Else
    Txt_Tipo_Cambio.Visible = False
    Lbl_Tipo_Cambio.Visible = False
    Txt_Total_Pesos.Visible = False
    Lbl_Total_Pesos.Visible = False
End If
End Sub

Private Sub Txt_Monto_Nota_Change()
    Txt_Pago.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) - Val(Txt_Monto_Nota.Text), "##.00")
End Sub

Private Sub Txt_Monto_Nota_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Monto_Nota, True)
End Sub

Private Sub Txt_Monto_Nota_LostFocus()
    Txt_Monto_Nota.Text = Format(Txt_Monto_Nota.Text, "###,##0.00")
End Sub

Private Sub Txt_Nota_Credito_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Nota_Credito, False)
End Sub

Private Sub Txt_Pago_Change()
    If Txt_Moneda = "Dolares" And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) > 0 And Val(Txt_Tipo_Cambio) > 0 Then
        Txt_Total_Pesos.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) * Val(Txt_Tipo_Cambio), "#.00")
    Else
        Txt_Total_Pesos.Text = ""
    End If
    Txt_Saldo.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) - Val(Txt_Monto_Nota.Text), "#,###,###.00")
End Sub

Private Sub Txt_Pago_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pago.Text, True)
End Sub

Private Sub Txt_Pago_LostFocus()
    Txt_Pago.Text = Format(Txt_Pago.Text, "###,###,###.00")
End Sub

Private Sub Txt_Tipo_Cambio_Change()
    If UCase(Trim(Txt_Moneda.Text)) = UCase("Dolares") And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) > 0 And Val(Txt_Tipo_Cambio) > 0 Then
        Txt_Total_Pesos.Text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) * Val(Txt_Tipo_Cambio), "#.00")
    Else
        Txt_Total_Pesos.Text = ""
    End If
End Sub

'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Alta_Pago()
'   DESCRIPCIÓN: Realiza el alta de un pago a proveedor
'   PARÁMETROS:Generales de Adm_Proveedores_Facturas, Adm_Notas_Credito y Adm_Movimientos
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO:Rafael Muñoz
'   FECHA_MODIFICO:29-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Private Sub Alta_Pago()
Dim Rs_Modifica_Adm_Facturas_Proveedores As rdoResultset        '#  Modifica las facturas
Dim Rs_Agrega_Adm_Notas_Credito_Proveedores As rdoResultset     '#  Agrega las notas de credito
Dim Rs_Agrega_Adm_Movimientos As rdoResultset                   '#  Agrega las el movimiento
Dim Importe As Double
Dim No_Movimiento As String
Dim Valor As Integer
Dim Descripcion_Cuenta As String
Dim Pago As Double
Dim Factura As String
Dim MesAño As String
Dim Total_Contabilizar As Double
Dim Total_Pago_Banco As Double
Dim Total_Pago_Proveedor As Double
Dim Iva_Factura As Double
Dim Respuesta As Integer
Dim Rs_Edita_Banco As rdoResultset
Dim cont_Repeticiones As Integer

On Error GoTo Handler
    If Cmb_Banco.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) > 0 And _
    Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) <= Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) And _
    Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) > 0 And Trim(Txt_Referencia.Text) <> "" _
    Or (Trim(Txt_Nota_Credito.Text) <> "" And Val(Txt_Monto_Nota.Text) > 0 And _
    Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) > 0) Then
        If Trim(Txt_Nota_Credito.Text) = "" And Val(Txt_Monto_Nota.Text) > 0 Then
            MsgBox "Falta el No. de Nota de credito para aplicarla", vbInformation, "Validacion"
            Exit Sub
        End If
        'INICIA LA TRANSACCION
        'DA DE ALTA LA NOTA DE CREDITO SI EXISTE
        If Trim(Txt_Nota_Credito.Text) <> "" And Val(Txt_Monto_Nota.Text) > 0 Then
            If Val(Txt_Monto_Nota.Text) > Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(1, 4), ",")) Then
                MsgBox "La nota de credito no puede ser mayor que la primer factura a pagar"
                Exit Sub
            End If
            Importe = Val(Txt_Monto_Nota.Text)
            Conexion_Base.BeginTrans
            'Actualiza la factura
            Mi_SQL = "SELECT No_Factura,Abono,Saldo,Pagada,Fecha_Pago,Usuario_Modifico,Fecha_Modifico,Total,Tipo_Cambio,IVA " & _
            "FROM Adm_Proveedores_Facturas WHERE No_Factura = '" & Grid_Facturas_Pagar.TextMatrix(1, 0) & "' " & _
            "AND Proveedor_ID = '" & Format((Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex)), "00000") & "' "
            Set Rs_Modifica_Adm_Facturas_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            With Rs_Modifica_Adm_Facturas_Proveedores
                If Not .EOF Then
                    .Edit
                        If Fix(Importe) >= Fix(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(1, 4), ","))) Then
                            .rdoColumns("Pagada") = "S"
                            .rdoColumns("Saldo") = 0 'COLUMNA 4
                            .rdoColumns("Abono") = .rdoColumns("Total") 'COLUMNA 3
                            .rdoColumns("Fecha_Pago") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                            Grid_Facturas_Pagar.TextMatrix(1, 4) = 0
                            Grid_Facturas_Pagar.TextMatrix(1, 3) = .rdoColumns("Total")
                        Else
                            .rdoColumns("Saldo") = .rdoColumns("Saldo") - Importe
                            .rdoColumns("Abono") = .rdoColumns("Abono") + Importe
                            Grid_Facturas_Pagar.TextMatrix(1, 4) = .rdoColumns("Saldo")
                            Grid_Facturas_Pagar.TextMatrix(1, 3) = .rdoColumns("Abono")
                        End If
                        .rdoColumns("Tipo_Cambio") = Val(Txt_Tipo_Cambio.Text)
                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                        .rdoColumns("Fecha_Modifico") = Now()
                    .Update
                End If
                .Close
            End With
            Set Rs_Modifica_Adm_Facturas_Proveedores = Nothing
            'DA DE ALTA LA NOTA DE CREDITO
            Set Rs_Agrega_Adm_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Agregar("Adm_Proveedores_Notas_Credito")
            With Rs_Agrega_Adm_Notas_Credito_Proveedores
                .AddNew
'                    Txt_Nota_Credito.Text = Conectar_Ayudante.Maximo_Catalogo("Adm_Proveedores_Notas_Credito WHERE Proveedor_ID='" & Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000") & "' ", "No_Nota_Credito")
                    .rdoColumns("No_Nota_Credito") = Format(Trim(Txt_Nota_Credito.Text), "0000000000")
                    .rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                    .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                    .rdoColumns("Importe") = Val(Txt_Monto_Nota.Text)
                    .rdoColumns("No_Factura") = Grid_Facturas_Pagar.TextMatrix(1, 0)
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now()
                .Update
                .Close
            End With
            Set Rs_Agrega_Adm_Notas_Credito_Proveedores = Nothing
        End If
        'BLOQUE PARA PAGOS A FACTURAS
        If Cmb_Banco.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) > 0 And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) <= Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) >= 0 And Trim(Txt_Referencia.Text) <> "" Then
            'BLoque para dar de alta los movimientos
            Set Rs_Agrega_Adm_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
            'Da de alta el movimiento
            Importe = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
            Pago = 0
            Factura = ""
            'Total_Pago= 0
            Iva_Factura = 0
            For Ciclos = 1 To Grid_Facturas_Pagar.Rows - 1
                If Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Ciclos, 4), ",")) > 0 Then
                    Factura = Grid_Facturas_Pagar.TextMatrix(Ciclos, 0)
                    No_Movimiento = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                     '#  Agrega el movimiento
                    'Actualiza la factura
                    Mi_SQL = "SELECT No_Factura,Abono,Saldo,Pagada,Fecha_Pago,Usuario_Modifico,Fecha_Modifico," & _
                    "Total,Tipo_Cambio,IVA,Tipo FROM Adm_Proveedores_Facturas " & _
                    "WHERE No_Factura = '" & Grid_Facturas_Pagar.TextMatrix(Ciclos, 0) & "' " & _
                    "AND Proveedor_ID = '" & Format((Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex)), "00000") & "' "
                    Set Rs_Modifica_Adm_Facturas_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Modifica_Adm_Facturas_Proveedores
                        If Not .EOF Then
                            .Edit
                                If .rdoColumns("IVA") > 0 Then
                                    Iva_Factura = Iva_Factura + ((.rdoColumns("Saldo") / 1.15) * 0.15)
                                End If
                                If Fix(Importe) >= Fix(Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Ciclos, 4), ","))) Then
                                    .rdoColumns("Pagada") = "S"
                                    .rdoColumns("Saldo") = 0
                                    .rdoColumns("Abono") = .rdoColumns("Total")
                                    .rdoColumns("Fecha_Pago") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                                    Importe = Importe - Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Ciclos, 4), ","))
                                    Pago = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Ciclos, 4), ","))
                                Else
                                    Pago = Importe
                                    .rdoColumns("Saldo") = .rdoColumns("Saldo") - Importe
                                    .rdoColumns("Abono") = .rdoColumns("Abono") + Importe
                                    Importe = 0
'                                    Ciclos = Grid_Facturas_Pagar.Rows
                                End If
                                .rdoColumns("Tipo_Cambio") = Val(Txt_Tipo_Cambio.Text)
                                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                .rdoColumns("Fecha_Modifico") = Now()
                            .Update
                        End If
'                        .Close
                    End With
                    'Set Rs_Modifica_Adm_Facturas_Proveedores = Nothing
                    Set Rs_Agrega_Adm_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
                        
                    With Rs_Agrega_Adm_Movimientos
                        .AddNew
                            .rdoColumns("No_Movimiento") = No_Movimiento
                            .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                            .rdoColumns("Banco_ID") = Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000")
                            .rdoColumns("Banco") = Cmb_Banco.Text
                            .rdoColumns("No_Factura") = Factura
                            .rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                            .rdoColumns("Estatus") = "A"
                            If Format(Pago, "###,###.00") >= Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Ciclos, 4), ",")) Then
                                .rdoColumns("Concepto") = UCase("PAGO")
                            Else
                                .rdoColumns("Concepto") = UCase("ANTICIPO")
                            End If
                            .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.Text
                            .rdoColumns("Tipo_Pago") = Cmb_Tipo_Pago.Text
                            .rdoColumns("Dias_Pago") = Val(Txt_Dias.Text)
                            .rdoColumns("Referencia") = Trim(Txt_Referencia.Text)
                            .rdoColumns("Tipo") = "E"
                            If Trim(UCase(Txt_Moneda.Text)) = Trim(UCase("Pesos")) Then
                                .rdoColumns("Cantidad") = Pago
                            Else
                                .rdoColumns("Cantidad") = Pago * Val(Txt_Tipo_Cambio.Text)
                            End If
                            If Cmb_Forma_Pago.Text = "Cheque" Then
                                .rdoColumns("Beneficiario") = Cmb_Proveedor.Text
                            End If
                            .rdoColumns("Comentarios") = Txt_Comentarios.Text
                            .rdoColumns("Usuario_Creo") = Nombre_Usuario
                            .rdoColumns("Fecha_Creo") = Now()
                        .Update
                    End With
                    Rs_Agrega_Adm_Movimientos.Close
                End If
            Next Ciclos
        End If
        Set Rs_Agrega_Adm_Movimientos = Nothing
        If Cmb_Forma_Pago.Text = "Cheque" Then
            'SE EDITA EL CONSECUTIVO DEL CHEQUE
            Mi_SQL = " SELECT * FROM Cat_Bancos WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "' "
            Set Rs_Edita_Banco = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            With Rs_Edita_Banco
                If Not Rs_Edita_Banco.EOF Then
                    .Edit
                        .rdoColumns("Consecutivo_Cheque") = Txt_Referencia.Text
                    .Update
                End If
            End With
            Rs_Edita_Banco.Close
        End If
        '*****************************************************************
        '**********GENERA LA POLIZA CONTABLE*****************************
'        If Txt_Cuenta_Proveedor.Text <> "" And Txt_Cuenta_Banco.Text <> "" And Cuenta_IVA_Sujeto <> "" And Cuenta_IVA_Acreditado <> "" Then
'            'Determina el IVA de la poliza
'            If Grid_Facturas_Pagar.Rows = 2 Then
'                If Iva_Factura > 0 Then
'                    If Val(Txt_Total.Text) > Val(Txt_Pago.Text) Then
'                        Iva_Factura = Val(Format((Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) / 1.15) * 0.15, "#.00"))
'                    End If
'                End If
'            End If
'            If Val(Txt_Tipo_Cambio.Text) = 0 Then Txt_Tipo_Cambio.Text = 1
'            Total_Pago = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) * Val(Txt_Tipo_Cambio.Text), "#######.00")
'            'Busca el consecutivo de la poliza
'            With MiConsulta8
'                Set .ActiveConnection = cn
'                .SQL = "SELECT MAX(No_Poliza) FROM Cont_Polizas WHERE Tipo = 'Eg' AND "
'                .SQL = .SQL & " MesAño = '" & Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 1, 2) & Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 9, 2) & "'"
'                .RowsetSize = 1
'                .LockType = rdConcurRowVer
'                .CursorType = rdUseOdbc
'                Set MiConsecutivo = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
'            End With
'            If Not IsNull(MiConsecutivo(0)) Then
'                Num_Poliza = Format(MiConsecutivo(0) + 1, "00000")
'            Else
'                Num_Poliza = "00001"
'            End If
'            'Asigna el numero de cheque como no de poliza
'            Num_Poliza = Trim(Txt_Referencia.Text)
'            Cont_Partidas = 0
'            MesAño = Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 1, 2) & Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 9, 2)
'            MiConsecutivo.Close
'            '-------------------------------------------------------------------
'            'Da de alta la poliza de cabecera
'            With MiConsulta6
'                Set .ActiveConnection = cn
'                .SQL = "SELECT *"
'                .SQL = .SQL & " FROM Cont_Polizas"
'                .RowsetSize = 1
'                .LockType = rdConcurRowVer
'                .CursorType = rdUseOdbc
'                Set MiPoliza = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
'            End With
'            MiPoliza.AddNew  ' Create new record.
'                MiPoliza("No_Poliza") = Num_Poliza
'                MiPoliza("Tipo") = "Eg"
'                MiPoliza("MesAño") = MesAño
'                MiPoliza("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'                MiPoliza("Concepto") = Mid(Cmb_Proveedor.Text, 8)
'                MiPoliza("Total_Debe") = Total_Pago + Iva_Factura
'                MiPoliza("Total_Haber") = Total_Pago + Iva_Factura
'                MiPoliza("No_Partidas") = 4
'                MiPoliza("No_Movimiento") = No_Movimiento
'                MiPoliza("Usuario_Creo") = Usuario_Sistema
'                MiPoliza("Fecha_Creo") = Now
'            MiPoliza.Update  ' Save changes.
'            MiPoliza.Close
'            '---------------------------------------------------------
'            'Da de alta los detalles de poliza
'            With MiConsulta7
'                Set .ActiveConnection = cn
'                .SQL = "SELECT *"
'                .SQL = .SQL & " FROM Cont_Polizas_Detalles"
'                .RowsetSize = 1
'                .LockType = rdConcurRowVer
'                .CursorType = rdUseOdbc
'                Set MisDetalles = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
'            End With
'
'           'Da de alta la partida del proveedor
'            Cont_Partidas = Cont_Partidas + 1
'            MisDetalles.AddNew  ' Create new record.
'            MisDetalles("No_Poliza") = Num_Poliza
'            MisDetalles("Tipo") = "Eg"
'            MisDetalles("Partida") = Cont_Partidas
'            MisDetalles("MesAño") = MesAño
'            MisDetalles("Cuenta") = Txt_Cuenta_Proveedor.Text
'            MisDetalles("Concepto") = Mid(Cmb_Proveedor.Text, 8)
'            MisDetalles("Debe") = Total_Pago
'            MisDetalles("Haber") = 0
'            MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'            MisDetalles("Referencia") = Trim(Mid(Txt_Concepto.Text, 19))
'            MisDetalles.Update  ' Save changes.
'            If Iva_Factura > 0 Then
'                'DA DE ALTA LA PARTIDA LA IVA ACREDITABLE
'                 Cont_Partidas = Cont_Partidas + 1
'                 MisDetalles.AddNew  ' Create new record.
'                 MisDetalles("No_Poliza") = Num_Poliza
'                 MisDetalles("Tipo") = "Eg"
'                 MisDetalles("Partida") = Cont_Partidas
'                 MisDetalles("MesAño") = MesAño
'                 MisDetalles("Cuenta") = Cuenta_IVA_Acreditado
'                 MisDetalles("Concepto") = Mid(Cmb_Proveedor.Text, 8)
'                 MisDetalles("Debe") = Format(Iva_Factura, "########.00")
'                 MisDetalles("Haber") = 0
'                 MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'                 MisDetalles("Referencia") = Trim(Mid(Txt_Concepto.Text, 19))
'                 MisDetalles.Update  ' Save changes.
'                 'DA DE ALTA LA PARTIDA DEL IVA SUJETO A ACREDITAR
'                 Cont_Partidas = Cont_Partidas + 1
'                 MisDetalles.AddNew  ' Create new record.
'                 MisDetalles("No_Poliza") = Num_Poliza
'                 MisDetalles("Tipo") = "Eg"
'                 MisDetalles("Partida") = Cont_Partidas
'                 MisDetalles("MesAño") = MesAño
'                 MisDetalles("Cuenta") = Cuenta_IVA_Sujeto
'                 MisDetalles("Concepto") = Mid(Cmb_Proveedor.Text, 8)
'                 MisDetalles("Debe") = 0
'                 MisDetalles("Haber") = Format(Iva_Factura, "########.00")
'                 MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'                 MisDetalles("Referencia") = Trim(Mid(Txt_Concepto.Text, 19))
'                 MisDetalles.Update  ' Save changes
'            End If
'            'DA DE ALTA LA PARTIDA DE LOS BANCOS
'            Cont_Partidas = Cont_Partidas + 1
'            MisDetalles.AddNew  ' Create new record.
'            MisDetalles("No_Poliza") = Num_Poliza
'            MisDetalles("Tipo") = "Eg"
'            MisDetalles("Partida") = Cont_Partidas
'            MisDetalles("MesAño") = MesAño
'            MisDetalles("Cuenta") = Txt_Cuenta_Banco.Text
'            MisDetalles("Concepto") = Mid(Cmb_Proveedor.Text, 8)
'            MisDetalles("Debe") = 0
'            MisDetalles("Haber") = Total_Pago
'            MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'            MisDetalles("Referencia") = Trim(Mid(Txt_Concepto.Text, 19))
'            MisDetalles.Update  ' Save changes.
'            MisDetalles.Close
'        End If
        Conexion_Base.CommitTrans
            
        If Cmb_Banco.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) > 0 And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) <= Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.Text, ",")) And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) > 0 And Trim(Txt_Referencia.Text) <> "" Then
            If Cmb_Forma_Pago.Text = "Cheque" Then
                Respuesta = MsgBox("¿Desea Imprimir el Cheque?", vbYesNo + vbQuestion, " ADMINISTRACIÓN")
                If Respuesta = 6 Then
                    If Formato <> "" Then
                        ''Call Imprime_Cheque("Num_Poliza", "Eg", "MesAño")
                        For cont_Repeticiones = 1 To 2
                            Call Imprime_Cheques
                        Next
                    Else
                        MsgBox "No se encuentra el formato de impresion del cheque", vbExclamation, "ADMINISTRACIÓN"
                        Exit Sub
                    End If
                End If
            End If
        End If
        STb_Pagos.Tab = 1
        STb_Pagos.Enabled = False
        'Btn_Nuevo.Visible = True
        Fra_Proveedor.Enabled = False
        Fra_Nota_Credito.Enabled = False
        Fra_Datos_Pago.Enabled = False
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        MsgBox "Pago Capturado", vbInformation, "ADMINISTRACION"
    Else
        MsgBox "Datos incompletos para realizar el pago", vbExclamation, "ADMINISTRACIÓN"
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Private Sub Txt_Tipo_Cambio_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Tipo_Cambio.Text, True)
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprime_Cheques
'DESCRIPCIÓN            : Imprime Cheque del banco
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 7-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Imprime_Cheques()
Dim Rs_Consulta_Formatos As rdoResultset                        'Manejo de Registro
Dim Rs_Consulta_Formatos_General As rdoResultset                'Manejo de Registro
Dim Rs_Consulta_Formatos_Detalles As rdoResultset               'Manejo de Registro
Dim Longitud As Integer
Dim Salto As Double
Dim Fuente As Double

On Error GoTo Handler

    'Consulta los fosrmatos
    Mi_SQL = "SELECT * FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE Nombre='" & Formato & "'"
    Set Rs_Consulta_Formatos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Formatos.EOF Then
        Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Nombre='" & Formato & "'"
        Mi_SQL = Mi_SQL & " AND Tipo='General'"
        Set Rs_Consulta_Formatos_General = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Nombre='" & Formato & "'"
        Mi_SQL = Mi_SQL & " AND Tipo='Detalle'"
        Set Rs_Consulta_Formatos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Formatos
            'Comienza la impresion
            Printer.ScaleMode = vbCentimeters
            'Configura la fuente
            Printer.FontSize = .rdoColumns("Tamaño_Generales")
            Printer.Font = .rdoColumns("Letra_Generales")
            Fuente = .rdoColumns("Tamaño_Generales")
            If .rdoColumns("Estilo_Generales") = "Negrita" Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
            Salto = .rdoColumns("Separacion_Detalles")
        End With
        'Inicia la impresión
        With Rs_Consulta_Formatos_General
            If Cmb_Banco.ListIndex > -1 Then
            End If
            While Not .EOF
                Printer.CurrentX = .rdoColumns("X")
                Printer.CurrentY = .rdoColumns("Y")
                Longitud = .rdoColumns("Longitud")
                If .rdoColumns("Campo") = "FECHA" Then Printer.Print Format(DTP_Fecha.Value, "dd-MMM-yyyy")
                If .rdoColumns("Campo") = "NOMBRE" Then Printer.Print Mid(Cmb_Proveedor.Text, 1, Longitud)
                'CAMPOS PARA LA IMPRESION EN EL AREA DE POLIZA
                If .rdoColumns("Campo") = "NOMBRE_POLIZA" Then Printer.Print Mid(Cmb_Proveedor.Text, 1, Longitud)
                If .rdoColumns("Campo") = "CONCEPTO" Then Printer.Print Mid(Txt_Concepto.Text, 1, Longitud)
                If .rdoColumns("Campo") = "NUMERO_CHEQUE" Then Printer.Print Mid("CH. " & Txt_Referencia.Text, 1, Longitud)
                If .rdoColumns("Campo") = "NOMBRE_BANCO" Then Printer.Print Mid(Cmb_Banco.Text, 1, Longitud)
                'VALIDA SI LA MONEDA ES EN PESOS
                If UCase(Trim(Txt_Moneda.Text)) = "PESOS" Then
                    If .rdoColumns("Campo") = "CANTIDAD" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                    If .rdoColumns("Campo") = "CANTIDAD_POLIZA" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                    If .rdoColumns("Campo") = "CANTIDAD_LETRA" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
                Else
                    If .rdoColumns("Campo") = "CANTIDAD" Then Printer.Print Format(Txt_Total_Pesos, "###,###,###.00")
                    If .rdoColumns("Campo") = "CANTIDAD_LETRA" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total_Pesos.Text, ","))
                End If
                .MoveNext
            Wend
        End With
        Printer.EndDoc
    End If
    Rs_Consulta_Formatos.Close
    Rs_Consulta_Formatos_General.Close
    Rs_Consulta_Formatos_Detalles.Close
    Exit Sub
Handler:
    Printer.EndDoc
    MsgBox Err.Description, vbCritical
End Sub

