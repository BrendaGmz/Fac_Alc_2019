VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Cobranza 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cobranza"
   ClientHeight    =   8520
   ClientLeft      =   -120
   ClientTop       =   165
   ClientWidth     =   8850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   8850
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Documento"
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
      Left            =   240
      TabIndex        =   39
      Top             =   360
      Width           =   1575
      Begin VB.Label Lbl_Folio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.CommandButton Btn_Capturar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capturar"
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
      Picture         =   "Frm_Adm_Cobranza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "A"
      Top             =   7800
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
      Left            =   7440
      Picture         =   "Frm_Adm_Cobranza.frx":3537
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "A"
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Facturas_Pendientes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas Pendientes de Pago"
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
      TabIndex        =   8
      Top             =   1080
      Width           =   8655
      Begin VB.CommandButton Btn_Buscar_Factura 
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
         Height          =   420
         Left            =   6240
         Picture         =   "Frm_Adm_Cobranza.frx":6C36
         Style           =   1  'Graphical
         TabIndex        =   52
         Tag             =   "C"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox Txt_Buscar_Factura 
         Height          =   405
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   51
         Top             =   1200
         Width           =   1200
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
         Left            =   7200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frm_Adm_Cobranza.frx":A1C2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Facturas_Pendientes 
         Height          =   885
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1561
         _Version        =   393216
         Rows            =   0
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Fra_Facturas_Pagar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   8655
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
         Picture         =   "Frm_Adm_Cobranza.frx":D478
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   1500
      End
      Begin VB.TextBox Txt_Total 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1245
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Facturas_Pagar 
         Height          =   885
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1561
         _Version        =   393216
         Rows            =   0
         Cols            =   13
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Lbl_Total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar:"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   1290
         Width           =   1005
      End
   End
   Begin VB.Frame Fra_Cliente 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   6855
      Begin VB.ComboBox Cmb_Cliente 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   180
         Width           =   6015
      End
      Begin VB.Label Lbl_Nombre_Clientes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Fra_Datos_Pagos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos del Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   8655
      Begin VB.ComboBox Cmb_FacRef 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":1072A
         Left            =   7440
         List            =   "Frm_Adm_Cobranza.frx":1072C
         TabIndex        =   49
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox Cmb_Serie 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":1072E
         Left            =   6720
         List            =   "Frm_Adm_Cobranza.frx":10730
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2760
         Width           =   735
      End
      Begin VB.ComboBox Cmb_Relacionados 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":10732
         Left            =   1590
         List            =   "Frm_Adm_Cobranza.frx":1073C
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2760
         Width           =   2280
      End
      Begin VB.ComboBox Cmb_Banco_Des 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":10766
         Left            =   6720
         List            =   "Frm_Adm_Cobranza.frx":10768
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1680
         Width           =   1830
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   7695
         TabIndex        =   43
         Text            =   "T. Cambio"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Cmb_Tipo_Moneda 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":1076A
         Left            =   6720
         List            =   "Frm_Adm_Cobranza.frx":10774
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Txt_Cuenta_Destino 
         Height          =   285
         Left            =   1590
         TabIndex        =   38
         Top             =   1680
         Width           =   2280
      End
      Begin VB.TextBox Txt_Cuenta_Origen 
         Height          =   285
         Left            =   1590
         TabIndex        =   37
         Top             =   1320
         Width           =   2280
      End
      Begin VB.ComboBox Cmb_Tipo_Cadena 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":10782
         Left            =   1590
         List            =   "Frm_Adm_Cobranza.frx":1078C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2040
         Width           =   2280
      End
      Begin VB.TextBox Txt_Certificado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1590
         TabIndex        =   29
         Top             =   2400
         Width           =   2280
      End
      Begin VB.TextBox Txt_Cadena 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Top             =   2400
         Width           =   1830
      End
      Begin VB.TextBox Txt_Sello 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   27
         Top             =   2040
         Width           =   1830
      End
      Begin VB.ComboBox Cmb_Banco 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":107A7
         Left            =   6720
         List            =   "Frm_Adm_Cobranza.frx":107A9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox Txt_Comentarios 
         Height          =   285
         Left            =   1590
         TabIndex        =   21
         Top             =   960
         Width           =   2280
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   285
         Left            =   6720
         MaxLength       =   20
         TabIndex        =   6
         Top             =   990
         Width           =   1830
      End
      Begin VB.TextBox Txt_Pago 
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   285
         Width           =   1830
      End
      Begin VB.ComboBox Cmb_Forma_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Cobranza.frx":107AB
         Left            =   1590
         List            =   "Frm_Adm_Cobranza.frx":107B8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   2280
      End
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   270
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   110428163
         CurrentDate     =   38038
      End
      Begin VB.Label Lbl_UUID_Relacion 
         BackColor       =   &H80000005&
         Caption         =   "Relacionado"
         Height          =   255
         Left            =   5310
         TabIndex        =   50
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Tipo Relación"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Destino"
         Height          =   195
         Left            =   5310
         TabIndex        =   45
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Tipo Moneda"
         Height          =   255
         Left            =   5280
         TabIndex        =   42
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Cuenta origen"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Cuenta destino"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Lbl_Tipo_Cadena_Pago 
         BackColor       =   &H80000005&
         Caption         =   "Tipo Cadena"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Caption         =   "Certificado"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Cadena"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5310
         TabIndex        =   32
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Caption         =   "Sello"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5310
         TabIndex        =   31
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Lbl_Comentarios 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label Lbl_Banco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Origen"
         Height          =   195
         Left            =   5310
         TabIndex        =   19
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Lbl_Referencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   5310
         TabIndex        =   17
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label Lbl_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pago               $"
         Height          =   195
         Index           =   4
         Left            =   5310
         TabIndex        =   16
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Lbl_Forma_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Lbl_Fecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Label Lbl_Cobranza_Clientes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPLEMENTO PARA RECEPCIÓN DE PAGOS"
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
      Left            =   1005
      TabIndex        =   18
      Top             =   60
      Width           =   7695
   End
End
Attribute VB_Name = "Frm_Adm_Cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cliente_ID As String
Dim Tipo As String
Dim Id_cliente As String

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Btn_Agregar_Click()
'DESCRIPCIÓN: Agrega al grid los datos de las facturas que están pendientes por pagar
'PARÁMETROS:
'CREO:  Sergio Godínez Banda
'FECHA_CREO:    8-Agosto-2007
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Agregar_Click()
Dim Cont_Facturas As Integer                         'Usada para contar las facturas del grid
Dim Suma As Double                                  'Usada para sumar el saldo de la factura
Dim Pago As String

    If Cmb_Cliente.ListIndex > -1 Then
        If Grid_Facturas_Pendientes.RowSel > 0 Then
           Pago = Format(Conectar_Ayudante.Quitar_Caracter(InputBox("Ingrese el monto del pago." & vbNewLine), ","), "#.00")
            If Val(Pago) <= 0 Then
                MsgBox "Introduzca el monto del pago a realizar"
                Exit Sub
            End If
            If Val(Pago) > Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5), ",")) Then
                MsgBox "El pago debe ser menor o igual al saldo de la factura."
                Exit Sub
            End If
            If Grid_Facturas_Pagar.Rows = 0 Then
                Grid_Facturas_Pagar.AddItem "Documento" & Chr(9) & "Electrónica" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) _
                    & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Tipo" & Chr(9) & "Fecha Pago" & Chr(9) & "UUID" & Chr(9) & "Parcialidad" & Chr(9) & "Pago" & Chr(9) & "Saldo" & Chr(9) & "Serie_fact"
            End If
            If Grid_Facturas_Pagar.Rows > 1 Then
                If Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 6) = Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.Rows - 1, 6) Then
                    If Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0) <> Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.Rows - 1, 0) Then
                         Grid_Facturas_Pagar.AddItem Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 1) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 2) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 3) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 6) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 7) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 8) & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 9) & Chr(9) _
                        & Format(Val(Pago), "#0.00") & Chr(9) & Format(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5) - Val(Pago), "#0.00") & Chr(9) _
                        & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 10) & Chr(9)
                    Else
                        MsgBox "Factura agregada al pago", vbInformation
                        Exit Sub
                    End If
                Else
                    'Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4) & Chr(9)
                    MsgBox "Favor de seleccionar documentos del mismo tipo", vbInformation
                    Exit Sub
                End If
            Else
            Grid_Facturas_Pagar.AddItem Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 0) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 1) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 2) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 3) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 4) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 6) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 7) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 8) & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 9) & Chr(9) _
                & Format(Pago, "#0.00") & Chr(9) & Format(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5) - Pago, "#0.00") & Chr(9) _
                & Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 10) & Chr(9)
            End If
            'Calcula el saldo total
            Suma = 0
            Suma = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pendientes.TextMatrix(Grid_Facturas_Pendientes.RowSel, 5), ","))
            'Agrega totales a Total y pago
            Txt_Total.text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ",")) + Val(Pago), "###,###.00")  'Suma
            Txt_Total.text = Format(Txt_Total.text, "###,###.00")
'            Txt_Pago.text = Format(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","), "###,###.00")
            Txt_Pago = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ",")) + Val(Pago), "###,###.00")
                            
            'remueve del otro grid la factura
             If Grid_Facturas_Pendientes.Rows = 2 Then
                Grid_Facturas_Pendientes.Rows = 0
             Else
                Grid_Facturas_Pendientes.RemoveItem (Grid_Facturas_Pendientes.RowSel)
             End If
        End If
        Tipo = Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.Rows - 1, 6)
        Formatea_Columnas_Grid
    Else
        Grid_Facturas_Pendientes.Rows = 2
    End If
End Sub

Private Sub Btn_Buscar_Factura_Click()
Dim Mi_SQL As String
Dim Rs_MiTabla As rdoResultset
Dim Importe As Double
Dim Tipo_Documento As String
Dim Electronico As String
Dim Parcialidad As Long
Dim encontro As Boolean
        If Cmb_Cliente.ListIndex > -1 Then
            If Txt_Buscar_Factura.text <> "" Then
                Grid_Facturas_Pendientes.Rows = 0
                Grid_Facturas_Pendientes.AddItem "Documento" & Chr(9) & "Electrónica" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & _
                Chr(9) & "Saldo" & Chr(9) & "Tipo" & Chr(9) & "Fecha Pago" & Chr(9) & "UUID" & Chr(9) & "Parcialidad"

                Mi_SQL = "SELECT No_Factura, No_Factura_Electronica, Fecha, Total, Abono, Saldo, Tipo_Pago, Tipo_Documento, Fecha_Pago,Timbre_UUID,No_Parcialidad,forma_pago, serie"
                Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas "
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND (No_Factura='" & Format(Txt_Buscar_Factura, "0000000000") & "' OR "
                Mi_SQL = Mi_SQL & " No_Factura_Electronica='" & Format(Txt_Buscar_Factura, "0000000000") & "') "
                Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
                Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
                Mi_SQL = Mi_SQL & " ORDER BY No_Factura"
                Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_MiTabla
                    If Not .EOF Then
'                        Txt_Total.text = 0
                        encontro = True
                        While Not .EOF
                            If Not IsNull(.rdoColumns("Tipo_Documento")) Then
                                Tipo_Documento = .rdoColumns("Tipo_Documento")
                            Else
                                Tipo_Documento = "FACTURA"
                            End If
                            If Not IsNull(.rdoColumns("No_Factura_Electronica")) Then
                                Electronico = Val(.rdoColumns("No_Factura_Electronica"))
                            Else
                                Electronico = ""
                            End If
                            If Not IsNull(.rdoColumns("No_Parcialidad")) Then
                                Parcialidad = .rdoColumns("No_Parcialidad") + 1
                            Else
                                Parcialidad = 1
                            End If
                             Grid_Facturas_Pendientes.AddItem .rdoColumns("No_Factura") & Chr(9) & Electronico & Chr(9) _
                            & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "###,##0.00") & Chr(9) _
                            & Format(.rdoColumns("Abono"), "###,##0.00") & Chr(9) & Format(.rdoColumns("Saldo"), "###,##0.00") & Chr(9) _
                            & Tipo_Documento & Chr(9) & .rdoColumns("Fecha_Pago") & Chr(9) & .rdoColumns("Timbre_UUID") & Chr(9) & Parcialidad & Chr(9) & .rdoColumns("Serie")
                            .MoveNext
                            Grid_Facturas_Pendientes.FixedRows = 1
                        Wend
                    Else
                        encontro = False
                    End If
                End With
                Rs_MiTabla.Close
                
                'Consulta En la tabla de Remisiones
                Mi_SQL = "SELECT No_Remision, Fecha, Total, Abono, Saldo, Tipo_Pago,Tipo_Documento,Fecha_Pago"
                Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones "
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND No_Remision='" & Format(Txt_Buscar_Factura, "0000000000") & "'"
                Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
                Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
                Mi_SQL = Mi_SQL & " ORDER BY No_Remision "
                Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    With Rs_MiTabla
                        If Not .EOF Then
                            encontro = True
'                            Txt_Total.text = 0
                            ''Grid_Facturas_Pendientes.AddItem "Documento" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Tipo"
                            While Not .EOF
                                Grid_Facturas_Pendientes.AddItem .rdoColumns("No_Remision") & Chr(9) & "" & Chr(9) _
                                    & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "###,##0.00") & Chr(9) _
                                    & Format(.rdoColumns("Abono"), "###,##0.00") & Chr(9) & Format(.rdoColumns("Saldo"), "###,##0.00") & Chr(9) _
                                    & .rdoColumns("Tipo_Documento") & Chr(9) & .rdoColumns("Fecha_Pago") & Chr(9) & ""
                                .MoveNext
                                Grid_Facturas_Pendientes.FixedRows = 1
                            Wend
                        End If
                    End With
                Rs_MiTabla.Close
                If encontro = False Then
                    MsgBox "No se encontró la factura"
                    Cmb_Cliente_Click
                    Exit Sub
                End If
                Formatea_Columnas_Grid
            Else
                MsgBox "Teclee el número de factura a buscar"
                Txt_Buscar_Factura.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Seleccione un cliente"
            Exit Sub
        End If
        
       
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Btn_Capturar_Click()
'DESCRIPCIÓN            : Almacena en la tabla Adm_Factura_Clientes los movimientos realizados
'                         a las facturas
'PARÁMETROS             :
'CREO                   :  Julio Cruz
'FECHA_CREO             :  18-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Btn_Capturar_Click()
Dim Mi_SQL As String
Dim Rs_MiTabla As rdoResultset
Dim Rs_MiFactura As rdoResultset
Dim Rs_Movimientos As rdoResultset
Dim Rs_Clientes_Facturas_Detalles As rdoResultset
Dim Importe As Double
Dim No_Movimiento As Double
Dim Resultado As Integer
Dim Cargo As Double
Dim Saldo_Factura As Double
     
     'Consulta la serie del rango activo actual
'    Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
'    Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
'    Mi_SQL = Mi_SQL & " AND Tipo = 'PAGOS'"
'    Set Rs_Consulta_Serie = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'        If Not Rs_Consulta_Serie.EOF Then
'            CFD_Generales.Serie = Trim(Rs_Consulta_Serie.rdoColumns("Serie"))
'        End If
'    Rs_Consulta_Serie.Close
'    'Valida si aun existen folios de facturas disponibles para utilizar
'    Call Aviso_Termino_Folios("PAGOS")
    'si la bandera esta habilitada muestra mensaje y cancela la operación
'    If Folios_Terminados = True Then
'        CFD_Generales.Folio = ""
'        MsgBox "No se encontraron folios de factura disponibles, favor de verificar", vbCritical
'        Btn_Salir_Click
'        Exit Sub
'    End If
    If Cmb_Cliente.ListIndex = -1 Then
       MsgBox "Selecciona el cliente.", vbExclamation
       Exit Sub
       Cmb_Cliente.SetFocus
    End If
     If Grid_Facturas_Pagar.Rows = 0 Then
        MsgBox "Favor de agregar las facturas a pagar.", vbExclamation
        Exit Sub
    End If
    If Cmb_Tipo_Cadena.ListIndex > 0 Then
        If Txt_Certificado.text = "" Then
            MsgBox "Falta certificado de pago, favor de verificar", vbExclamation
            Txt_Certificado.SetFocus
            Exit Sub
        End If
        If Txt_Cadena.text = "" Then
            MsgBox "Falta cadena de pago, favor de verificar", vbExclamation
            Txt_Cadena.SetFocus
            Exit Sub
        End If
        If Txt_Sello.text = "" Then
            MsgBox "Falta sello de pago, favor de verificar", vbExclamation
            Txt_Sello.SetFocus
            Exit Sub
        End If
    End If
    If Cmb_Relacionados.ListIndex > 0 Then
        If Cmb_Serie.ListIndex = -1 Then
            MsgBox "Selecciona la serie del documento relacionado"
            Cmb_Serie.SetFocus
            Exit Sub
        End If
        If Cmb_FacRef.text = "" Then
            MsgBox "Selecciona el folio del documento relacionado"
            Cmb_FacRef.SetFocus
            Exit Sub
        Else
            Busca_UUID
        End If
    End If
    If Cmb_Tipo_Moneda.ListIndex = -1 Then
        MsgBox "Selecciona el tipo de moneda"
        Cmb_Tipo_Moneda.SetFocus
        Exit Sub
    End If
    If Cmb_Tipo_Moneda.text <> "MXN" And (Text2.text = "" Or Text2.text = "T. Cambio") Then
        MsgBox "Selecciona el tipo de cambio"
        Cmb_Tipo_Moneda.SetFocus
        Exit Sub
    End If
    Set Conectar_Ayudante = New Ayudante
        On Error GoTo handler
        If CDbl(Txt_Total.text) > 0 And CDbl(Txt_Pago.text) <= CDbl(Txt_Total.text) _
            And CDbl(Txt_Pago.text) > 0 Then
            Resultado = MsgBox("Seguro de capturar el cobro", vbQuestion + vbYesNo, "MAGICAL BRIDE")
            If Resultado = 6 Then
                Conexion_Base.BeginTrans
                'Actualiza las facturas
                Importe = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ","))
                Cargo = 0
                For I = 1 To Grid_Facturas_Pagar.Rows - 1
                    With MiConsulta2
                        If Trim(Grid_Facturas_Pagar.TextMatrix(I, 6)) = "FACTURA" Then
                            Mi_SQL = "SELECT No_Factura, Pagada, Total, Abono, Saldo, Fecha_Pago, Forma_Pago, Usuario_Modifico, Fecha_Modifico,No_Parcialidad"
                            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
                            Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Grid_Facturas_Pagar.TextMatrix(I, 0) & "'"
                        End If
                        If Trim(Grid_Facturas_Pagar.TextMatrix(I, 6)) = "REMISION" Then
                            Mi_SQL = "SELECT No_Remision, Pagada, Total, Abono, Saldo, Fecha_Pago, Forma_Pago, Usuario_Modifico, Fecha_Modifico"
                            Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones"
                            Mi_SQL = Mi_SQL & " WHERE No_Remision = '" & Grid_Facturas_Pagar.TextMatrix(I, 0) & "'"
                        End If
                        Set Rs_MiFactura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    End With
                    'Actualiza la factura
                    With Rs_MiFactura
                        Rs_MiFactura.Edit
                            If Format(Grid_Facturas_Pagar.TextMatrix(I, 10), "###,###.00") >= Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ",")) Then
                                .rdoColumns("Pagada") = "S"
                                .rdoColumns("Saldo") = 0
                                .rdoColumns("Abono") = .rdoColumns("Total")
                                .rdoColumns("Fecha_Pago") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                                .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                .rdoColumns("Fecha_Modifico") = Format(Now, "MM/dd/yyyy")
                                Saldo_Factura = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ","))
                                Importe = Format(Format(Importe, "###,###.00") - Format(Saldo_Factura, "###,###.0000"), "###,###.00")
                                Cargo = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ","))
                                If Trim(Grid_Facturas_Pagar.TextMatrix(I, 6)) = "FACTURA" Then .rdoColumns("No_Parcialidad") = Grid_Facturas_Pagar.TextMatrix(I, 9)
                            Else
                                .rdoColumns("Saldo") = Format(Val(.rdoColumns("Saldo")), "###,##0.00") - Format(Grid_Facturas_Pagar.TextMatrix(I, 10), "###,##0.00") 'Format(Importe, "###,##0.00")
                                .rdoColumns("Abono") = .rdoColumns("Abono") + Format(Grid_Facturas_Pagar.TextMatrix(I, 10), "###,##0.00") 'Importe
                                .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                .rdoColumns("Fecha_Modifico") = Format(Now, "MM/dd/yyyy")
                                Cargo = Importe
                                Importe = 0
                                If Trim(Grid_Facturas_Pagar.TextMatrix(I, 6)) = "FACTURA" Then .rdoColumns("No_Parcialidad") = Grid_Facturas_Pagar.TextMatrix(I, 9)
                            End If
                        .Update
                    End With
                    If Grid_Facturas_Pagar.TextMatrix(I, 6) = "FACTURA" Then  'SE AGREGA EL MOVIMIENTO DE LA FACTURA
                        'Agrega a Adm_Movimiento
                        Set Rs_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
                            Rs_Movimientos.AddNew
                                Rs_Movimientos.rdoColumns("No_Movimiento") = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                                Rs_Movimientos.rdoColumns("Referencia") = Txt_Referencia.text
                                Rs_Movimientos.rdoColumns("No_Factura") = Grid_Facturas_Pagar.TextMatrix(I, 0)
                                Rs_Movimientos.rdoColumns("Fecha") = Format(Now, "MM/dd/yyyy")
                                Rs_Movimientos.rdoColumns("Tipo_Pago") = "I"
                                Rs_Movimientos.rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                                Rs_Movimientos.rdoColumns("Banco") = Cmb_Banco.text
                                Rs_Movimientos.rdoColumns("Cliente_ID") = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000")
                                Rs_Movimientos.rdoColumns("Estatus") = "A"
                                Rs_Movimientos.rdoColumns("Comentarios") = Txt_Comentarios.text
                                If Cmb_Forma_Pago.text = "Cheque" Then Rs_Movimientos.rdoColumns("Beneficiario") = Cmb_Cliente.text
                                If Format(Cargo, "###,###.00") >= Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ",")) Then
                                    Rs_Movimientos.rdoColumns("Concepto") = UCase("Pago")
                                Else
                                    Rs_Movimientos.rdoColumns("Concepto") = UCase("Abono")
                                End If
                                Rs_Movimientos.rdoColumns("Cantidad") = Grid_Facturas_Pagar.TextMatrix(I, 10) ' Cargo
'                                Rs_Movimientos.rdoColumns("Tipo_Pago") = Trim(Cmb_Tipo_Pago.text)
'                                Rs_Movimientos.rdoColumns("Dias_Pago") = Val(Txt_Dias.text)
                                Rs_Movimientos.rdoColumns("No_Complemento_Pago") = Format(Conectar_Ayudante.Maximo_Catalogo("Complemento_Pago", "No_Factura"), "#0000000000")
                                Rs_Movimientos.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Movimientos.rdoColumns("Fecha_Creo") = Now
                           Rs_Movimientos.Update
                           Rs_Movimientos.Close
                         Rs_MiFactura.Close
                     Else 'SE AGREGA EL MOVIMIENTO DE LA REMISION
                        'Agrega a Adm_Movimiento
                        Set Rs_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
                            Rs_Movimientos.AddNew
                                Rs_Movimientos.rdoColumns("No_Movimiento") = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                                Rs_Movimientos.rdoColumns("Referencia") = Txt_Referencia.text
                                Rs_Movimientos.rdoColumns("No_Remision") = Grid_Facturas_Pagar.TextMatrix(I, 0)
                                Rs_Movimientos.rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                                Rs_Movimientos.rdoColumns("Tipo_Pago") = "I"
                                Rs_Movimientos.rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                                Rs_Movimientos.rdoColumns("Banco") = Cmb_Banco.text
                                Rs_Movimientos.rdoColumns("Cliente_ID") = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000")
                                Rs_Movimientos.rdoColumns("Estatus") = "A"
                                Rs_Movimientos.rdoColumns("Comentarios") = Txt_Comentarios.text
                                If Cmb_Forma_Pago.text = "Cheque" Then Rs_Movimientos.rdoColumns("Beneficiario") = Cmb_Cliente.text
                                If Format(Cargo, "###,###.00") >= Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ",")) Then
                                    Rs_Movimientos.rdoColumns("Concepto") = UCase("Pago")
                                Else
                                    Rs_Movimientos.rdoColumns("Concepto") = UCase("Abono")
                                End If
                                Rs_Movimientos.rdoColumns("Cantidad") = Cargo
'                                Rs_Movimientos.rdoColumns("Tipo_Pago") = Trim(Cmb_Tipo_Pago.text)
'                                Rs_Movimientos.rdoColumns("Dias_Pago") = Val(Txt_Dias.text)
                                Rs_Movimientos.rdoColumns("Usuario_Creo") = Nombre_Usuario
                                Rs_Movimientos.rdoColumns("Fecha_Creo") = Now
                           Rs_Movimientos.Update
                           Rs_Movimientos.Close
                         Rs_MiFactura.Close
                     End If
                Next I
                If Tipo = "FACTURA" Then
                    Registra_Pago
                End If
                Cmb_Cliente.text = ""
                Grid_Facturas_Pendientes.Rows = 0
                Grid_Facturas_Pagar.Rows = 0
                Txt_Pago.text = ""
                Txt_Referencia.text = ""
                Cmb_Banco.ListIndex = -1
                Txt_Total.text = ""
'                Cmb_Tipo_Pago.ListIndex = -1
                Cmb_Tipo_Cadena.ListIndex = 0
                Cmb_Tipo_Cadena_Click
                Cmb_Forma_Pago.ListIndex = -1
                Cmb_Relacionados.ListIndex = -1
                Cmb_Banco_Des.ListIndex = -1
                Txt_Cuenta_Destino.text = ""
                Txt_Cuenta_Origen.text = ""
                Text2.text = ""
                Text2_LostFocus
                DTP_Fecha.Value = Now
                Cmb_Tipo_Moneda.ListIndex = -1
                Conexion_Base.CommitTrans
                MsgBox "Cobranza Capturada", vbInformation
            Else
                MsgBox "Operación Abortada", vbExclamation
            End If
        Else
            MsgBox "Datos imcompleto para realizar el pago", vbExclamation
        End If
    Exit Sub
handler:
MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
'    For Each Er In rdoErrors
        MsgBox Err.Description
'    Next Er
End Sub
Private Sub Registra_Pago()
 Dim Rs_Alta_Nota_Credito As rdoResultset            'Variable para el manejo de la tabla
Dim Rs_Consulta_Parametro_Folio As rdoResultset     'Variable para el manejo de la tabla
Dim Rs_Alta_Detalles_Nota_Credito As rdoResultset   'Variable para el manejo de la tabla
Dim Rs_Actualiza_Factura As rdoResultset
Dim Rs_Modifica_Factura As rdoResultset
Dim Rs_Consulta_Regimen As rdoResultset
Dim Rs_Consulta_Unidad As rdoResultset
Dim Autorizacion_Electronica As String              'Almacena los valores para factura electronica
Dim Año_Electronico As String                       'Almacena los valores para factura electronica
Dim Fecha_Xml As Date                               'Almacena la fecha del xml
Dim Str_Cadena_Original As String                   'Almacena la cadena original de la factura
Dim Str_Cadena_UTF As String                        'Almacena la cadena en formato utf
Dim Str_Cadena_MD5 As String                        'Almacena la cadena en formato md5
Dim Str_Cadena_Sello As String                      'Almacena la cadena del sello digital
Dim Cont_Detalles_Notas_Credito As Integer          'Almacena el conteo de las partidas
Dim Fecha_Timbrado As Date                          'Almacena la fecha del timbrado
Dim Grupo_Fecha_Timbrado() As String                'Almacena la fecha del timbrado separada por T
Dim Iva As Double
Dim SubTotal As Double
Dim Retenciones As Double
Dim Descuento As Double
Dim Hora_Generacion As Date                         'almacena la hora en que se genera el documento
Dim UM As String
Dim Misql As String
Dim Rs_Consulta_Clientes As rdoResultset
Dim crxReport As CRAXDDRT.report

    MDIFrm_Apl_Principal.MousePointer = 11
    Saldo_Anterior = 0
    Saldo_Actual = 0
    UUID = ""
    Call Conectar_Ayudante.Limpia_Variables
    Conexion_Base.BeginTrans
    CFD_Generales.Tipo_Factura = "Pagos"
    'Asigna los datos
    CFD_Generales.Version = "3.3"
'    CFD_Generales.Serie = "P"
    CFD_Generales.Folio = Conectar_Ayudante.Maximo_Catalogo("Complemento_Pago WHERE Serie = '" & CFD_Generales.Serie & "'", "No_Factura")
    'Asigna la fecha del xml
    Hora_Generacion = DateAdd("n", -5, Now)
    Fecha_Xml = Format(Now, "MM/dd/yyyy") & " " & Format(Hora_Generacion, "HH:mm:ss")
    CFD_Generales.Fecha = Format(Fecha_Xml, "yyyy-MM-dd") & "T" & Format(Fecha_Xml, "HH:mm:ss")
'    CFD_Generales.Forma_Pago = Mid(Forma_Pago, 1, 2)
    CFD_Generales.Condiciones_Pago = ""
    
    CFD_Generales.SubTotal = 0
    CFD_Generales.Descuento = Format(Descuento, "#0.00")
    'If Grid_Notas_Credito_Electronicas.Rows = 0 Then
       ' CFD_Generales.Total = Format(Val(Quitar_Caracter(Txt_Total_Factura.text, ",")), "#0.00")
       ' Else
        'CFD_Generales.Total = Format(Val(Quitar_Caracter(Txt_Total.text, ",")), "#0.00")
       ' End If
    CFD_Generales.Tipo_Moneda = "XXX"
    ReDim CFD_Relacionados_Conceptos(0)
'    CFD_Relacionados.Existe = False
'    CFD_Relacionados.Relacionados = ""
'    CFD_Relacionados.UUID_Relacionados = ""
            
'    CFD_Generales.Descuento = 0
    CFD_Generales.Tipo_Comprobante = "P"
    'CFD_Generales.Metodo_Pago = CFD_Elimina_Espacios("PUE")
    'CFD_Generales.No_Cuenta_Pago = Trim(Cuenta_Pago)
    
    'Asigna los datos del EMISOR al CFD para generar el xml
    CFD_Emisor.Nombre = CFD_Elimina_Espacios(Nombre_Emisor)
    CFD_Emisor.RFC = CFD_Elimina_Espacios(RFC_Emisor)
    CFD_Emisor.Expedido_En = CFD_Elimina_Espacios(Codigo_Postal_Emisor)
    CFD_Emisor.Calle = CFD_Elimina_Espacios(Calle_Emisor)
    CFD_Emisor.No_Exterior = CFD_Elimina_Espacios(No_Exterior_Emisor)
    CFD_Emisor.No_Interior = CFD_Elimina_Espacios(No_Interior_Emisor)
    CFD_Emisor.Colonia = CFD_Elimina_Espacios(Colonia_Emisor)
    CFD_Emisor.cp = CFD_Elimina_Espacios(Codigo_Postal_Emisor)
    CFD_Emisor.Localidad = CFD_Elimina_Espacios(Municipio_Emisor)
    CFD_Emisor.Municipio = CFD_Elimina_Espacios(Municipio_Emisor)
    CFD_Emisor.Estado = CFD_Elimina_Espacios(Estado_Emisor)
    CFD_Emisor.Pais = "México"
    CFD_Emisor.Referencia = ""
    CFD_Emisor.Regimen_Fiscal = Regimen_Emisor
    
'      Mi_SQL = "SELECT Regimen_Fiscal FROM Cat_Parametros_Factura_Electronica"
'    Set Rs_Consulta_Regimen = Recordset_Consultar(Mi_SQL)
'    CFD_Emisor.Regimen_Fiscal = Mid(Rs_Consulta_Regimen.rdoColumns("Regimen_Fiscal"), 1, 3)
'    Rs_Consulta_Regimen.Close
'
'    'Asigna los datos del EXPEDIDOEN al CFD para generar el xml
'    CFD_ExpedidoEn.Calle = CFD_Elimina_Espacios(Calle_Sucursal)
'    CFD_ExpedidoEn.No_Exterior = CFD_Elimina_Espacios(No_Exterior_Sucursal)
'    CFD_ExpedidoEn.No_Interior = CFD_Elimina_Espacios(No_Interior_Sucursal)
'    CFD_ExpedidoEn.colonia = CFD_Elimina_Espacios(Colonia_Sucursal)
'    CFD_ExpedidoEn.Ciudad = CFD_Elimina_Espacios(Ciudad_Sucursal)
'    CFD_ExpedidoEn.estado = CFD_Elimina_Espacios(Estado_Sucursal)
'    CFD_ExpedidoEn.Pais = CFD_Elimina_Espacios(Pais_Sucursal)
'    CFD_ExpedidoEn.cp = CFD_Elimina_Espacios(CP_Sucursal)
        
    'Asigna los datos del RECEPTOR al CFD para generar el xml
    Misql = "Select * FROM Cat_Clientes WHERE Cliente_ID='" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
    Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Misql)
    With Rs_Consulta_Clientes
         CFD_Receptor.Nombre = CFD_Elimina_Espacios(.rdoColumns("Nombre"))
        CFD_Receptor.RFC = CFD_Elimina_Espacios(.rdoColumns("RFC"))
        CFD_Receptor.Calle = CFD_Elimina_Espacios(.rdoColumns("Direccion"))
'        CFD_Receptor.No_Exterior = .rdoColumns("Numero_exterior")
'        CFD_Receptor.No_Interior = .rdoColumns("Numero_interior")
        CFD_Receptor.Colonia = CFD_Elimina_Espacios(.rdoColumns("Colonia"))
        CFD_Receptor.cp = CFD_Elimina_Espacios(.rdoColumns("CP"))
    End With
    Rs_Consulta_Clientes.Close
'    CFD_Receptor.Calle = CFD_Elimina_Espacios(Txt_Calle.text)
'    CFD_Receptor.No_Exterior = CFD_Elimina_Espacios(Txt_No_Exterior.text)
'    CFD_Receptor.No_Interior = CFD_Elimina_Espacios(Txt_No_Interior.text)
'    CFD_Receptor.colonia = CFD_Elimina_Espacios(Txt_Colonia.text)
'    CFD_Receptor.cp = CFD_Elimina_Espacios(Txt_Codigo_Postal.text)
'    CFD_Receptor.Localidad = CFD_Elimina_Espacios(Txt_Municipio.text)
'    CFD_Receptor.Municipio = CFD_Elimina_Espacios(Txt_Municipio.text)
'    CFD_Receptor.estado = CFD_Elimina_Espacios(Txt_Estado.text)
'    CFD_Receptor.Pais = CFD_Elimina_Espacios(Txt_Pais.text)
'    CFD_Receptor.Referencia = ""
    CFD_Receptor.Uso_CFDI = "P01 Por Definir"
  
        'Asigna el conteo de partidas al arreglo
        ReDim CFD_Conceptos(1)
        CFD_Conceptos(1).Cod_prod = "84111506"
        CFD_Conceptos(1).No_Identificacion = ""
        CFD_Conceptos(1).Cantidad = "1"
        CFD_Conceptos(1).Unidad_Medida = "ACT"
        CFD_Conceptos(1).Descripcion = "Pago"
        CFD_Conceptos(1).Valor_Unitario = 0
        CFD_Conceptos(1).Importe = 0
        'CFD_Conceptos(1).IVA_Producto = False
   
    'CFD_Impuestos_Retenidos = 0
    IVA_EXENTO = True
    ReDim CFD_Impuestos_Retenidos(0)
    ReDim CFD_Impuestos(0)
    'Datos del pago
    CFD_Pagos.Fecha_Pago = Format(DTP_Fecha, "yyyy-MM-dd") & "T" & Format(0, "HH:mm:ss")
    CFD_Pagos.Forma_Pago_Pagos = Cmb_Forma_Pago.text 'Cmb_Tipo_Pago.text
    CFD_Pagos.Moneda_Pago = Cmb_Tipo_Moneda.text
    If Cmb_Tipo_Moneda.text <> "MXN" Then
        CFD_Pagos.Tipo_Cambio_Pago = Text2.text
    Else
        CFD_Pagos.Tipo_Cambio_Pago = ""
    End If
    CFD_Pagos.Monto = Txt_Pago.text 'Text1.text
    CFD_Pagos.Num_Operacion = Trim(Txt_Referencia.text) 'Txt_No_Cheque.text
'    CFD_Pagos.RfcEmisorCtaOrd = CFD_Receptor.RFC
    'CFD_Pagos.NomBancoOrdExt = Txt_Nombre_Banco.text
    CFD_Pagos.CtaOrdenante = Trim(Txt_Cuenta_Origen.text)
'    CFD_Pagos.RfcEmisorCtaBen = CFD_Emisor.RFC
    CFD_Pagos.CtaBeneficiario = Trim(Txt_Cuenta_Destino.text)
'    CFD_Generales.Condiciones_Pago = Txt_Comentarios.text
    ReDim CFD_Pagos_DR(Grid_Facturas_Pagar.Rows - 1)
    For I = 1 To Grid_Facturas_Pagar.Rows - 1
        CFD_Pagos_DR(I).ID_Doc = Grid_Facturas_Pagar.TextMatrix(I, 8)
        'CFD_Pagos_DR(i).Serie = FG_Facturas_Agregadas.TextMatrix(i, 0)
        CFD_Pagos_DR(I).Folio = Format(Grid_Facturas_Pagar.TextMatrix(I, 1), "0000000000")
        'If Grid_Facturas_Pagar.TextMatrix(I, 5) = "Pesos" Then
        CFD_Pagos_DR(I).Moneda_DR = "MXN"
          '  CFD_Pagos_DR(I).Tipo_Cambio_DR = ""
        'Else
        '    CFD_Pagos_DR(I).Moneda_DR = "USD"
         '   CFD_Pagos_DR(I).Tipo_Cambio_DR = FG_Facturas_Agregadas.TextMatrix(I, 7)
        'End If
        'CFD_Pagos_DR(i).Tipo_Cambio_DR = FG_Facturas_Agregadas.TextMatrix(i, 7)
        CFD_Pagos_DR(I).Metodo_Pago_DR = "PPD Pago en parcialidades o diferido"
        CFD_Pagos_DR(I).No_Parcialidad = Grid_Facturas_Pagar.TextMatrix(I, 9)
        CFD_Pagos_DR(I).Saldo_Anterior = Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(I, 5), ",")
        CFD_Pagos_DR(I).Saldo_Insoluto = Grid_Facturas_Pagar.TextMatrix(I, 11) 'Grid_Facturas_Pagar.TextMatrix(I, 5) - FG_Facturas_Agregadas.TextMatrix(I, 1)
        CFD_Pagos_DR(I).Importe_Pagado = Grid_Facturas_Pagar.TextMatrix(I, 10)
        Saldo_Anterior = Saldo_Anterior + Val(CFD_Pagos_DR(I).Saldo_Anterior)
        Saldo_Actual = Saldo_Actual + Val(CFD_Pagos_DR(I).Saldo_Insoluto)
     Next
    
    'Crea el sello digital con toda la informacion
    CFD_Generales.No_Certificado = CFD_Consulta_Serie_Certificado(Ruta_Certificado)
    CFD_Generales.Certificado = CFD_Consulta_Certificado(Ruta_Certificado)
    Str_Cadena_Original = CFD_Cadena_Original("PAGOS")
    Str_Cadena_UTF = CFD_Valida_Caracteres_UTF(Str_Cadena_Original)
    Str_Cadena_MD5 = CFD_Genera_MD5(Str_Cadena_UTF)
    Str_Cadena_Sello = CFD_Genera_Sello(Str_Cadena_UTF, Ruta_Llave_Privada) ', Year(Fecha_Xml))
    CFD_Generales.Cadena_Original = Str_Cadena_UTF
    CFD_Generales.Sello = Str_Cadena_Sello
    If CFD_Generales.Tipo_Moneda = "MXN" Or CFD_Generales.Tipo_Moneda = "XXX" Then
        CFD_Generales.Importe_Letra = Conectar_Ayudante.Convierte_Cantidad_Letras(Format(CStr(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago, ","))), "#0.00")) ', 1)
    Else
        CFD_Generales.Importe_Letra = Conectar_Ayudante.Convierte_Cantidad_Letras(Format(CStr(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago, ","))), "#0.00")) ', 2)
    End If
    'Registra del pagos
    For I = 1 To Grid_Facturas_Pagar.Rows - 1
        Set Rs_Alta_Nota_Credito = Conectar_Ayudante.Recordset_Agregar("Complemento_Pago")
            With Rs_Alta_Nota_Credito
                .AddNew
                    .rdoColumns("Clave") = Format(Conectar_Ayudante.Maximo_Catalogo("Complemento_Pago", "Clave"), "#0000000000")
                    .rdoColumns("No_Factura") = Format(CFD_Generales.Folio, "#0000000000")
                    .rdoColumns("Serie") = CFD_Generales.Serie
                    .rdoColumns("RFC") = CFD_Receptor.RFC
                    .rdoColumns("Cliente_ID") = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") 'Format(Cmb_Clientes.ItemData(Cmb_Clientes.ListIndex), "00000")
                    .rdoColumns("Cliente") = Trim(Cmb_Cliente.text)
                    .rdoColumns("Direccion") = CFD_Receptor.Calle
                    .rdoColumns("No_Exterior") = CFD_Receptor.No_Exterior
                    .rdoColumns("No_Interior") = CFD_Receptor.No_Interior
                    .rdoColumns("Colonia") = CFD_Receptor.Colonia
                    .rdoColumns("Municipio") = CFD_Receptor.Localidad
                   ' .rdoColumns("Estado") = CFD_Receptor.estado
    '                .rdoColumns("Pais") = Txt_Pais.text
                    .rdoColumns("Codigo_Postal") = CFD_Receptor.cp
                    .rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy HH:mm:ss")
                    '.rdoColumns("Tipo_Nota_Credito") = Cmb_Tipo_Nota_Credito.text
                    '.rdoColumns("Orden_Compra") = Trim(Txt_Orden_Compra.text)
                    If CFD_Pagos_DR(I).Serie <> "" Then .rdoColumns("Serie_Factura") = CFD_Pagos_DR(I).Serie
                    .rdoColumns("No_Factura_Ref") = Format(Val(CFD_Pagos_DR(I).Folio), "0000000000")
                    .rdoColumns("Tipo_Moneda") = Trim(CFD_Pagos.Moneda_Pago)
    '                .rdoColumns("Tipo_Cambio") = Val(CFD_Pagos.Tipo_Cambio_Pago)
                     .rdoColumns("Estatus") = "ACTIVA"
                    .rdoColumns("Fecha_Creo_Xml") = Format(Fecha_Xml, "yyyy/MM/dd HH:mm:ss")
                    .rdoColumns("No_Certificado") = CFD_Generales.No_Certificado
                    '.rdoColumns("No_Autorizacion") = CFD_Generales.No_Aprobacion
                    '.rdoColumns("Año_Autorizacion") = CFD_Generales.Año_Aprobacion
    '                .rdoColumns("Certificado") = CFD_Generales.Certificado
'                    .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios.text))
                    .rdoColumns("Ruta_Codigo_BD") = Ruta_Pdfs & "\CFDI_" & Trim(CFD_Generales.Serie & "_" & CFD_Generales.Folio) & ".bmp"
                    .rdoColumns("Metodo_Pago") = "PPD Pago en parcialidades o diferido"
                    '.rdoColumns("No_Cuenta_Pago") = Cuenta_Pago
                    .rdoColumns("Usuario_Creo") = NombreVendedor
                    .rdoColumns("Fecha_Creo") = Now
                    .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                    .rdoColumns("Monto_Pagado") = Format(CFD_Pagos_DR(I).Importe_Pagado, "#0.00")
                    .rdoColumns("Num_Operacion") = Txt_No_Cheque
                    .rdoColumns("RFC_Emisor_Cta_Ord") = CFD_Pagos.RfcEmisorCtaOrd
                    '.rdoColumns("Nom_Banco_Ord_Ext") = Txt_Nombre_Banco.text
                    .rdoColumns("Cta_Ordenante") = Txt_Cuenta_Origen.text
                    .rdoColumns("RFC_Emisor_Cta_Ben") = CFD_Pagos.RfcEmisorCtaBen
                    .rdoColumns("Cta_Beneficiario") = Txt_Cuenta_Destino.text
                     If Cmb_Tipo_Cadena.ListIndex > 0 Then .rdoColumns("Tipo_Cad_Pago") = Cmb_Tipo_Cadena.text
                    .rdoColumns("Cert_Pago") = Txt_Certificado.text
                    .rdoColumns("Cadena_Pago") = Txt_Cadena.text
                    .rdoColumns("Sello_Pago") = Txt_Sello.text
                    .rdoColumns("UUID_Relacionado") = Trim(CFD_Pagos_DR(I).ID_Doc)
                    If CFD_Relacionados.Existe Then
                        .rdoColumns("Tipo_Relacion") = Trim("04 Sustitucion de los CFDI previos")
                        .rdoColumns("UUID_Relacion") = CFD_Relacionados.UUID_Relacionados
                    End If
                    .rdoColumns("No_Parcialidad") = CFD_Pagos_DR(I).No_Parcialidad
                    .rdoColumns("Saldo_Actual") = CFD_Pagos_DR(I).Saldo_Insoluto
                    .rdoColumns("Saldo_Anterior") = CFD_Pagos_DR(I).Saldo_Anterior
                    .rdoColumns("Banco_ord") = Cmb_Banco.text
                    .rdoColumns("Banco_Ben") = Cmb_Banco_Des.text
                    .rdoColumns("No_Factura_Electronica") = Grid_Facturas_Pagar.TextMatrix(I, 1)
                    .rdoColumns("Serie_Factura_Electronica_Relacionada") = Grid_Facturas_Pagar.TextMatrix(I, 12)
                    
                .Update
            End With
        Rs_Alta_Nota_Credito.Close
    Next
    
    'Registra las partidas de la nota de credito
       
    'Crea el xml con los datos del recibo de pagos
    Call CFD_Crea_Xml("CFDI_" & Trim(CFD_Generales.Serie) & "_" & Val(CFD_Generales.Folio), "PAGOS", "") ', "", "")
    
    'Convierte la fecha de timbrado
    Grupo_Fecha_Timbrado = Split(Timbrado_FechaTimbrado, "T")
    Fecha_Timbrado = Grupo_Fecha_Timbrado(0) & " " & Grupo_Fecha_Timbrado(1)
    
    'Actualiza la factura con el timbrado
    'Mi_SQL = "SELECT * FROM Complemento_Pagos WHERE Serie = '" & Trim(Cmb_Serie.text) & "' AND No_Nota_Credito = '" & Format(Val(Txt_No_Nota_Credito.text), "0000000000") & "'"
     Mi_SQL = "SELECT * FROM Complemento_Pago WHERE No_Factura='" & Format(Val(CFD_Generales.Folio), "0000000000") & "' and Serie='" & CFD_Generales.Serie & "'"
    Set Rs_Modifica_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Modifica_Factura
            While Not .EOF
                .Edit
                    .rdoColumns("Timbre_Version") = Timbrado_VersionSat
                    .rdoColumns("Timbre_UUID") = Timbrado_UUID
                    .rdoColumns("Timbre_FechaTimbrado") = Format(Fecha_Timbrado, "MM/dd/yyyy HH:mm:ss")
                    .rdoColumns("Timbre_selloCFD") = Timbrado_selloCFD
                    .rdoColumns("Timbre_NoCertificadoSA") = Timbrado_noCertificadoSAT
                    .rdoColumns("Timbre_selloSAT") = Timbrado_selloSAT
                .Update
                .MoveNext
            Wend
        End With
    Rs_Modifica_Factura.Close
    Conexion_Base.CommitTrans
'    'Regresa los controles inicial
'    Fra_Datos_Factura_Generales.Enabled = False
'    Fra_Datos_Clientes.Enabled = False
'    'Fra_Consultas_Listado.Enabled = False
'    'Fra_Descuentos.Enabled = False
'    Fra_Comentarios.Enabled = False
'    Fra_Totales.Enabled = False
'    Btn_Nueva.Caption = "Nueva"
'    Btn_Salir.Caption = "Salir"
'    Btn_Ver_Pdf.Enabled = True
'    Btn_Ver_Xml.Enabled = True
'    Btn_Cancelar.Enabled = True
'    Btn_Busqueda.Enabled = True
    MDIFrm_Apl_Principal.MousePointer = 0
    'MsgBox "Recibo de pago generado exitosamente", vbInformation
''    Valida_Termino_Serie
    'Call Aviso_Termino_Folios("NOTA CREDITO", Trim(Cmb_Serie.text))
    'Crea el pdf con los datos de la nota de credito
    Call CFD_Crea_PDF("CFDI_" & Trim(CFD_Generales.Serie) & "_" & Val(CFD_Generales.Folio), "PAGOS", "PAGOS", Year(Fecha_Xml))
 
    
    If MsgBox("¿Desea visualizar el PDF?", vbQuestion + vbYesNo) = vbYes Then
        'Realiza la impresión de la factura
        'CFD_Documento.save App.Path & "\CFDs\" & "CFDI_" & Trim(CFD_Generales.Serie) & "_" & Val(CFD_Generales.Folio) & ".pdf"
        ShellExecute ByVal 0&, "open", App.Path & "\PDF\" & "CFDI_" & Trim(CFD_Generales.Serie) & "_" & Val(CFD_Generales.Folio) & ".pdf", vbNullString, vbNullString, SW_SHOWMAXIMIZED
        
    End If
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    'Obtiene el error
'    If Err.Number = 7777 Then
'    MsgBox Err.Description
    
    Err.Raise &HFFFFFF01, "Error", Err.Description
'    Else
'        For Each Rdo_Error In rdoErrors
'            MsgBox Rdo_Error.Description
'        Next
'    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Btn_Eliminar_Click()
'DESCRIPCIÓN: Elimina los registros de la base de datos asi como del grid
'PARÁMETROS:
'CREO:  Sergi Godínez Banda
'FECHA_CREO:    8-Agosto-2007
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Eliminar_Click()
Dim Cont_Facturas As Integer                              'Usada para contar facturas del Grid
Dim Suma As Double                                       'Usada para realizar la suma del saldo

    If Grid_Facturas_Pagar.RowSel > 0 Then
        If Grid_Facturas_Pendientes.Rows = 0 Then
        Grid_Facturas_Pendientes.AddItem "Documento" & Chr(9) & "Electrónica" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & _
            Chr(9) & "Saldo" & Chr(9) & "Tipo" & Chr(9) & "Fecha Pago"
        End If
        Grid_Facturas_Pendientes.AddItem Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 0) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 1) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 2) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 3) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 4) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 5) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 6) & Chr(9) _
            & Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 7)
        
         'Calcula el saldo total
            Suma = 0
            Suma = Suma + Val(Conectar_Ayudante.Quitar_Caracter(Grid_Facturas_Pagar.TextMatrix(Grid_Facturas_Pagar.RowSel, 10), ","))
        'Asigna la suma a Total y Pago
            Txt_Total.text = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ",")) - Suma
            Txt_Total.text = Format(Txt_Total.text, "###,##0.00")
            Txt_Pago = Txt_Total.text
        'remueve del otro grid la factura
        If Grid_Facturas_Pagar.Rows = 2 Then
            Grid_Facturas_Pagar.Rows = 0
        Else
            Grid_Facturas_Pagar.RemoveItem (Grid_Facturas_Pagar.RowSel)
        End If
        Grid_Facturas_Pendientes.FixedRows = 1
    End If
End Sub

Private Sub Btn_Salir_Click()
    Unload Frm_Adm_Cobranza
End Sub

Public Function Consulta_Folio()
Dim Rs_Consulta_Serie As rdoResultset
     'Consulta la serie de factura usada actualmente
        Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Tipo = 'PAGOS' AND Estatus = 'ACTIVO'"
        Set Rs_Consulta_Serie = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Serie.EOF Then
                   Lbl_Folio.Caption = Rs_Consulta_Serie.rdoColumns("Serie")
                   CFD_Generales.Serie = Rs_Consulta_Serie.rdoColumns("Serie")
            Else
                'No hay serie activa y muestra mensaje
                MsgBox "No se encontró una serie activa para utilizar, favor de verificar", vbExclamation
                Cmb_Cliente.text = ""
                Exit Function
            End If
        Rs_Consulta_Serie.Close
        Lbl_Folio.Caption = Lbl_Folio.Caption & Conectar_Ayudante.Maximo_Catalogo("Complemento_Pago WHERE Serie = '" & CFD_Generales.Serie & "'", "No_Factura")
        
        'Valida si aun existen folios de facturas disponibles para utilizar
         Call Aviso_Termino_Folios("PAGOS")
         'si la bandera esta habilitada muestra mensaje y cancela la operación
         If Folios_Terminados = True Then
             Lbl_Folio.Caption = ""
             Cmb_Cliente.text = ""
             Grid_Facturas_Pendientes.Rows = 0
             MsgBox "No se encontraron folios de factura disponibles, favor de verificar", vbCritical
             Consulta_Folio = False
             Exit Function
         Else
             Consulta_Folio = True
         End If
                    
End Function

Private Sub Cmb_Banco_Click()
    CFD_Pagos.RfcEmisorCtaOrd = Busca_Banco(Cmb_Banco.text)
End Sub
Public Function Busca_Banco(Nombre As String)
    Dim Rs_Consulta As rdoResultset
    Mi_SQL = "SELECT RFC FROM Cat_Bancos WHERE Nombre='" & Nombre & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
        If Not IsNull(Rs_Consulta.rdoColumns("RFC")) Then
            Busca_Banco = Rs_Consulta.rdoColumns("RFC")
        Else
            Busca_Banco = ""
        End If
    End If
    Rs_Consulta.Close
End Function
Private Sub Cmb_Banco_Des_Click()
    CFD_Pagos.RfcEmisorCtaBen = Busca_Banco(Cmb_Banco_Des.text)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Cmb_Cliente_Click()
'DESCRIPCIÓN            : Permite elegir el nombre del cliente y asi mostrar en pantalla
'                         las facturas que tiene
'PARÁMETROS             :
'CREO                   :  Julio Cruz
'FECHA_CREO             :  18-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Cmb_Cliente_Click()
Dim Mi_SQL As String
Dim Rs_MiTabla As rdoResultset
Dim Importe As Double
Dim Tipo_Documento As String
Dim Electronico As String
Dim Parcialidad As Long
    
        Set Conectar_Ayudante = New Ayudante
        Txt_Buscar_Factura.text = ""
        Grid_Facturas_Pendientes.Rows = 0
        Cmb_Relacionados.ListIndex = -1
        If (Grid_Facturas_Pagar.Rows = 0 And Btn_Capturar.Caption = "Capturar") Or Id_cliente <> Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") Then Grid_Facturas_Pagar.Rows = 0
        Grid_Facturas_Pendientes.AddItem "Documento" & Chr(9) & "Electrónica" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & _
            Chr(9) & "Saldo" & Chr(9) & "Tipo" & Chr(9) & "Fecha Pago" & Chr(9) & "UUID" & Chr(9) & "Parcialidad"
        
        'Consulta En la tabla de facturas
        Mi_SQL = "SELECT No_Factura, No_Factura_Electronica, Fecha, Total, Abono, Saldo, Tipo_Pago, Tipo_Documento, Fecha_Pago,Timbre_UUID,No_Parcialidad,forma_pago, serie"
        Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas "
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
        Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
        Mi_SQL = Mi_SQL & " ORDER BY No_Factura "
        Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_MiTabla
                 If (Grid_Facturas_Pagar.Rows = 0 And Btn_Capturar.Caption = "Capturar") Or Id_cliente <> Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") Then Txt_Total.text = 0
                While Not .EOF
                    If Not IsNull(.rdoColumns("Tipo_Documento")) Then
                        Tipo_Documento = .rdoColumns("Tipo_Documento")
                    Else
                        Tipo_Documento = "FACTURA"
                    End If
                    If Not IsNull(.rdoColumns("No_Factura_Electronica")) Then
                        Electronico = Val(.rdoColumns("No_Factura_Electronica"))
                    Else
                        Electronico = ""
                    End If
                    If Not IsNull(.rdoColumns("No_Parcialidad")) Then
                        Parcialidad = .rdoColumns("No_Parcialidad") + 1
                    Else
                        Parcialidad = 1
                    End If
                    'Cargar el método de pago al combo
    '                If Not IsNull(.rdoColumns("Tipo_Pago")) And Grid_Facturas_Pendientes.Rows = 1 Then
    '                    If .rdoColumns("Tipo_Pago") = "CONTADO" Or .rdoColumns("Tipo_Pago") = "Pago en 1 exhibiciones" Or .rdoColumns("Tipo_Pago") = "Pago en 1 parcialidades" Or .rdoColumns("Tipo_Pago") = "Pago en una sola exhibicion" Or .rdoColumns("Tipo_Pago") = "PUE Pago en una sola exhibicion" Then
    '                        Cmb_Tipo_Pago.ListIndex = 0
    '                    Else
    '                         Cmb_Tipo_Pago.ListIndex = 1
    '                    End If
    '                End If
                    'Cargar forma de pago al combo
                    If Not IsNull(.rdoColumns("forma_pago")) And Grid_Facturas_Pendientes.Rows = 1 Then
                        If .rdoColumns("Forma_Pago") = "01    Efectivo" Or .rdoColumns("Forma_Pago") = "01   Efectivo" Or .rdoColumns("Forma_Pago") = "Efectivo" Then
                            Cmb_Forma_Pago.ListIndex = 0
                        ElseIf .rdoColumns("Forma_Pago") = "03 Transferencia" Or .rdoColumns("Forma_Pago") = "Transferencia" Then
                             Cmb_Forma_Pago.ListIndex = 2
                        Else
                            Cmb_Forma_Pago.ListIndex = 1
                        End If
                    End If
                    'Agrega la partida al grid
                    Grid_Facturas_Pendientes.AddItem .rdoColumns("No_Factura") & Chr(9) & Electronico & Chr(9) _
                        & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "###,##0.00") & Chr(9) _
                        & Format(.rdoColumns("Abono"), "###,##0.00") & Chr(9) & Format(.rdoColumns("Saldo"), "###,##0.00") & Chr(9) _
                        & Tipo_Documento & Chr(9) & .rdoColumns("Fecha_Pago") & Chr(9) & .rdoColumns("Timbre_UUID") & Chr(9) & Parcialidad & Chr(9) & .rdoColumns("Serie")
                    .MoveNext
                    Grid_Facturas_Pendientes.FixedRows = 1
                Wend
            End With
        Rs_MiTabla.Close
        'Consulta En la tabla de Remisiones
        Mi_SQL = "SELECT No_Remision, Fecha, Total, Abono, Saldo, Tipo_Pago,Tipo_Documento,Fecha_Pago"
        Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones "
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
        Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
        Mi_SQL = Mi_SQL & " ORDER BY No_Factura "
        Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_MiTabla
                ''Grid_Facturas_Pendientes.Rows = 0
                ''Grid_Facturas_Pagar.Rows = 0
                 If (Grid_Facturas_Pagar.Rows = 0 And Btn_Capturar.Caption = "Capturar") Or Id_cliente <> Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") Then Txt_Total.text = 0
                ''Grid_Facturas_Pendientes.AddItem "Documento" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Tipo"
                While Not .EOF
                    Grid_Facturas_Pendientes.AddItem .rdoColumns("No_Remision") & Chr(9) & "" & Chr(9) _
                        & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "###,##0.00") & Chr(9) _
                        & Format(.rdoColumns("Abono"), "###,##0.00") & Chr(9) & Format(.rdoColumns("Saldo"), "###,##0.00") & Chr(9) _
                        & .rdoColumns("Tipo_Documento") & Chr(9) & .rdoColumns("Fecha_Pago") & Chr(9) & ""
                    .MoveNext
                    Grid_Facturas_Pendientes.FixedRows = 1
                Wend
            End With
        Rs_MiTabla.Close
        Formatea_Columnas_Grid
        If Consulta_Folio = False And Cmb_Cliente.text <> "" Then
            Cmb_Cliente.text = ""
            Exit Sub
        End If
        Id_cliente = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000")
End Sub

Public Sub Formatea_Columnas_Grid()

    If Grid_Facturas_Pendientes.Rows > 1 Then
        Grid_Facturas_Pendientes.ColWidth(0) = 1000     'Documento
        Grid_Facturas_Pendientes.ColAlignment(0) = 3
        Grid_Facturas_Pendientes.ColWidth(1) = 1000     'Electronico
        Grid_Facturas_Pendientes.ColAlignment(1) = 3
        Grid_Facturas_Pendientes.ColWidth(2) = 1000     'Fecha
        Grid_Facturas_Pendientes.ColAlignment(2) = 3
        Grid_Facturas_Pendientes.ColWidth(3) = 1150     'Total
        Grid_Facturas_Pendientes.ColWidth(4) = 1150     'Abono
        Grid_Facturas_Pendientes.ColWidth(5) = 900      'Saldo
        Grid_Facturas_Pendientes.ColWidth(6) = 900      'Tipo
        Grid_Facturas_Pendientes.ColWidth(7) = 950      'Fecha_Pago
        Grid_Facturas_Pendientes.ColWidth(8) = 0      'UUID
         Grid_Facturas_Pendientes.ColWidth(9) = 0      'Parcialidad
         Grid_Facturas_Pendientes.ColWidth(10) = 0      'Serie
        Grid_Facturas_Pendientes.FixedRows = 1
    Else
        Grid_Facturas_Pendientes.Rows = 0
    End If
    
    If Grid_Facturas_Pagar.Rows > 1 Then
        Grid_Facturas_Pagar.ColWidth(0) = 1000
        Grid_Facturas_Pagar.ColAlignment(0) = 3
        Grid_Facturas_Pagar.ColWidth(1) = 1000     'Electronico
        Grid_Facturas_Pagar.ColAlignment(1) = 3
        Grid_Facturas_Pagar.ColWidth(2) = 1000     'Fecha
        Grid_Facturas_Pagar.ColAlignment(2) = 3
        Grid_Facturas_Pagar.ColWidth(3) = 1150     'Total
        Grid_Facturas_Pagar.ColWidth(4) = 0 '1150     'Abono
        Grid_Facturas_Pagar.ColWidth(5) = 0 '900      'Saldo
        Grid_Facturas_Pagar.ColWidth(6) = 900      'Tipo
        Grid_Facturas_Pagar.ColWidth(7) = 950      'Fecha_Pago
        Grid_Facturas_Pagar.ColWidth(8) = 0      'UUID
        Grid_Facturas_Pagar.ColWidth(9) = 0      'Parcialidad
        Grid_Facturas_Pagar.ColWidth(10) = 900   'Pago
        Grid_Facturas_Pagar.ColWidth(11) = 900   'Saldo
        Grid_Facturas_Pagar.ColWidth(12) = 0   'Serie
        
        Grid_Facturas_Pagar.FixedRows = 1
    Else
        Grid_Facturas_Pagar.Rows = 0
    End If
End Sub

Private Sub Cmb_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Consulta_Clientes
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Clientes
    'DESCRIPCIÓN: Consulta los datos de la tabla Cat_Clientes
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 8-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Clientes()
Dim Mi_SQL As String
Dim Rs_MiTabla As rdoResultset
Dim No_Nombre As Integer
    
    Set Conectar_Ayudante = New Ayudante
    
    'Consulta en la tabla de facturas
    Mi_SQL = "SELECT Distinct Cat_Clientes.Cliente_ID, Cat_Clientes.Nombre, Adm_Clientes_Facturas.Cliente_ID, Adm_Clientes_Facturas.Pagada, Adm_Clientes_Facturas.Cancelada "
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes, Adm_Clientes_Facturas "
    Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID = Adm_Clientes_Facturas.Cliente_ID and  Pagada = 'N'"
    Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre "
    Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta del catalogo de Clientes
    Cmb_Cliente.Clear
    If Not Rs_MiTabla.EOF Then
        While Not Rs_MiTabla.EOF
            Cmb_Cliente.AddItem Rs_MiTabla!Nombre
            Cmb_Cliente.ItemData(Cmb_Cliente.NewIndex) = Rs_MiTabla!Cliente_ID
            Rs_MiTabla.MoveNext
        Wend
        Rs_MiTabla.Close
    End If
    'Consulta en la tabla de remisiones
    Mi_SQL = "SELECT Distinct Cat_Clientes.Cliente_ID, Cat_Clientes.Nombre, Adm_Clientes_Remisiones.Cliente_ID, Adm_Clientes_Remisiones.Pagada, Adm_Clientes_Remisiones.Cancelada "
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes, Adm_Clientes_Remisiones "
    Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID = Adm_Clientes_Remisiones.Cliente_ID and  Pagada = 'N'"
    Mi_SQL = Mi_SQL & " AND Cancelada = 'N'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre "
    Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta del catalogo de Clientes
    ''Cmb_Cliente.Clear
    If Not Rs_MiTabla.EOF Then
        While Not Rs_MiTabla.EOF
            Cmb_Cliente.AddItem Rs_MiTabla!Nombre
            Cmb_Cliente.ItemData(Cmb_Cliente.NewIndex) = Rs_MiTabla!Cliente_ID
            Rs_MiTabla.MoveNext
        Wend
        Rs_MiTabla.Close
    End If
End Sub



'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cmb_Forma_Pago_Click()
    'DESCRIPCIÓN: Muestra o esconde la etiqueta y el textbox de Banco, según la forma
    '             de pago que se elija
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 8-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Cmb_Forma_Pago_Click()
    If Cmb_Forma_Pago.text = "Efectivo" Then
        Txt_Referencia.Visible = False
        Txt_Referencia.text = ""
        Cmb_Banco.Visible = False
        Cmb_Banco.text = ""
        Lbl_Referencia.Visible = False
        Lbl_Banco.Visible = False
    Else
        If Cmb_Forma_Pago.text = "Cheque" Then
            Lbl_Referencia.Caption = "No. Cheque"
            Lbl_Banco.Caption = "Banco del Cheque"
        Else
            Lbl_Referencia.Caption = "Referencia"
            Lbl_Banco.Caption = "Banco Origen"
        End If
        Txt_Referencia.Visible = True
        Cmb_Banco.Visible = True
        Lbl_Referencia.Visible = True
        Lbl_Banco.Visible = True
    End If
End Sub
Public Sub Busca_UUID()
    Dim Rs_Consulta As rdoResultset
    Mi_SQL = "SELECT Timbre_UUID FROM Complemento_Pago WHERE Serie='" & Cmb_Serie.text & "' and No_Factura='" & Cmb_FacRef.text & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
         CFD_Relacionados.UUID_Relacionados = Rs_Consulta.rdoColumns("Timbre_UUID")
    End If
    Rs_Consulta.Close
End Sub
Private Sub Cmb_Relacionados_Click()
Dim Rs_Consulta As rdoResultset
    If Cmb_Relacionados.ListIndex > 0 Then
        Lbl_UUID_Relacion.Enabled = True
        Cmb_Serie.Enabled = True
        Cmb_FacRef.Enabled = True
        Cmb_Serie.Clear
        Mi_SQL = "SELECT DISTINCT Serie FROM Complemento_Pago WHERE Serie is not null"
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta.EOF
         Cmb_Serie.AddItem (Rs_Consulta.rdoColumns("Serie"))
         Rs_Consulta.MoveNext
        Wend
        Rs_Consulta.Close
        CFD_Relacionados.Existe = True
        CFD_Relacionados.Relacionados = Cmb_Relacionados.text
    Else
        Lbl_UUID_Relacion.Enabled = False
        Cmb_Serie.Enabled = False
        Cmb_FacRef.Enabled = False
        CFD_Relacionados.Existe = False
    End If
    Cmb_Serie.ListIndex = -1
    Cmb_FacRef.ListIndex = -1
End Sub
Public Sub Carga_Facturas()
    Dim Rs_Consulta As rdoResultset
    Cmb_FacRef.Clear
    If Cmb_Cliente.ListIndex > -1 Then
        Mi_SQL = "SELECT distinct(No_Factura) FROM Complemento_Pago WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "#00000") & "' and Timbre_UUID is not null and Timbre_UUID<>'' and Estatus='CANCELADA' and Serie='" & Cmb_Serie.text & "'"
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta.EOF
             Cmb_FacRef.AddItem (Rs_Consulta.rdoColumns("No_Factura"))
             Rs_Consulta.MoveNext
        Wend
         Rs_Consulta.Close
    Else
        MsgBox "Seleccione un cliente"
        Exit Sub
    End If
End Sub

Private Sub Cmb_Serie_Click()
    If Cmb_Serie.ListIndex > -1 And Cmb_Cliente.ListIndex > -1 Then Carga_Facturas
End Sub

Private Sub Cmb_Tipo_Cadena_Click()
    Txt_Certificado.text = ""
    Txt_Cadena.text = ""
    Txt_Sello.text = ""
    If Cmb_Tipo_Cadena.ListIndex > 0 Then
        Txt_Certificado.Enabled = True
        Txt_Cadena.Enabled = True
        Txt_Sello.Enabled = True
        Label11.Enabled = True
        Label13.Enabled = True
        Label12.Enabled = True
   Else
        Txt_Certificado.Enabled = False
        Txt_Cadena.Enabled = False
        Txt_Sello.Enabled = False
        Label11.Enabled = False
        Label13.Enabled = False
        Label12.Enabled = False
    End If
End Sub

Private Sub Cmb_Tipo_Pago_Click()
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset

    If Cmb_Tipo_Pago.ListIndex = 1 Then
        Txt_Dias.Visible = True
        ''Txt_Dias.Text = 15
        Lbl_Dias.Visible = True
        Mi_SQL = " SELECT * FROM Cat_Clientes WHERE Cliente_ID = '" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "'"
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta.EOF Then
                Txt_Dias.text = Rs_Consulta!Dias_Credito
            End If
    Else
        Txt_Dias.Visible = False
        Lbl_Dias.Visible = False
    End If
End Sub



Private Sub Form_Load()
    Me.Top = 0
    Me.Height = 9090 ' Altura del frame principal
    Me.Width = 9090 'Ancho del frame principal
    Consulta_Clientes
    DTP_Fecha.Value = Now
    Cmb_Forma_Pago.ListIndex = 0
    Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco_Des, 1, "Nombre")
    Cmb_Forma_Pago.Clear
    Call Conectar_Ayudante.Llena_Combo_Item("Clave,Clave + ' ' + Descripcion", "Cat_Formas_Pago", Cmb_Forma_Pago, 1, "Descripcion")
    'Cmb_Forma_Pago.ListIndex = 2
    'Cmb_Tipo_Pago.ListIndex = 0
    Call Conectar_Ayudante.Limpia_Variables
End Sub

Private Sub Grid_Facturas_Pagar_DblClick()
    Btn_Eliminar_Click
End Sub

Private Sub Grid_Facturas_Pendientes_DblClick()
    Btn_Agregar_Click
End Sub

Private Sub Text2_GotFocus()
    If Text2.text = "T. Cambio" Then
      Text2.text = ""
      Text2.BackColor = &H80000005
      Text2.ForeColor = &H80000008
    End If
End Sub
Private Sub Text2_LostFocus()
    If Text2.text = "" Then
      Text2.text = "T. Cambio"
      Text2.BackColor = &H8000000B
      Text2.ForeColor = &H80000011
    End If
End Sub

Private Sub Txt_Buscar_Factura_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Buscar_Factura, True)
End Sub

Private Sub Txt_Pago_Change()
'    If Val(Txt_Pago.text) > Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ",")) Then
'        MsgBox "El pago no puede ser mayor que el total a pagar"
''        Text1.text = ""
'        Exit Sub
'    Else
''        Text1.text = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ",")) - Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ",")), "#0.00")
'    End If
End Sub

Private Sub Txt_Pago_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pago, True)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Text2, True)
End Sub

Private Sub Txt_Total_Change()
'    Txt_Pago.text = Format(Conectar_Ayudante.Quitar_Caracter(Txt_Total.text, ","), "###,###.00")
End Sub


