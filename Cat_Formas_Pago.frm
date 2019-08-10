VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Formas_Pago 
   Caption         =   "Catalogo Formas de Pago"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   39
      Top             =   5040
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid Grid_Formas 
         Height          =   2055
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   0
         Cols            =   15
         FixedRows       =   0
         ForeColorSel    =   -2147483638
         BackColorBkg    =   16777215
      End
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2310
      Picture         =   "Cat_Formas_Pago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "M"
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7425
      Picture         =   "Cat_Formas_Pago.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4005
      Picture         =   "Cat_Formas_Pago.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "B"
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5715
      Picture         =   "Cat_Formas_Pago.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "C"
      Top             =   7680
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   600
      Picture         =   "Cat_Formas_Pago.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "A"
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Vigencia 
      Caption         =   "Fechas de vigencia"
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
      Height          =   1575
      Left            =   5160
      TabIndex        =   28
      Top             =   3480
      Width           =   3855
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   33
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio de Vigencia"
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
         Index           =   9
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Fra_Beneficiario 
      Caption         =   "Cuenta Beneficiario"
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
      Height          =   1455
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   4815
      Begin VB.ComboBox Cmb_Beneficiario 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0552
         Left            =   2520
         List            =   "Cat_Formas_Pago.frx":055F
         TabIndex        =   23
         Text            =   "Cmb_Beneficiario"
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox Cmb_Patron_Beneficiario 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0575
         Left            =   2520
         List            =   "Cat_Formas_Pago.frx":0582
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox Cmb_RFC_Beneficiario 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0598
         Left            =   2520
         List            =   "Cat_Formas_Pago.frx":05A5
         TabIndex        =   20
         Text            =   "Cmb_RFC_Beneficiario"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Patrón para cuenta"
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
         Index           =   8
         Left            =   720
         TabIndex        =   24
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta de beneficiario"
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
         Index           =   5
         Left            =   480
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "RFC del Emisor"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Fra_Ordenante 
      Caption         =   "Cuenta Ordenante"
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
      Height          =   2775
      Left            =   5160
      TabIndex        =   11
      Top             =   600
      Width           =   3855
      Begin VB.ComboBox Cmb_Banco 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":05BB
         Left            =   1920
         List            =   "Cat_Formas_Pago.frx":05C8
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox Cmb_Patron_Ordenante 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":05DE
         Left            =   1920
         List            =   "Cat_Formas_Pago.frx":05EB
         TabIndex        =   16
         Text            =   "Cmb_Patron_Ordenante"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Cmb_Ordenante 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0601
         Left            =   1920
         List            =   "Cat_Formas_Pago.frx":060E
         TabIndex        =   14
         Text            =   "Cmb_Ordenante"
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Cmb_RFC_Ordenante 
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0624
         Left            =   1920
         List            =   "Cat_Formas_Pago.frx":0631
         TabIndex        =   12
         Text            =   "Cmb_RFC_Ordenante"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre del banco emisor en caso de extranjero"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Patrón para cuenta"
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
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta ordenante"
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
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "RFC del emisor"
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
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.TextBox Txt_Forma_Pago 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      Begin VB.TextBox Txt_Descripcion 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2160
         TabIndex        =   27
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox Cmb_Cadena 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":0647
         Left            =   2160
         List            =   "Cat_Formas_Pago.frx":0654
         TabIndex        =   9
         Text            =   "Cmb_Cadena"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Bancarizado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":066A
         Left            =   2160
         List            =   "Cat_Formas_Pago.frx":0677
         TabIndex        =   7
         Text            =   "Cmb_Bancarizado"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Operacion 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Formas_Pago.frx":068D
         Left            =   2160
         List            =   "Cat_Formas_Pago.frx":069A
         TabIndex        =   6
         Text            =   "Cmb_Operacion"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo cadena de pago"
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
         Index           =   12
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Bancarizado"
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
         Index           =   7
         Left            =   1080
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Número de operación"
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
         Index           =   6
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Clave"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción "
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
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FORMAS DE PAGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Frm_Cat_Formas_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Forma_Pago.text <> "" Then
        Consulta_Formas (Txt_Forma_Pago.text)
        
    Else
        MsgBox "Ingrese la clave", vbInformation
    End If
End Sub
Public Sub Consulta_Formas(Cadena As String)
Dim Rs_Consulta_Cat_Formas As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Fra_Vigencia.Enabled = True
    Fra_Beneficiario.Enabled = True
    Fra_Ordenante.Enabled = True
    Txt_Descripcion.Enabled = True
    Cmb_Bancarizado.Enabled = True
    Cmb_Operacion.Enabled = True
    Cmb_Cadena.Enabled = True
    Grid_Formas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Formas_Pago"
    Mi_SQL = Mi_SQL & " WHERE (Clave =" & Cadena & ")"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Formas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Formas.EOF Then
        'Pone un encabezado en el grid
        Grid_Formas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Bancarizado" & Chr(9) & "No. Operación" & Chr(9) & "Cadena Pago" & Chr(9) & "RFC Emisor " & Chr(9) & "Cuenta Beneficiario" & Chr(9) & "Patron para Cuenta" & Chr(9) & "RFC Emisor Ordenante" & Chr(9) & "Cuenta Ordenante" & Chr(9) & "Patron para Cuenta" & Chr(9) & "Nombre Banco" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Formas
              Txt_Forma_Pago.text = .rdoColumns("Clave")
            Txt_Descripcion.text = .rdoColumns("Descripcion")
            Cmb_Bancarizado.text = .rdoColumns("Bancarizado")
            Cmb_Operacion.text = .rdoColumns("Numero_Operacion")
            Cmb_RFC_Ordenante.text = .rdoColumns("RFC_Emisor_Cuenta_Ordenante")
            Cmb_Ordenante.text = .rdoColumns("Cuenta_Ordenante")
            Cmb_Patron_Ordenante.text = .rdoColumns("Patron_Cuenta_Ordenante")
            Cmb_RFC_Beneficiario.text = .rdoColumns("RFC_Emisor_Cuenta_Beneficiario")
            Cmb_Beneficiario.text = .rdoColumns("Cuenta_Beneficiario")
            Cmb_Patron_Beneficiario.text = .rdoColumns("Patron_Cuenta_Beneficiario")
            Cmb_Cadena.text = .rdoColumns("Tipo_Cadena_Pago")
            Cmb_Banco.text = .rdoColumns("Banco_Emisor_Cuenta_Ordenante")
            Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
            Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
            If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                Chc_Vigencia.Value = 1
                Chc_Vigencia_Click
                Else
                Chc_Vigencia.Value = 0
                Chc_Vigencia_Click
                Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                End If
       
        'Llenado del grid
        While Not Rs_Consulta_Cat_Formas.EOF
            'With Rs_Consulta_Cat_Formas
                Grid_Formas.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Bancarizado") & Chr(9) & .rdoColumns("Numero_Operacion") & Chr(9) & .rdoColumns("Tipo_Cadena_Pago") & Chr(9) & .rdoColumns("RFC_Emisor_Cuenta_Beneficiario") & Chr(9) & .rdoColumns("Cuenta_Beneficiario") & Chr(9) & .rdoColumns("Patron_Cuenta_Ordenante") & Chr(9) & .rdoColumns("RFC_Emisor_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Cuenta_Ordenante") & Chr(9) & .rdoColumns("Patron_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Banco_Emisor_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
                Grid_Formas.FixedRows = 1
                Rs_Consulta_Cat_Formas.MoveNext
            
        Wend
        End With
        'Tamaño de las columnas en el grid
        Grid_Formas.FixedCols = 1
        Grid_Formas.ColWidth(0) = 800
        Grid_Formas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Formas.ColWidth(1) = 3000
        Grid_Formas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Formas.ColWidth(2) = 1000
        Grid_Formas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Formas.ColWidth(3) = 1200
        Grid_Formas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Formas.ColWidth(4) = 1200
        Grid_Formas.ColAlignment(4) = flexAlignLeftCenter
        Grid_Formas.ColWidth(5) = 1000
        Grid_Formas.ColAlignment(5) = flexAlignLeftCenter
        Grid_Formas.ColWidth(6) = 1500
        Grid_Formas.ColAlignment(6) = flexAlignLeftCenter
        Grid_Formas.ColWidth(7) = 1500
        Grid_Formas.ColAlignment(7) = flexAlignLeftCenter
        Grid_Formas.ColAlignment(8) = flexAlignLeftCenter
        Grid_Formas.ColWidth(8) = 1800
        Grid_Formas.ColAlignment(9) = flexAlignLeftCenter
        Grid_Formas.ColWidth(9) = 1500
        Grid_Formas.ColAlignment(10) = flexAlignLeftCenter
        Grid_Formas.ColWidth(10) = 1500
        Grid_Formas.ColAlignment(11) = flexAlignLeftCenter
        Grid_Formas.ColWidth(11) = 1200
        Grid_Formas.ColAlignment(12) = flexAlignLeftCenter
        Grid_Formas.ColWidth(12) = 1400
        Grid_Formas.ColAlignment(13) = flexAlignLeftCenter
        Grid_Formas.ColWidth(13) = 1200
        Grid_Formas.ColAlignment(14) = flexAlignLeftCenter
        Grid_Formas.ColWidth(14) = 1000
        
        
    Else
        MsgBox "La clave no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Formas.Close
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Forma_Pago.text <> "" Then
        Cambiar_Estado (Txt_Forma_Pago.text)
        
    Else
        MsgBox "Ingrese la clave", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Formas As rdoResultset   'Manejo de registro
    
    Grid_Formas.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Formas_Pago WHERE (Clave =" & Cadena & ")"
    Set Rs_Consulta_Cat_Formas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Formas.rdoColumns("Estado") = "ACTIVO" Then
        Mi_SQL = "UPDATE Cat_Formas_Pago SET Estado='INACTIVO' WHERE (Clave =" & Cadena & ")"
    Else
        Mi_SQL = "UPDATE Cat_Formas_Pago SET Estado='ACTIVO' WHERE (Clave=" & Cadena & ")"
        End If
    Set Rs_Consulta_Cat_Formas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Formas.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Actualizar" And Txt_Forma_Pago.text <> "" And Trim(Txt_Descripcion.text) <> "" And Cmb_Bancarizado.text <> "" And Cmb_Operacion.text <> "" And Cmb_Cadena.text <> "" And Cmb_RFC_Beneficiario.text <> "" And Cmb_Beneficiario.text <> "" And Cmb_Patron_Beneficiario <> "" And Cmb_RFC_Ordenante.text <> "" And Cmb_Ordenante.text <> "" And Cmb_Patron_Ordenante.text <> "" Then
                Modifica_Formas
                
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        
End Sub
Public Sub Modifica_Formas()
Dim Rs_Modificacion_Cat_Formas As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Formas_Pago"
    Mi_SQL = Mi_SQL & " WHERE Clave=" & Txt_Forma_Pago.text & ""
    Set Rs_Modificacion_Cat_Formas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Formas.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Formas
            .Edit
                .rdoColumns("Clave") = Txt_Forma_Pago.text
                .rdoColumns("Descripcion") = Txt_Descripcion.text
                .rdoColumns("Bancarizado") = Cmb_Bancarizado.text
                .rdoColumns("Numero_Operacion") = Cmb_Operacion.text
                .rdoColumns("RFC_Emisor_Cuenta_Ordenante") = Cmb_RFC_Ordenante.text
                .rdoColumns("Cuenta_Ordenante") = Cmb_Ordenante.text
                .rdoColumns("Patron_Cuenta_Ordenante") = Cmb_Patron_Ordenante.text
                .rdoColumns("RFC_Emisor_Cuenta_Beneficiario") = Cmb_RFC_Beneficiario.text
                .rdoColumns("Cuenta_Beneficiario") = Cmb_Beneficiario.text
                .rdoColumns("Patron_Cuenta_Beneficiario") = Cmb_Patron_Beneficiario.text
                .rdoColumns("Tipo_Cadena_Pago") = Cmb_Cadena.text
                .rdoColumns("Banco_Emisor_Cuenta_Ordenante") = Cmb_Banco.text
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
                If Chc_Vigencia.Value = 1 Then
                 .rdoColumns("Fecha_Fin_Vigencia") = Null
                Else
                .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
                End If
            
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Formas.Close
    MsgBox "Modificación exitosa", vbInformation
     Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
    Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Grid_Formas.Rows = 0
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Fra_Vigencia.Enabled = True
                Fra_Beneficiario.Enabled = True
                Fra_Ordenante.Enabled = True
                Txt_Descripcion.Enabled = True
                Cmb_Bancarizado.Enabled = True
                Cmb_Operacion.Enabled = True
                Cmb_Cadena.Enabled = True
                Txt_Forma_Pago.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Formas_Pago", "Clave"), "00")
                Txt_Forma_Pago.Locked = True
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion.text) <> "" And Cmb_Bancarizado.text <> "" And Cmb_Operacion.text <> "" And Cmb_Cadena.text <> "" And Cmb_RFC_Beneficiario.text <> "" And Cmb_Beneficiario.text <> "" And Cmb_Patron_Beneficiario <> "" And Cmb_RFC_Ordenante.text <> "" And Cmb_Ordenante.text <> "" And Cmb_Patron_Ordenante.text <> "" Then
                Alta_Formas
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Formas()
Dim Rs_Alta_Cat_Forma As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Clave FROM Cat_Formas_Pago where Clave =" & Trim(Txt_Forma_Pago.text) & ""
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Patente_Aduanal.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Forma = Conectar_Ayudante.Recordset_Agregar("Cat_Formas_Pago")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Forma
        .AddNew
            
            .rdoColumns("Clave") = Txt_Forma_Pago.text
            .rdoColumns("Descripcion") = Txt_Descripcion.text
            .rdoColumns("Bancarizado") = Cmb_Bancarizado.text
            .rdoColumns("Numero_Operacion") = Cmb_Operacion.text
            .rdoColumns("RFC_Emisor_Cuenta_Ordenante") = Cmb_RFC_Ordenante.text
            .rdoColumns("Cuenta_Ordenante") = Cmb_Ordenante.text
            .rdoColumns("Patron_Cuenta_Ordenante") = Cmb_Patron_Ordenante.text
            .rdoColumns("RFC_Emisor_Cuenta_Beneficiario") = Cmb_RFC_Beneficiario.text
            .rdoColumns("Cuenta_Beneficiario") = Cmb_Beneficiario.text
            .rdoColumns("Patron_Cuenta_Beneficiario") = Cmb_Patron_Beneficiario.text
            .rdoColumns("Tipo_Cadena_Pago") = Cmb_Cadena.text
            .rdoColumns("Banco_Emisor_Cuenta_Ordenante") = Cmb_Banco.text
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Forma.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Formas.Rows = 0 Then
        Grid_Formas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Bancarizado" & Chr(9) & "No. Operación" & Chr(9) & "Cadena Pago" & Chr(9) & "RFC Emisor " & Chr(9) & "Cuenta Beneficiario" & Chr(9) & "Patron para Cuenta" & Chr(9) & "RFC Emisor Ordenante" & Chr(9) & "Cuenta Ordenante" & Chr(9) & "Patron para Cuenta" & Chr(9) & "Nombre Banco" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Formas.AddItem Txt_Forma_Pago.text & Chr(9) & Txt_Descripcion.text & Chr(9) & Cmb_Bancarizado.text & Chr(9) & Cmb_Operacion.text & Chr(9) & Cmb_Cadena.text & Chr(9) & Cmb_RFC_Beneficiario.text & Chr(9) & Cmb_Beneficiario.text & Chr(9) & Cmb_Patron_Beneficiario.text & Chr(9) & Cmb_RFC_Ordenante.text & Chr(9) & Cmb_Ordenante.text & Chr(9) & Cmb_Patron_Ordenante.text & Chr(9) & Cmb_Banco.text & Chr(9) & Dtp_Inicio_Vigencia.Value & Chr(9) & "" & Chr(9) & "ACTIVO"
            Else
            Grid_Formas.AddItem Txt_Forma_Pago.text & Chr(9) & Txt_Descripcion.text & Chr(9) & Cmb_Bancarizado.text & Chr(9) & Cmb_Operacion.text & Chr(9) & Cmb_Cadena.text & Chr(9) & Cmb_RFC_Beneficiario.text & Chr(9) & Cmb_Beneficiario.text & Chr(9) & Cmb_Patron_Beneficiario.text & Chr(9) & Cmb_RFC_Ordenante.text & Chr(9) & Cmb_Ordenante.text & Chr(9) & Cmb_Patron_Ordenante.text & Chr(9) & Cmb_Banco.text & Chr(9) & Dtp_Inicio_Vigencia.Value & Chr(9) & Dtp_Fin_Vigencia.Value & Chr(9) & "ACTIVO"
           End If
    
        Grid_Formas.FixedCols = 1
        Grid_Formas.ColWidth(0) = 800
        Grid_Formas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Formas.ColWidth(1) = 3000
        Grid_Formas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Formas.ColWidth(2) = 1000
        Grid_Formas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Formas.ColWidth(3) = 1200
        Grid_Formas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Formas.ColWidth(4) = 1200
        Grid_Formas.ColAlignment(4) = flexAlignLeftCenter
        Grid_Formas.ColWidth(5) = 1000
        Grid_Formas.ColAlignment(5) = flexAlignLeftCenter
        Grid_Formas.ColWidth(6) = 1500
        Grid_Formas.ColAlignment(6) = flexAlignLeftCenter
        Grid_Formas.ColWidth(7) = 1500
        Grid_Formas.ColAlignment(7) = flexAlignLeftCenter
        Grid_Formas.ColAlignment(8) = flexAlignLeftCenter
        Grid_Formas.ColWidth(8) = 1800
        Grid_Formas.ColAlignment(9) = flexAlignLeftCenter
        Grid_Formas.ColWidth(9) = 1500
        Grid_Formas.ColAlignment(10) = flexAlignLeftCenter
        Grid_Formas.ColWidth(10) = 1500
        Grid_Formas.ColAlignment(11) = flexAlignLeftCenter
        Grid_Formas.ColWidth(11) = 1200
        Grid_Formas.ColAlignment(12) = flexAlignLeftCenter
        Grid_Formas.ColWidth(12) = 1400
        Grid_Formas.ColAlignment(13) = flexAlignLeftCenter
        Grid_Formas.ColWidth(13) = 1200
        Grid_Formas.ColAlignment(14) = flexAlignLeftCenter
        Grid_Formas.ColWidth(14) = 1000
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        Btn_Modificar.Caption = "Modificar"
        Btn_Modificar.Enabled = False
        Btn_Consultar.Enabled = True
        Btn_Eliminar.Enabled = False
        Fra_Vigencia.Enabled = False
        Fra_Beneficiario.Enabled = False
        Fra_Ordenante.Enabled = False
        Txt_Descripcion.Enabled = False
        Cmb_Bancarizado.Enabled = False
        Cmb_Operacion.Enabled = False
        Cmb_Cadena.Enabled = False
        Txt_Forma_Pago.Locked = False
        Grid_Formas.Rows = 0
        Consulta
    End If
End Sub

Private Sub Form_Load()
    Consulta
    Chc_Vigencia.Value = 1
    Chc_Vigencia_Click
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub
Private Sub Chc_Vigencia_Click()
    If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia.Enabled = False
    Else
        Dtp_Fin_Vigencia.Enabled = True
        Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
        Dtp_Fin_Vigencia.Value = Dtp_Inicio_Vigencia.Value
    End If
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Formas As rdoResultset   'Manejo de registro
    
    Grid_Formas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Formas_Pago ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Formas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Formas.EOF Then
        'Pone un encabezado en el grid
        Grid_Formas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Bancarizado" & Chr(9) & "No. Operación" & Chr(9) & "Cadena Pago" & Chr(9) & "RFC Emisor " & Chr(9) & "Cuenta Beneficiario" & Chr(9) & "Patron para Cuenta" & Chr(9) & "RFC Emisor Ordenante" & Chr(9) & "Cuenta Ordenante" & Chr(9) & "Patron para Cuenta" & Chr(9) & "Nombre Banco" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Formas.EOF
            With Rs_Consulta_Cat_Formas
                Grid_Formas.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Bancarizado") & Chr(9) & .rdoColumns("Numero_Operacion") & Chr(9) & .rdoColumns("Tipo_Cadena_Pago") & Chr(9) & .rdoColumns("RFC_Emisor_Cuenta_Beneficiario") & Chr(9) & .rdoColumns("Cuenta_Beneficiario") & Chr(9) & .rdoColumns("Patron_Cuenta_Ordenante") & Chr(9) & .rdoColumns("RFC_Emisor_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Cuenta_Ordenante") & Chr(9) & .rdoColumns("Patron_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Banco_Emisor_Cuenta_Ordenante") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
                Grid_Formas.FixedRows = 1
                Rs_Consulta_Cat_Formas.MoveNext
            End With
        Wend
        'Tamaño de las columnas en el grid
        Grid_Formas.FixedCols = 1
        Grid_Formas.ColWidth(0) = 800
        Grid_Formas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Formas.ColWidth(1) = 3000
        Grid_Formas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Formas.ColWidth(2) = 1000
        Grid_Formas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Formas.ColWidth(3) = 1200
        Grid_Formas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Formas.ColWidth(4) = 1200
        Grid_Formas.ColAlignment(4) = flexAlignLeftCenter
        Grid_Formas.ColWidth(5) = 1000
        Grid_Formas.ColAlignment(5) = flexAlignLeftCenter
        Grid_Formas.ColWidth(6) = 1500
        Grid_Formas.ColAlignment(6) = flexAlignLeftCenter
        Grid_Formas.ColWidth(7) = 1500
        Grid_Formas.ColAlignment(7) = flexAlignLeftCenter
        Grid_Formas.ColAlignment(8) = flexAlignLeftCenter
        Grid_Formas.ColWidth(8) = 1800
        Grid_Formas.ColAlignment(9) = flexAlignLeftCenter
        Grid_Formas.ColWidth(9) = 1500
        Grid_Formas.ColAlignment(10) = flexAlignLeftCenter
        Grid_Formas.ColWidth(10) = 1500
        Grid_Formas.ColAlignment(11) = flexAlignLeftCenter
        Grid_Formas.ColWidth(11) = 1200
        Grid_Formas.ColAlignment(12) = flexAlignLeftCenter
        Grid_Formas.ColWidth(12) = 1400
        Grid_Formas.ColAlignment(13) = flexAlignLeftCenter
        Grid_Formas.ColWidth(13) = 1200
        Grid_Formas.ColAlignment(14) = flexAlignLeftCenter
        Grid_Formas.ColWidth(14) = 1000
        
        
    End If
    Rs_Consulta_Cat_Formas.Close
End Sub

Private Sub Grid_Formas_Click()
Dim Rs_Consulta_Cat_Formas As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Formas.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Fra_Vigencia.Enabled = True
        Fra_Beneficiario.Enabled = True
        Fra_Ordenante.Enabled = True
        Txt_Descripcion.Enabled = True
        Cmb_Bancarizado.Enabled = True
        Cmb_Operacion.Enabled = True
        Cmb_Cadena.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Formas_Pago"
        Mi_SQL = Mi_SQL & " WHERE Clave=" & Grid_Formas.TextMatrix(Grid_Formas.RowSel, 0) & ""
        Set Rs_Consulta_Cat_Formas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Formas.EOF Then
            With Rs_Consulta_Cat_Formas
            
            Txt_Forma_Pago.text = .rdoColumns("Clave")
            Txt_Descripcion.text = .rdoColumns("Descripcion")
            Cmb_Bancarizado.text = .rdoColumns("Bancarizado")
            Cmb_Operacion.text = .rdoColumns("Numero_Operacion")
            Cmb_RFC_Ordenante.text = .rdoColumns("RFC_Emisor_Cuenta_Ordenante")
            Cmb_Ordenante.text = .rdoColumns("Cuenta_Ordenante")
            Cmb_Patron_Ordenante.text = .rdoColumns("Patron_Cuenta_Ordenante")
            Cmb_RFC_Beneficiario.text = .rdoColumns("RFC_Emisor_Cuenta_Beneficiario")
            Cmb_Beneficiario.text = .rdoColumns("Cuenta_Beneficiario")
            Cmb_Patron_Beneficiario.text = .rdoColumns("Patron_Cuenta_Beneficiario")
            Cmb_Cadena.text = .rdoColumns("Tipo_Cadena_Pago")
            Cmb_Banco.text = .rdoColumns("Banco_Emisor_Cuenta_Ordenante")
            Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
            Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
            If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                Chc_Vigencia.Value = 1
                Chc_Vigencia_Click
                Else
                Chc_Vigencia.Value = 0
                Chc_Vigencia_Click
                Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                End If
            End With
        End If
        Rs_Consulta_Cat_Formas.Close
    End If
End Sub
