VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Proveedores_Anticipos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Anticipos a Proveedores"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8130
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
      Left            =   6585
      Picture         =   "Frm_Adm_Anticipos_Proveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "A"
      Top             =   3120
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
      Picture         =   "Frm_Adm_Anticipos_Proveedores.frx":36FF
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "A"
      Top             =   3105
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Datos_Anticipo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos del Anticipo"
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
      Height          =   2610
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   7815
      Begin VB.ComboBox Cmb_Proveedor 
         Height          =   315
         Left            =   1350
         TabIndex        =   8
         Top             =   240
         Width           =   6315
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   285
         Left            =   5220
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1710
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_Forma_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Anticipos_Proveedores.frx":6C36
         Left            =   1350
         List            =   "Frm_Adm_Anticipos_Proveedores.frx":6C43
         TabIndex        =   13
         Top             =   1710
         Width           =   2355
      End
      Begin VB.ComboBox Cmb_Banco 
         Height          =   315
         Left            =   1350
         TabIndex        =   11
         Top             =   720
         Width           =   6315
      End
      Begin VB.TextBox Txt_Concepto 
         Height          =   285
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   15
         Top             =   2205
         Width           =   6240
      End
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Top             =   1215
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   16449539
         CurrentDate     =   38038
      End
      Begin VB.TextBox Txt_Pago 
         Height          =   285
         Left            =   5220
         TabIndex        =   12
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Lbl_Proveedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_Referencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   3840
         TabIndex        =   6
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Lbl_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pago               $"
         Height          =   195
         Index           =   4
         Left            =   3870
         TabIndex        =   5
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label Lbl_Forma_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1710
         Width           =   1080
      End
      Begin VB.Label Lbl_Fecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   1215
         Width           =   450
      End
      Begin VB.Label Lbl_Banco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Lbl_Comentarios 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   1
         Top             =   2220
         Width           =   870
      End
   End
   Begin VB.Label Lbl_Proveedores_Anticipos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANTICIPOS A PROVEEDORES"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   4365
   End
End
Attribute VB_Name = "Frm_Adm_Proveedores_Anticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Leyenda As String
Dim Formato As String
Dim Mi_Ayudante As Ayudante

'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Imprime_Cheque(Num_Poliza,Tipo_Poliza,MesAño)
'   DESCRIPCIÓN: Consulta los formatos para la impresion de los cheques
'   PARÁMETROS:Formatos
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO:Rafael Muñoz
'   FECHA_MODIFICO:29-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Imprime_Cheque()
Dim Rs_Consulta_Formatos As rdoResultset                    'Manejo del Registro
Dim Rs_Consulta_Formatos_Generales As rdoResultset          'Manejo del Registro
Dim Rs_Consulta_Formatos_Detalles As rdoResultset           'Manejo del Registro
Dim Longitud As Integer                                     'Almacena la longitud de la cadena
Dim Inicio As Integer

    'Consulta el formato
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
    Set Rs_Consulta_Formatos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    'Consulta los generales del formato
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'General'"
    Set Rs_Consulta_Formatos_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    'Consulta los detalles del formato
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'Detalle'"
    Set Rs_Consulta_Formatos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

If Not Rs_Consulta_Formatos.EOF Then
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
    'Inicia la impresión
    'Imprime la fecha de la factura y la ciudad
    With Rs_Consulta_Formatos_Generales
        While Not .EOF
            Printer.CurrentX = .rdoColumns("X")
            Printer.CurrentY = .rdoColumns("Y")
            Longitud = .rdoColumns("Longitud")
            If .rdoColumns("Campo") = "Lugar" Then Printer.Print "IRAPUATO, GTO."
            If .rdoColumns("Campo") = "Fecha" Then Printer.Print Format(DTP_Fecha.Value, "dd-MMM-yyyy")
            If .rdoColumns("Campo") = "Nombre" Then Printer.Print Mid(Cmb_Proveedor.Text, 7, Longitud)
            If .rdoColumns("Campo") = "Cantidad" Then Printer.Print Format(Txt_Pago, "###,###,###.00")
            If .rdoColumns("Campo") = "Cantidad_Letra" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
            If .rdoColumns("Campo") = "Concepto" Then Printer.Print Mid(Txt_Concepto.Text, 1, Longitud)
            If .rdoColumns("Campo") = "Leyenda" Then
                If Leyenda = "SI" Then Printer.Print "PARA ABONO A CUENTA DEL BENEFICIARIO"
            End If
            .MoveNext
        Wend
    End With
End If
End Sub
Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Cmb_Banco.ListIndex = -1
        Cmb_Proveedor.ListIndex = -1
        Fra_Datos_Anticipo.Enabled = False
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
    End If
End Sub

'Private Sub Cmb_Banco_Click()
'Dim Rs_Consulta_Cat_Bancos As rdoResultset      '#  Consulta la informacion del banco seleccionado
'
'    If Cmb_Banco.ListIndex = -1 Then Exit Sub
'    Mi_SQL = "SELECT * FROM Cat_Bancos"
'    Mi_SQL = Mi_SQL & " WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "'"
'    Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    With Rs_Consulta_Cat_Bancos
'        If Not .EOF Then
'            Formato = .rdoColumns("Formato")
'            Txt_Cuenta_Banco = .rdoColumns("No_Cuenta")
'            If Not IsNull(.rdoColumns("Cuenta_Contable")) Then Txt_Cuenta_Banco.Text = .rdoColumns("Cuenta_Contable")
'        Else
'            Formato = ""
'            Txt_Cuenta_Banco = ""
'        End If
'        If Formato = "" And Cmb_Banco.Text <> "" Then MsgBox "Este Banco no tiene formato de impresion establecido", vbCritical, "TRANPORTES HERNIE"
'        .Close
'    End With
'    Set Rs_Consulta_Cat_Bancos = Nothing
'End Sub

'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Consulta_Proveedores()
'   DESCRIPCIÓN: Realiza una consulta para llenar el combo de proveedores
'   PARÁMETROS:Generales de Cat_Proveedores
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO:Rafael Muñoz
'   FECHA_MODIFICO:29-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Consulta_Proveedores()
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor, 1, "Nombre")
End Sub

Private Sub Btn_Sallir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Cmb_Banco.ListIndex = -1
        Cmb_Proveedor.ListIndex = -1
        Fra_Datos_Anticipo.Enabled = False
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
    End If
End Sub

Private Sub Cmb_Banco_Click()
Dim Rs_Consulta_Cat_Bancos As rdoResultset      '#  Consulta la informacion del banco seleccionado
    
    If Cmb_Banco.ListIndex = -1 Then Exit Sub
    Mi_SQL = "SELECT Formato FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "'"
    Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Bancos
        If Not .EOF Then
            Formato = .rdoColumns("Formato")
        Else
            Formato = ""
        End If
        If Formato = "" And Cmb_Banco.Text <> "" Then MsgBox "Este Banco no tiene formato de impresion establecido", vbCritical
        .Close
    End With
    Set Rs_Consulta_Cat_Bancos = Nothing
End Sub

Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Consulta_Bancos
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Forma_Pago_Change()
    If Cmb_Forma_Pago.Text = "Cheque" Then
        Lbl_Referencia.Caption = "No. Cheque"
    Else
        If Cmb_Forma_Pago.Text = "Efectivo" Then
            Lbl_Referencia.Caption = "Recibe"
        Else
            Lbl_Referencia.Caption = "Referencia"
        End If
    End If
End Sub

Private Sub Cmb_Forma_Pago_Click()
If Cmb_Forma_Pago.Text = "Cheque" Then
    Lbl_Referencia.Caption = "No. Cheque"
Else
    If Cmb_Forma_Pago.Text = "Efectivo" Then
        Lbl_Referencia.Caption = "Recibe"
    Else
        Lbl_Referencia.Caption = "Referencia"
    End If
End If
End Sub

Private Sub Cmb_Proveedor_Click()
    Dim Rs_Consulta_Cat_Proveedores As rdoResultset '#  Consulta el catalogo de proveedores
    
    '#  si no se ha seleccionado el proveedor
    If Cmb_Proveedor.ListIndex = -1 Then Exit Sub
    'Consulta datos del proveedor
    Mi_SQL = "SELECT Dias_Credito, Forma_Pago "
    Mi_SQL = Mi_SQL & " FROM Cat_Proveedores "
    Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & Format(Cmb_Proveedor.ItemData(c), "00000") & "'"
    Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Proveedores
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Forma_Pago")) Then Cmb_Forma_Pago.Text = .rdoColumns("Forma_Pago")
        End If
        .Close
    End With
    Set Rs_Consulta_Cat_Proveedores = Nothing
End Sub

Private Sub Cmb_Proveedor_KeyPress(KeyAscii As Integer)
Dim Despliega_Lista As Long

    If KeyAscii = 13 Then
        Consulta_Proveedores
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Proveedor.hwnd, &H14F, True, 0)
    End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Salir.Caption = "Cancelar"
        Fra_Datos_Anticipo.Enabled = True
        Cmb_Banco.ListIndex = -1
        Cmb_Proveedor.ListIndex = -1
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Call Conectar_Ayudante.Asigna_Item_Combo("BANAMEX", Cmb_Banco)
        Cmb_Proveedor.SetFocus
    Else
        Call Alta_Anticipo
    End If
End Sub

Private Sub Form_Initialize()
    Set Mi_Ayudante = New Ayudante
    Set Mi_Ayudante.Forma = Me
    Call Cmb_Proveedor_KeyPress(13)
    Call Cmb_Banco_KeyPress(13)
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Height = 4350
    Me.Width = 8250
    Me.Left = (Screen.Width - Me.Width) / 2
    Consulta_Bancos
    DTP_Fecha.Value = Now
End Sub
'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Consulta_Bancos()
'   DESCRIPCIÓN: Consulta para llenar el combo de bancos
'   PARÁMETROS:Generales de Cat_Bancos
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO:Rafael Muñoz
'   FECHA_MODIFICO:29-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Consulta_Bancos()
    Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Nombre")
End Sub

Private Sub Form_Resize()
    Mi_Ayudante.Redimensionar_Controles
End Sub

Private Sub Txt_Pago_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pago.Text, True)
End Sub

Private Sub Txt_Pago_LostFocus()
    Txt_Pago.Text = Format(Txt_Pago.Text, "#,###,###.00")
End Sub

'*******************************************************************************
'   NOMBRE DE LA FUNCIÓN: Alta_Anticipo()
'   DESCRIPCIÓN: Da de alta un anticipo y agrega un movimiento
'   PARÁMETROS:Generales de Adm_Proveedores_Anticipos y Adm_Movimientos
'   CREO      :Joel Romero
'   FECHA_CREO:
'   MODIFICO:Rafael Muñoz
'   FECHA_MODIFICO:29-Diciembre-2007
'   CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Private Sub Alta_Anticipo()
Dim Rs_Agrega_Adm_Movimiento As rdoResultset    '#  Agrega el anticipo movimiento
Dim Rs_Agrega_Adm_Anticipos_Proveedores As rdoResultset   '#  Agrego el anticipo
Dim MiFactura As rdoResultset
Dim Importe As Double
Dim No_Movimiento As String
Dim No_Anticipo As String
Dim Valor As Integer
Dim Descripcion_Cuenta As String
Dim Respuesta As Integer
Dim MiPoliza As rdoResultset
Dim MisDetalles As rdoResultset
Dim MiConsecutivo As rdoResultset
Dim Num_Poliza As String
Dim MesAño As String
Dim Total_Contabilizar As Double
Dim cont_Repeticiones As Integer

On Error GoTo Handler
    If Cmb_Banco.ListIndex > -1 And Cmb_Proveedor.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")) > 0 And _
    Trim(Txt_Referencia.Text) <> "" Then
        Conexion_Base.BeginTrans
        Set Rs_Agrega_Adm_Movimiento = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
        With Rs_Agrega_Adm_Movimiento
        No_Movimiento = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
            .AddNew
                .rdoColumns("No_Movimiento") = No_Movimiento
                .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                .rdoColumns("Banco_ID") = Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000")
                .rdoColumns("Banco") = Cmb_Banco.Text
                .rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                .rdoColumns("Estatus") = "A"
                .rdoColumns("Concepto") = "ANTICIPO"
                .rdoColumns("Tipo") = "E"
                .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.Text
                .rdoColumns("Referencia") = Trim(Txt_Referencia.Text)
                If Cmb_Forma_Pago.Text = "Cheque" Then .rdoColumns("Beneficiario") = Cmb_Proveedor.Text
                .rdoColumns("Comentarios") = Txt_Concepto.Text
                .rdoColumns("Cantidad") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now()
            .Update
            .Close
        End With
        Set Rs_Agrega_Adm_Movimiento = Nothing
        'Guarda el Anticipo a Proveedor
        Set Rs_Agrega_Adm_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Agregar("Adm_Proveedores_Anticipos")
        With Rs_Agrega_Adm_Anticipos_Proveedores
         No_Anticipo = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Proveedores_Anticipos", "No_Anticipo"), "0000000000")
            .AddNew
                .rdoColumns("No_Anticipo") = No_Anticipo
                .rdoColumns("No_Movimiento") = No_Movimiento
                .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                .rdoColumns("Estatus") = "ACTIVO"
                .rdoColumns("Total") = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
                .rdoColumns("Concepto_Anticipo") = Trim(Txt_Concepto.Text)
                .rdoColumns("Aplicado") = "N"
                .rdoColumns("Banco_ID") = Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000")
                .rdoColumns("Fecha") = Format(Now, "MM/dd/yyyy")
                .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.Text
                .rdoColumns("Referencia") = Trim(Txt_Referencia.Text)
                .rdoColumns("Concepto") = "ANTICIPO"
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now()
            .Update
            .Close
        End With
        Set Rs_Agrega_Adm_Anticipos_Proveedores = Nothing
          '*****************************************************************
        '**********GENERA LA POLIZA CONTABLE*****************************
'        If Txt_Cuenta_Proveedor.Text <> "" And Txt_Cuenta_Banco.Text <> "" And Cuenta_IVA_Acreditado <> "" Then
'            Total_Contabilizar = Format(Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ",")))
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
'            Cont_Partidas = 0
'            MesAño = Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 1, 2) & Mid(Format(Dtp_Fecha.Value, "MM/dd/yyyy"), 9, 2)
'            MiConsecutivo.Close
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
'                MiPoliza("Concepto") = "POLIZA DE EGRESOS"
'                MiPoliza("Total_Debe") = Total_Contabilizar
'                MiPoliza("Total_Haber") = Total_Contabilizar
'                MiPoliza("No_Partidas") = 3
'                MiPoliza("No_Movimiento") = No_Movimiento
'                MiPoliza("Usuario_Creo") = Usuario_Sistema
'                MiPoliza("Fecha_Creo") = Now
'            MiPoliza.Update  ' Save changes.
'            MiPoliza.Close
'
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
'            MisDetalles("Concepto") = Txt_Concepto.Text
'            MisDetalles("Debe") = Format((Total_Contabilizar / 1.15), "########.00")
'            MisDetalles("Haber") = 0
'            MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'            MisDetalles.Update  ' Save changes.
'           'DA DE ALTA LA PARTIDA LA IVA ACREDITABLE
'            Cont_Partidas = Cont_Partidas + 1
'            MisDetalles.AddNew  ' Create new record.
'            MisDetalles("No_Poliza") = Num_Poliza
'            MisDetalles("Tipo") = "Eg"
'            MisDetalles("Partida") = Cont_Partidas
'            MisDetalles("MesAño") = MesAño
'            MisDetalles("Cuenta") = Cuenta_IVA_Acreditado
'            MisDetalles("Concepto") = Txt_Concepto.Text
'            MisDetalles("Debe") = Format((Total_Contabilizar / 1.15) * 0.15, "########.00")
'            MisDetalles("Haber") = 0
'            MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'            MisDetalles.Update  ' Save changes.
'            'DA DE ALTA LA PARTIDA DE LOS BANCOS
'            Cont_Partidas = Cont_Partidas + 1
'            MisDetalles.AddNew  ' Create new record.
'            MisDetalles("No_Poliza") = Num_Poliza
'            MisDetalles("Tipo") = "Eg"
'            MisDetalles("Partida") = Cont_Partidas
'            MisDetalles("MesAño") = MesAño
'            MisDetalles("Cuenta") = Txt_Cuenta_Banco.Text
'            MisDetalles("Concepto") = Txt_Concepto.Text
'            MisDetalles("Debe") = 0
'            MisDetalles("Haber") = Total_Contabilizar
'            MisDetalles("Fecha") = Format(Dtp_Fecha.Value, "MM/dd/yyyy")
'            MisDetalles.Update  ' Save changes.
'            MisDetalles.Close
'        End If
        Conexion_Base.CommitTrans
        If Cmb_Forma_Pago.Text = "Cheque" Then
            If Formato <> "" Then
                Respuesta = MsgBox("¿Desea Imprimir el Cheque?", vbYesNo + vbQuestion, " ADMINISTRACIÓN")
                If Respuesta = 6 Then
                    ''Imprime_Cheque
                    For cont_Repeticiones = 1 To 2
                        Call Imprime_Cheques
                    Next
                   ' Call Impresion_Poliza(Num_Poliza, "Eg", MesAño)
                    Printer.EndDoc
                End If
            Else
                MsgBox "No se encuentra el formato de impresion del cheque", vbExclamation, "ADMINISTRACIÓN"
                Exit Sub
            End If
        End If
        MsgBox "Anticipo Capturado", vbInformation, "ADMINISTRACIÓN"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Fra_Datos_Anticipo.Enabled = False
        Cmb_Proveedor.Text = ""
        Cmb_Banco.Text = ""
    Else
        MsgBox "Datos incompletos para realizar el Anticipo", vbExclamation, "ADMINISTRACIÓN"
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
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
                'MONEDA ES EN PESOS
                If .rdoColumns("Campo") = "CANTIDAD" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                If .rdoColumns("Campo") = "CANTIDAD_POLIZA" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                If .rdoColumns("Campo") = "CANTIDAD_LETRA" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.Text, ","))
                .MoveNext
            Wend
        End With
        Printer.EndDoc
    End If
Handler:
    MsgBox Err.Description, vbCritical
    Printer.EndDoc
End Sub
