VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Adm_Otros_Pagos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OTROS PAGOS"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Otros_Pagos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4290
      ScaleWidth      =   9690
      TabIndex        =   7
      Top             =   0
      Width           =   9690
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
         Left            =   90
         Picture         =   "Frm_Adm_Otros_Pagos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "A"
         Top             =   2250
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
         Left            =   5505
         Picture         =   "Frm_Adm_Otros_Pagos.frx":3537
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "A"
         Top             =   2250
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Otros_Pagos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del pago"
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
         Height          =   2175
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   6765
         Begin VB.ComboBox Cmb_Proveedor 
            Height          =   315
            Left            =   1275
            TabIndex        =   17
            Top             =   1440
            Width           =   5415
         End
         Begin VB.TextBox Txt_Referencia 
            Height          =   330
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1087
            Width           =   1770
         End
         Begin VB.ComboBox Cmb_Forma_Pago 
            Height          =   315
            ItemData        =   "Frm_Adm_Otros_Pagos.frx":6C36
            Left            =   1275
            List            =   "Frm_Adm_Otros_Pagos.frx":6C43
            TabIndex        =   3
            Top             =   1080
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
            Left            =   4920
            TabIndex        =   2
            Top             =   750
            Width           =   1770
         End
         Begin VB.ComboBox Cmb_Banco 
            Height          =   315
            Left            =   1275
            TabIndex        =   0
            Top             =   405
            Width           =   5415
         End
         Begin VB.TextBox Txt_Concepto 
            Height          =   315
            Left            =   1275
            MaxLength       =   100
            TabIndex        =   5
            Top             =   1800
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker DTP_Fecha 
            Height          =   315
            Left            =   1275
            TabIndex        =   1
            Top             =   735
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   40304643
            CurrentDate     =   38038
         End
         Begin VB.Label Lbl_Forma_Pago 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago"
            Height          =   240
            Index           =   3
            Left            =   135
            TabIndex        =   16
            Top             =   1132
            Width           =   1140
         End
         Begin VB.Label Lbl_Referencia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            Height          =   240
            Left            =   3900
            TabIndex        =   15
            Top             =   1132
            Width           =   840
         End
         Begin VB.Label Lbl_Concepto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beneficiario"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   14
            Top             =   1500
            Width           =   825
         End
         Begin VB.Label Lbl_Pago 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pago      $"
            Height          =   195
            Index           =   4
            Left            =   3915
            TabIndex        =   13
            Top             =   795
            Width           =   735
         End
         Begin VB.Label Lbl_Concepto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   12
            Top             =   1860
            Width           =   690
         End
         Begin VB.Label Lbl_Banco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco "
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   11
            Top             =   465
            Width           =   510
         End
         Begin VB.Label Lbl_Fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   10
            Top             =   795
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "Frm_Adm_Otros_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Formato As String

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Cmb_Banco.ListIndex = -1
        Cmb_Forma_Pago.ListIndex = -1
        Btn_Nuevo.Caption = "Capturar"
        Btn_Salir.Caption = "Cancelar"
        Fra_Otros_Pagos.Enabled = True
        Cmb_Banco.SetFocus
    Else
        Call Alta_Pago
    End If
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        Fra_Otros_Pagos.Enabled = False
        Cmb_Banco.ListIndex = -1
        Cmb_Forma_Pago.ListIndex = -1
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
Dim Despliega_Lista As Long

    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Banco.hwnd, &H14F, True, 0)
    End If
End Sub


Private Sub Cmb_Forma_Pago_Click()
Dim Mi_SQL  As String
Dim Rs_Consulta_Bancos As rdoResultset

    If Cmb_Forma_Pago.text = "Cheque" Then
        Lbl_Referencia.Caption = "No. Cheque"
        If Cmb_Banco.ListIndex > -1 Then
            Mi_SQL = " SELECT * FROM Cat_Bancos  WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "'  "
            Set Rs_Consulta_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Bancos.EOF Then
                If Not IsNull(Rs_Consulta_Bancos!Numero_Inial_Cheque) Then
                    If Val(Rs_Consulta_Bancos!Numero_Inial_Cheque) > Val(Rs_Consulta_Bancos!Consecutivo_Cheque) Then
                       Txt_Referencia.text = Val(Rs_Consulta_Bancos!Numero_Inial_Cheque)
                    Else
                       Txt_Referencia.text = Val(Rs_Consulta_Bancos!Consecutivo_Cheque) + 1
                    End If
                Else
                    Txt_Referencia.text = 1
                End If
            End If
            Rs_Consulta_Bancos.Close
        End If
    Else
        If Cmb_Forma_Pago.text = "Efectivo" Then
            Lbl_Referencia.Caption = "Recibe"
        Else
            Lbl_Referencia.Caption = "Referencia"
        End If
    End If
End Sub


Private Sub Cmb_Proveedor_KeyPress(KeyAscii As Integer)
Dim Despliega_Lista As Long

    'SE DEPLEGA LA LISTA DEL COMBO
    Despliega_Lista = SendMessageLong(Cmb_Proveedor.hwnd, &H14F, True, 0)
End Sub


Private Sub Form_Load()
    Me.Height = 3435
    Me.Width = 7050
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) \ 2
    Me.Top = (MDIFrm_Apl_Principal.Height - Me.Height) \ 4
    Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Estatus='ACTIVO' AND Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor, 1, "Estatus='ACTIVO' AND Nombre")
    DTP_Fecha.Value = Now
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Alta_Pago()
'DESCRIPCIÓN            : Realiza el alta de un pago a proveedor
'PARÁMETROS             : Generales de Adm_Proveedores_Facturas, Adm_Notas_Credito y Adm_Movimientos
'CREO                   : Julio Cruz
'FECHA_CREO             : 26 de Enero del 2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Alta_Pago()
Dim Rs_Modifica_Adm_Facturas_Proveedores As rdoResultset        '#  Modifica las facturas
Dim Rs_Agrega_Adm_Notas_Credito_Proveedores As rdoResultset     '#  Agrega las notas de credito
Dim Rs_Agrega_Adm_Movimientos As rdoResultset                   '#  Agrega las el movimiento
Dim No_Movimiento As String
Dim Pago As Double
Dim Respuesta As Integer
Dim Rs_Edita_Banco As rdoResultset
Dim Mi_SQL As String
Dim cont_Repeticiones As Integer

On Error GoTo Handler

    If Cmb_Banco.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ",")) > 0 And _
    Trim(Txt_Referencia.text) <> "" Then
        'INICIA LA TRANSACCION
        Conexion_Base.BeginTrans
        'BLOQUE PARA PAGOS
        If Cmb_Banco.ListIndex > -1 And _
        Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ",")) >= 0 And Trim(Txt_Referencia.text) <> "" Then
            'BLoque para dar de alta los movimientos
            Set Rs_Agrega_Adm_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
            'Da de alta el movimiento
            Pago = Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ","))
            No_Movimiento = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
            '#  Agrega el movimiento
            Set Rs_Agrega_Adm_Movimientos = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
            With Rs_Agrega_Adm_Movimientos
                .AddNew
                    .rdoColumns("No_Movimiento") = No_Movimiento
                    .rdoColumns("Banco_ID") = Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000")
                    .rdoColumns("Proveedor_ID") = Format(Cmb_Proveedor.ItemData(Cmb_Proveedor.ListIndex), "00000")
                    .rdoColumns("Banco") = Cmb_Banco.text
                    .rdoColumns("Fecha") = Format(DTP_Fecha.Value, "MM/dd/yyyy")
                    .rdoColumns("Estatus") = "A"
                    .rdoColumns("Concepto") = UCase("PAGO")
                    .rdoColumns("Forma_Pago") = Cmb_Forma_Pago.text
                    .rdoColumns("Referencia") = Trim(Txt_Referencia.text)
                    .rdoColumns("Tipo") = "E"
                    .rdoColumns("Cantidad") = Pago
                    If Cmb_Forma_Pago.text = "Cheque" Then
                        .rdoColumns("Beneficiario") = Cmb_Proveedor.text
                        'SE EDITA EL CONSECUTIVO DEL CHEQUE
                        Mi_SQL = " SELECT * FROM Cat_Bancos WHERE Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "' "
                        Set Rs_Edita_Banco = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        With Rs_Edita_Banco
                            If Not Rs_Edita_Banco.EOF Then
                                .Edit
                                    .rdoColumns("Consecutivo_Cheque") = Txt_Referencia.text
                                .Update
                            End If
                        End With
                        Rs_Edita_Banco.Close
                    End If
                    .rdoColumns("Beneficiario") = Cmb_Proveedor.text
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now()
                .Update
            End With
            Rs_Agrega_Adm_Movimientos.Close
        End If
        Set Rs_Agrega_Adm_Movimientos = Nothing
        Conexion_Base.CommitTrans
        If Cmb_Banco.ListIndex > -1 And Val(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ",")) > 0 And _
         Trim(Txt_Referencia.text) <> "" Then
            If Cmb_Forma_Pago.text = "Cheque" Then
                Respuesta = MsgBox("¿Desea Imprimir el Cheque?", vbYesNo + vbQuestion, " ADMINISTRACIÓN")
                If Respuesta = 6 Then
                    If Formato <> "" Then
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
        Fra_Otros_Pagos.Enabled = False
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
                If .rdoColumns("Campo") = "NOMBRE" Then Printer.Print Mid(Cmb_Proveedor.text, 1, Longitud)
                'CAMPOS PARA LA IMPRESION EN EL AREA DE POLIZA
                If .rdoColumns("Campo") = "NOMBRE_POLIZA" Then Printer.Print Mid(Cmb_Proveedor.text, 1, Longitud)
                If .rdoColumns("Campo") = "CONCEPTO" Then Printer.Print Mid(Txt_Concepto.text, 1, Longitud)
                If .rdoColumns("Campo") = "NUMERO_CHEQUE" Then Printer.Print Mid("CH. " & Txt_Referencia.text, 1, Longitud)
                If .rdoColumns("Campo") = "NOMBRE_BANCO" Then Printer.Print Mid(Cmb_Banco.text, 1, Longitud)
                'VALIDA SI LA MONEDA ES EN PESOS
                ''If UCase(Trim(Txt_Moneda.Text)) = "PESOS" Then
                    If .rdoColumns("Campo") = "CANTIDAD" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                    If .rdoColumns("Campo") = "CANTIDAD_POLIZA" Then Printer.Print Format(Txt_Pago, "###,###,##0.00")
                    If .rdoColumns("Campo") = "CANTIDAD_LETRA" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Pago.text, ","))
                ''Else
                    ''If .rdoColumns("Campo") = "CANTIDAD" Then Printer.Print Format(Txt_Total_Pesos, "###,###,###.00")
                    ''If .rdoColumns("Campo") = "CANTIDAD_LETRA" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Txt_Total_Pesos.Text, ","))
                ''End If
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

Private Sub Txt_Pago_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Pago.text, True)
End Sub


