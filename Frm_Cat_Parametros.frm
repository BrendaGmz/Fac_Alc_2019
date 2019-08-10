VERSION 5.00
Begin VB.Form Frm_Cat_Parametros 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PARÁMETROS"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
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
      Left            =   3150
      Picture         =   "Frm_Cat_Parametros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "A"
      Top             =   4290
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
      Left            =   90
      Picture         =   "Frm_Cat_Parametros.frx":36FF
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "M"
      Top             =   4290
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Parametros 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parámetros"
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
      Height          =   3795
      Left            =   120
      TabIndex        =   1
      Top             =   465
      Width           =   4425
      Begin VB.TextBox Txt_Porcentaje_IVA 
         Height          =   315
         Left            =   2175
         MaxLength       =   10
         TabIndex        =   22
         Top             =   2025
         Width           =   1125
      End
      Begin VB.TextBox Txt_Impuesto_Cedular 
         Height          =   315
         Left            =   2175
         MaxLength       =   4
         TabIndex        =   18
         Top             =   3390
         Width           =   1125
      End
      Begin VB.TextBox Txt_Retencion_Fletes 
         Height          =   315
         Left            =   2175
         MaxLength       =   4
         TabIndex        =   16
         Top             =   2490
         Width           =   1125
      End
      Begin VB.TextBox Txt_ISR 
         Height          =   315
         Left            =   2175
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2925
         Width           =   1125
      End
      Begin VB.TextBox Txt_Retencion_IVA 
         Height          =   315
         Left            =   2175
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1560
         Width           =   1125
      End
      Begin VB.TextBox Txt_Intentos_Fallidos 
         Height          =   285
         Left            =   2175
         TabIndex        =   4
         Top             =   285
         Width           =   1125
      End
      Begin VB.TextBox Txt_Vencimiento_Cuenta 
         Height          =   285
         Left            =   2175
         TabIndex        =   3
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox Txt_Cambio_Password 
         Height          =   285
         Left            =   2175
         TabIndex        =   2
         Top             =   1125
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Impuesto IVA"
         Height          =   195
         Left            =   165
         TabIndex        =   24
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   195
         Left            =   3465
         TabIndex        =   23
         Top             =   2085
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   195
         Left            =   3495
         TabIndex        =   21
         Top             =   2550
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   195
         Left            =   3480
         TabIndex        =   20
         Top             =   3450
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Impuesto Cedular"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   3450
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retencion fletes"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   2550
         Width           =   1155
      End
      Begin VB.Label Lbl_ISR 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retencion ISR"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   2985
         Width           =   1050
      End
      Begin VB.Label Lbl_Porcentaje_ISR 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   195
         Left            =   3480
         TabIndex        =   14
         Top             =   2985
         Width           =   135
      End
      Begin VB.Label Lbl_Retencion_IVA 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retencion IVA"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2085
         Width           =   1035
      End
      Begin VB.Label Lbl_Porcentaje_IVA 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   195
         Left            =   3495
         TabIndex        =   11
         Top             =   1620
         Width           =   120
      End
      Begin VB.Label Lbl_Intentos_Fallidos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Intentos Fallidos"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Lbl_Vencimiento_Cuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vencimiento Cuenta"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Lbl_Cambio_Password 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limite Cambio Password"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   1170
         Width           =   1710
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARÁMETROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1005
      TabIndex        =   0
      Top             =   45
      Width           =   2685
   End
End
Attribute VB_Name = "Frm_Cat_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Nuevo" Or Btn_Modificar.Caption = "Modificar" Then
    If Btn_Modificar.Caption = "Nuevo" Then
        Btn_Modificar.Caption = "Alta"
    Else
        Btn_Modificar.Caption = "Actualizar"
    End If
    Btn_Salir.Caption = "Cancelar"
    Fra_Parametros.Enabled = True
    Txt_Intentos_Fallidos.SetFocus
Else
    If Val(Txt_Cambio_Password.text) > 0 And Val(Txt_Intentos_Fallidos.text) > 0 And _
    Val(Txt_Vencimiento_Cuenta.text) > 0 Then
        Modifica_Parametros 'Modifica o da de alta los parámetros del sistema
    Else
        If Val(Txt_Cambio_Password.text) = 0 Then
            MsgBox "Proporcione la cantidad de dias" & Chr(13) & Chr(13) & _
                   "en que el password va hacer vigente", vbExclamation
            Txt_Cambio_Password.SetFocus
        Else
            If Val(Txt_Intentos_Fallidos.text) = 0 Then
                MsgBox "Proporcione el No. de Intentos que" & Chr(13) & Chr(13) & _
                       "puede tener el usuario para equivocarse", vbExclamation
                Txt_Intentos_Fallidos.SetFocus
            Else
                MsgBox "Proporcione el No. de Días que" & Chr(13) & Chr(13) & _
                "puede tener una cuenta sin estar en uso", vbExclamation
                Txt_Vencimiento_Cuenta.SetFocus
            End If
        End If
    End If
End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Btn_Modificar_Click
    'DESCRIPCIÓN: Modifica los valores que se tienen en los registros de la tabla
    '             de Cat_Parametros
    'PARÁMETROS :
    'CREO       : Miguel Segura
    'FECHA_CREO :24-Oct-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Parametros()
Dim Rs_Alta_Cat_Parametros_Sistema As rdoResultset 'Manejo del registro de Cat_Parametros_Sistema
Dim Rs_Modifica_Cat_Parametros As rdoResultset     'Modifica los datos de los registros que tiene la tabla
Set Conectar_Ayudante = New Ayudante

On Error GoTo Handler
    Conexion_Base.BeginTrans
    If Btn_Modificar.Caption = "Alta" Then
        'Alta de Parametros sistema
        Set Rs_Alta_Cat_Parametros_Sistema = Conectar_Ayudante.Recordset_Agregar("Cat_Parametros")
        'Llena la tabla de Cat_Parametros sistema con los datos contenidos en las cajas de textos
        With Rs_Alta_Cat_Parametros_Sistema
            .AddNew
                .rdoColumns("Parametro_ID") = "00001"
                .rdoColumns("Intentos_Fallidos") = Val(Txt_Intentos_Fallidos.text)
                .rdoColumns("Vencimiento_Cuenta_Usuario") = Val(Txt_Vencimiento_Cuenta.text)
                .rdoColumns("Limite_Cambio_Password") = Val(Txt_Cambio_Password.text)
                .rdoColumns("Retencion_IVA") = Val(Txt_Retencion_IVA.text)
                .rdoColumns("Retencion_ISR") = Val(Txt_ISR.text)
                .rdoColumns("Retencion_Fletes") = Val(Txt_Retencion_Fletes.text)
                .rdoColumns("Impuesto_Cedular") = Val(Txt_Impuesto_Cedular.text)
                .rdoColumns("Impuesto_IVA") = Val(Txt_Porcentaje_IVA.text)
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
        Rs_Alta_Cat_Parametros_Sistema.Close
    Else
        Mi_SQL = "SELECT * FROM Cat_Parametros"
        Set Rs_Modifica_Cat_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        'Llena la tabla de Cat_Parametros sistema con los datos contenidos en las cajas de textos
        With Rs_Modifica_Cat_Parametros
            .Edit
                .rdoColumns("Intentos_Fallidos") = Val(Txt_Intentos_Fallidos.text)
                .rdoColumns("Vencimiento_Cuenta_Usuario") = Val(Txt_Vencimiento_Cuenta.text)
                .rdoColumns("Limite_Cambio_Password") = Val(Txt_Cambio_Password.text)
                .rdoColumns("Retencion_IVA") = Val(Txt_Retencion_IVA.text)
                .rdoColumns("Retencion_ISR") = Val(Txt_ISR.text)
                .rdoColumns("Retencion_Fletes") = Val(Txt_Retencion_Fletes.text)
                .rdoColumns("Impuesto_Cedular") = Val(Txt_Impuesto_Cedular.text)
                .rdoColumns("Impuesto_IVA") = Val(Txt_Porcentaje_IVA.text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
        Rs_Modifica_Cat_Parametros.Close
    End If
    Conexion_Base.CommitTrans
    MsgBox "Parametros capturados", vbInformation
    Fra_Parametros.Enabled = False
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Consulta_Parametros 'Consulta los parámetros del sistema para actualizarlos de acuerdo a la modificación realizada
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Parametros", Frm_Cat_Parametros)
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
Dim Respuesta As Integer 'Indica la respuesta del usuario
If Btn_Salir.Caption = "Salir" Then
    Unload Me
Else
    Respuesta = MsgBox("Esta seguro de cancelar la operación", vbYesNo + vbQuestion)
    If Respuesta = 6 Then
        Consulta_Parametros 'Consulta los valores que tiene los parámetros
        Fra_Parametros.Enabled = False
        If Btn_Modificar.Caption = "Alta" Then
            Btn_Modificar.Caption = "Nuevo"
        Else
            Btn_Modificar.Caption = "Modificar"
        End If
        Btn_Salir.Caption = "Salir"
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Consulta_Parametros 'Consulta los parametrso que tiene el sistema
    Txt_Intentos_Fallidos = Intentos_Fallidos
    Txt_Vencimiento_Cuenta = Bloqueo_Por_No_Utilizar
    Txt_Cambio_Password.text = Bloqueo_Por_Expiración_Password
    Txt_Retencion_IVA.text = PG_Retencion_IVA
    Txt_ISR.text = PG_Retencion_ISR
    Txt_Retencion_Fletes.text = PG_Retencion_Flete
    Txt_Impuesto_Cedular.text = PG_Impuesto_Cedular
    Txt_Porcentaje_IVA.text = Porcentaje_IVA
    If Val(Txt_Intentos_Fallidos.text) > 0 Then
        Btn_Modificar.Caption = "Modificar"
    Else
        Btn_Modificar.Caption = "Nuevo"
    End If
End Sub

Private Sub Txt_Cambio_Password_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cambio_Password, False)
End Sub

Private Sub Txt_Intentos_Fallidos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Intentos_Fallidos, False)
End Sub

Private Sub Txt_Vencimiento_Cuenta_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Vencimiento_Cuenta, False)
End Sub
