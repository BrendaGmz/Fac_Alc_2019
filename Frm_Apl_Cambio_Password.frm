VERSION 5.00
Begin VB.Form Frm_Apl_Cambio_Password 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CAMBIO DE CONTRASEÑA"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   5280
   Begin VB.CommandButton Btn_Actualizar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Actualizar"
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
      Left            =   1800
      Picture         =   "Frm_Apl_Cambio_Password.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "M"
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.TextBox Txt_Contraseña_Anterior 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3045
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1050
      Width           =   1700
   End
   Begin VB.TextBox Txt_Login 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3045
      MaxLength       =   20
      TabIndex        =   0
      Top             =   675
      Width           =   1700
   End
   Begin VB.TextBox Txt_Confirmar_Contraseña_Nueva 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3045
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1785
      Width           =   1700
   End
   Begin VB.TextBox Txt_Contraseña_Nueva 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3045
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1700
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  CAMBIO DE PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   930
      TabIndex        =   8
      Top             =   165
      Width           =   3465
   End
   Begin VB.Label Lbl_Contraseña_Anterior 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Anterior"
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
      Left            =   540
      TabIndex        =   7
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label Lbl_Contraseña_Nueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Nueva"
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
      Left            =   540
      TabIndex        =   6
      Top             =   1515
      Width           =   1590
   End
   Begin VB.Label Lbl_Login 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   570
      TabIndex        =   5
      Top             =   765
      Width           =   480
   End
   Begin VB.Label Lbl_Confirmar_Contraseña_Nueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Contraseña Nueva"
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
      Left            =   540
      TabIndex        =   3
      Top             =   1860
      Width           =   2445
   End
End
Attribute VB_Name = "Frm_Apl_Cambio_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:Desbloqueo_Cuentas
    'DESCRIPCIÓN: Cambia los parametros de bloqueo a un estado que permita utilizar la cuenta
    'PARÁMETROS :
    'CREO       : Miguel Segura
    'FECHA_CREO        :26-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Actualizar_Click()
Dim Rs_Modificacion_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro de la tabla Cat_Usuarios, modifica el password del usuario logeado

Set Conectar_Ayudante = New Ayudante
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Valida que los tex no esten vacios
    If Trim(Txt_Login.Text) <> "" And Trim(Txt_Contraseña_Anterior.Text) <> "" And Trim(Txt_Contraseña_Nueva.Text) <> "" And Trim(Txt_Confirmar_Contraseña_Nueva.Text) <> "" Then
        If Conectar_Ayudante.Es_Alfanumerico(Txt_Contraseña_Nueva.Text) = True Then
            'Valida qu la contraseña sea de por lo menos 6 caracteres
            If Len(Txt_Contraseña_Nueva.Text) >= 6 Then
                'Valida que la confirmación de la contraseña sea igual a la nueva contraseña
                If Txt_Contraseña_Nueva.Text = Txt_Confirmar_Contraseña_Nueva.Text Then
                    'Consulta el Usuario actual seleccionado
                    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
                    Mi_SQL = Mi_SQL & " WHERE Login='" & Trim(Txt_Login.Text) & "' AND Password='" & Trim(Txt_Contraseña_Anterior.Text) & "' "
                    Set Rs_Modificacion_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    'Modifica los datos de la tabla Cat_Usuarios
                    With Rs_Modificacion_Apl_Cat_Usuarios
                        If Not .EOF Then
                            'Verifica que el password no sea el mismo que ya se tenia dado de alta
                            If .rdoColumns("Password") <> Txt_Contraseña_Nueva.Text Then
                                .Edit
                                    .rdoColumns("Password") = Trim(Txt_Contraseña_Nueva.Text)
                                    .rdoColumns("Estatus") = "ACTIVO"
                                    .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
                                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                    .rdoColumns("Fecha_Modifico") = Now
                                .Update
                                MsgBox "Password Modificado Exitosamente", vbInformation
                            Else
                                MsgBox "El nuevo password no puede ser el mismo que el que ha caducado!" & Chr(13) & Chr(13) & _
                                                   "Capture un password diferente al que ha caducado", vbCritical
                            End If
                        Else
                            MsgBox "No coincide el login o el password, favor de revisarlo", vbExclamation
                        End If
                    End With
                    Rs_Modificacion_Apl_Cat_Usuarios.Close
                Else
                    MsgBox "La confirmación del nuevo password no coinside con el nuevo password!" & Chr(13) & Chr(13) & _
                                               "Confirme nuevamente su password", vbCritical
                End If
            Else
                MsgBox "La longitud del password noes de por lo menos 6 caracteres!" & Chr(13) & Chr(13) & _
                                               "Capture un password de por lo menos 6 caracteres", vbCritical
            End If
        Else
            MsgBox "El password no esta compuesto por letras y numeros!" & Chr(13) & Chr(13) & _
                                               "Para obtener mayor seguridad, el password deve estar conformado por letras y numeros", vbCritical
        End If
    Else
        MsgBox "Faltan Datos para poder modificar su password!" & Chr(13) & Chr(13) & vbCritical
    End If
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Form_Load()
'Medidas del la Forma para que no puedan ser modificadas
'    Txt_Login.SetFocus
    Me.Width = 5400
    Me.Height = 3975
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub


Private Sub Txt_Confirmar_Contraseña_Nueva_KeyPress(KeyAscii As Integer)
Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub


Private Sub Txt_Contraseña_Anterior_KeyPress(KeyAscii As Integer)
Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub


Private Sub Txt_Contraseña_Nueva_KeyPress(KeyAscii As Integer)
Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub


Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub


