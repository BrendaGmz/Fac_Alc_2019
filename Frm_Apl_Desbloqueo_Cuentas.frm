VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Apl_Desbloqueo_Cuentas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Desbloqueo_Cuentas"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8550
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
      Left            =   4860
      Picture         =   "Frm_Apl_Desbloqueo_Cuentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "A"
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desbloquear"
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
      Left            =   2340
      Picture         =   "Frm_Apl_Desbloqueo_Cuentas.frx":36FF
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "M"
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Desbloqueo_Cuentas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   30
      TabIndex        =   0
      Top             =   465
      Width           =   8475
      Begin VB.Frame Fra_Usuarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   60
         TabIndex        =   9
         Top             =   1320
         Width           =   8300
         Begin MSFlexGridLib.MSFlexGrid Grid_Usuarios 
            Height          =   2295
            Left            =   105
            TabIndex        =   10
            Top             =   195
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Usuarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1050
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   8300
         Begin VB.TextBox Txt_Usuario_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1700
         End
         Begin VB.TextBox Txt_Nombre_Usuario 
            Height          =   315
            Left            =   1035
            TabIndex        =   4
            Top             =   630
            Width           =   7125
         End
         Begin VB.ComboBox Cmb_Estatus_Usuario 
            Height          =   315
            ItemData        =   "Frm_Apl_Desbloqueo_Cuentas.frx":6E30
            Left            =   6460
            List            =   "Frm_Apl_Desbloqueo_Cuentas.frx":6E3A
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   1700
         End
         Begin VB.Label Lbl_Usuario_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario ID"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Lbl_Nombre_Usuario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            TabIndex        =   7
            Top             =   690
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   5610
            TabIndex        =   6
            Top             =   300
            Width           =   645
         End
      End
   End
   Begin VB.Label Lbl_USUARIOS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desbloqueo de Cuentas"
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
      Left            =   2010
      TabIndex        =   1
      Top             =   45
      Width           =   4245
   End
End
Attribute VB_Name = "Frm_Apl_Desbloqueo_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:Desbloqueo_Cuentas
    'DESCRIPCIÓN: Cambia los parametros de bloqueo a un estado que permita utilizar la cuenta
    'PARÁMETROS :
    'CREO       : Miguel Segura
    'FECHA_CREO        :29-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Desbloqueo_Cuentas()
Dim Mi_SQL As String
Dim Rs_Modificacion_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro de la tabla Cat_Usuarios, modifica los valores del registro que tiene el usuario seleccionado

Set Conectar_Ayudante = New Ayudante
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & Trim(Txt_Usuario_ID.Text) & "'"
    Set Rs_Modificacion_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Apl_Cat_Usuarios
        .Edit
            .rdoColumns("Estatus") = Trim(Cmb_Estatus_Usuario.Text)
            .rdoColumns("Fecha_Ultimo_Acceso") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Sesion_Abierta") = "NO"
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Apl_Cat_Usuarios.Close
    Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 1) = Trim(UCase(Txt_Nombre_Usuario.Text))
    MsgBox "La cuenta se ha desbloqueado con exito", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Modificar_Click()
    'Funcion para desbloquear las uentas
    Call Desbloqueo_Cuentas
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
'Consulta los usuarios del sistema
Call Consulta_Usuarios
Me.Height = 5820
Me.Width = 8670
Me.Top = 100
Me.Left = (Screen.Width - Me.Width) / 2
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Usuarios
    'DESCRIPCIÓN: Consulta todos los Usuarios que hay en la tabla Cat_Usuarios
    '             llenando el Grid
    'PARÁMETROS : Nombre: Indica el nombre del rol que se pretende buscar
    'CREO       : Miguel Segura
    'FECHA_CREO : 26-Oct-06
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Usuarios()
Dim Rs_Consulta_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios
Set Conectar_Ayudante = New Ayudante
    
Grid_Usuarios.Rows = 0
'Consulta los datos generales del usuario
Mi_SQL = "SELECT Usuario_ID, Nombre, Login"
Mi_SQL = Mi_SQL & " FROM Apl_Cat_Usuarios"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Nombre"
Set Rs_Consulta_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Llena el grid con los datos obtenidos de la consulta
If Not Rs_Consulta_Apl_Cat_Usuarios.EOF Then
    'Coloca un encabezado en el grid
    Grid_Usuarios.AddItem "Usuario ID" & Chr(9) & "Nombre" & Chr(9) & "Login"
    While Not Rs_Consulta_Apl_Cat_Usuarios.EOF
        Grid_Usuarios.AddItem Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Usuario_ID") _
        & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Nombre") _
        & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Login")
        Grid_Usuarios.FixedRows = 1
        Rs_Consulta_Apl_Cat_Usuarios.MoveNext
    Wend
    'Configura el tamaño de las columnas del grid_usuarios
    Grid_Usuarios.ColWidth(0) = 1000
    Grid_Usuarios.ColWidth(1) = 5000
    Grid_Usuarios.ColWidth(2) = 1550
End If
'Cierra el manejador del registro
Rs_Consulta_Apl_Cat_Usuarios.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Grid_Usuarios_Click
    'DESCRIPCIÓN: Consulta todos los datos del usuario que fue seleccionado
    'PARÁMETROS :
    'CREO       : Jorge Razo
    'FECHA_CREO:
    'MODIFICO          : Yazmin Delgado Gómez
    'FECHA_MODIFICO    : 13-Oct-2007
    'CAUSA_MODIFICACIÓN: Porque se añadieron los campos de fecha a caducar y
    '                    estatus del usuario
'*******************************************************************************
Private Sub Grid_Usuarios_Click()
Dim Rs_Consulta_Alp_Cat_Usuarios As rdoResultset  'Consulta los datos del registro que fue selecciondo por el usuario

Set Conectar_Ayudante = New Ayudante
Call Conectar_Ayudante.Limpiar_Textos(Me)
'Si el grid_usuarios tiene más de un solo registro entonces consulta los datos del registro
'que fue seleccionado por el usuario
If Grid_Usuarios.Rows > 1 Then
    Txt_Usuario_ID.Text = Trim(Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 0))
    'consulta todos los valores del registro que tiene seleccionado el usuario
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID ='" & Trim(Txt_Usuario_ID.Text) & "'"
    Set Rs_Consulta_Alp_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si encuentra los valores entonces agrega los valores a los controles correspondientes de
    'la forma
    If Not Rs_Consulta_Alp_Cat_Usuarios.EOF Then
        With Rs_Consulta_Alp_Cat_Usuarios
            Txt_Nombre_Usuario.Text = .rdoColumns("Nombre")
            Cmb_Estatus_Usuario.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus_Usuario)
        End With
    End If
    Rs_Consulta_Alp_Cat_Usuarios.Close
End If
End Sub


