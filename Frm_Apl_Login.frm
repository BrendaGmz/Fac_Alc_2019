VERSION 5.00
Begin VB.Form Frm_Apl_Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1440
   ClientLeft      =   6270
   ClientTop       =   5265
   ClientWidth     =   3465
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFFFFF&
   FillStyle       =   2  'Horizontal Line
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Frm_Apl_Login.frx":0000
   ScaleHeight     =   850.799
   ScaleMode       =   0  'User
   ScaleWidth      =   3253.448
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txt_Login 
      Height          =   345
      Left            =   1695
      TabIndex        =   1
      Top             =   135
      Width           =   1605
   End
   Begin VB.CommandButton Btn_OK 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1710
      Picture         =   "Frm_Apl_Login.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Width           =   601
   End
   Begin VB.CommandButton Btn_Cancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2715
      Picture         =   "Frm_Apl_Login.frx":3FB1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   601
   End
   Begin VB.TextBox Txt_Password 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1695
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   1605
   End
   Begin VB.Label Lbl_Login 
      BackStyle       =   0  'Transparent
      Caption         =   "&Login"
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
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Lbl_Password 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
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
      Index           =   1
      Left            =   765
      TabIndex        =   2
      Top             =   585
      Width           =   1080
   End
End
Attribute VB_Name = "Frm_Apl_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Intentos As Integer 'Obtiene el número de intentos fallidos del usuario al acceder al sistema

'*************************************************************************************
    'NOMBRE DE LA FUNCIÓN: Actualiza_Ultimo_Acceso
    'DESCRIPCIÓN: Actualiza la fecha de ultimo acceso del usuario al sistema
    'PARÁMETROS : Login
    'CREO       : Miguel Segura Gonzalez
    'FECHA_CREO : 15-Octubre-2007
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*************************************************************************************
Private Sub Actualiza_Ultimo_Acceso()
Dim Rs_Modificar_Apl_Cat_Usuarios As rdoResultset         'Manejador de registro

On Error GoTo handler:
'Actaualiza la ultima fecha de acceso del usuario
    Mi_SQL = "SELECT Fecha_Ultimo_Acceso, Sesion_Abierta FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Login='" & Trim(UCase(Txt_Login)) & "'"
    Set Rs_Modificar_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    
    'Si durante la consulta no encontro al usuario manda un mensaje
    If Not Rs_Modificar_Apl_Cat_Usuarios.EOF Then
        With Rs_Modificar_Apl_Cat_Usuarios
            .Edit
                Rs_Modificar_Apl_Cat_Usuarios.rdoColumns("Sesion_Abierta") = "SI"
                Rs_Modificar_Apl_Cat_Usuarios.rdoColumns("Fecha_Ultimo_Acceso") = Format(Now, "MM/dd/yyyy")
            .Update
        End With
    End If
    
Exit Sub
handler:
    Debug.Print Err, Error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub


'***********************************************************************************
    'NOMBRE DE LA FUNCIÓN: Deshabilita_Usuario
    'DESCRIPCIÓN: Cambia el estatus del usuario de activo en inactivo
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 15-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'**********************************************************************************
Private Sub Deshabilita_Usuario()
Dim Rs_Modifica_Apl_Cat_Usuarios As rdoResultset 'Modifica el estatus del usuario de activo a inactivo

On Error GoTo handler:

'Consulta los datos generales del usuario al cual se le va a deshabilitar la cuenta
Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
Mi_SQL = Mi_SQL & " WHERE Login = '" & UCase(Txt_Login.text) & "'"
Set Rs_Modifica_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
'Cambia el estatus de activo a inactivo
If Not Rs_Modifica_Apl_Cat_Usuarios.EOF Then
    With Rs_Modifica_Apl_Cat_Usuarios
        .Edit
            .rdoColumns("Estatus") = "INACTIVO"
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    MsgBox "Usuario Deshabilitado", vbCritical
End If
Rs_Modifica_Apl_Cat_Usuarios.Close

Exit Sub
handler:
    Debug.Print Err, Error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub

'***********************************************************************************
    'NOMBRE DE LA FUNCIÓN: Habilitar_Configuración
    'DESCRIPCIÓN: Habilita los menus de acuerdo al usuario y sus permiso para
    '             acceder a ellos
    'PARÁMETROS : Usuario: Usuario_ID y es usada para consultar a cuales menu
    '             puede acceder el usuario
    'CREO       : Jorge Razo
    'FECHA_CREO : 12-Marzo-2005
    'MODIFICO          : Yazmin Abigail Delgado Gómez, Jorge Razo, Yazmin Delgado
    'FECHA_MODIFICO    : 16-Junio-2005, 17-Noviembre-2005, 28-Mayo-2007
    'CAUSA_MODIFICACIÓN: Porque no habilitaba los menus adecuadamente, marcaba error
    '                    al momento de deshabilitar algun menu
    '                  : Para que ocultara los menus no habiitados y no que los pusiera como deshabilitados
    '                  : Porque se cambio la forma de habilitar o deshabilitar
    '                    los menus y submenus del usuario
'**********************************************************************************
Private Function Habilita_Configuracion()
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta los menus y submenus a los cuales puede entrar el usuario
Dim Ctl As Control                              'Toma la forma del objeto al que esta apuntando en ese momento
Dim Encabezado As Integer                       'Almacenara el valor 1 al encontrar un encabezado y valida los siguientes menus para que los ocule o no

On Error GoTo handler:
    Set Conectar_Ayudante = New Ayudante
    '1. Busca en la forma si el objeto se llama menu o submenu
    '2. Por medio del Usuario_ID habilita o deshabilita los menus
    For Each Ctl In MDIFrm_Apl_Principal.Controls
        On Error Resume Next
        If UCase(Mid(Ctl.Name, 1, 4)) = "MENU" Or UCase(Mid(Ctl.Name, 1, 7)) = "SUBMENU" Then
            'Consulta que el menu que se esta seleccionado de la pantalla se encuentre
            'habilitado
            Mi_SQL = "SELECT * FROM Apl_Cat_Accesos"
            Mi_SQL = Mi_SQL & " WHERE Nombre_Sistema = '" & Ctl.Name & "'"
            Mi_SQL = Mi_SQL & " AND Rol_ID = '" & Rol_ID & "'"
            Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Apl_Cat_Accesos.EOF Then
                With Rs_Consulta_Apl_Cat_Accesos
                    If .rdoColumns("Habilitar") = "S" Then
                        Ctl.Visible = True
                    Else
                        Ctl.Visible = False
                    End If
                End With
            Else
                Ctl.Visible = False
            End If
            Rs_Consulta_Apl_Cat_Accesos.Close
        End If
    Next Ctl
    Exit Function
    
handler:
    Debug.Print Err, Error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Function

'*************************************************************************************
    'NOMBRE DE LA FUNCIÓN: Oculta_Menus
    'DESCRIPCIÓN: Dehabilita los menus más importantes del sistema para que cuando el
    '             usuario no se logie o no tenga permisos de entrar al sistema no pueda
    '             manipular la información que se existe en el sistema
    'PARÁMETROS:
    'CREO:        Jorge Razo
    'FECHA_CREO:  12-Marzo-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*************************************************************************************
'
Public Sub Oculta_Menus()
    MDIFrm_Apl_Principal.Submenu_Apl_Formato.Visible = False
    MDIFrm_Apl_Principal.Menu_Cat_Catalogos.Visible = False
    MDIFrm_Apl_Principal.Menu_Rep_Reportes.Visible = False
    MDIFrm_Apl_Principal.Menu_Ope_Ventas.Visible = False
    MDIFrm_Apl_Principal.Menu_Almacen.Visible = False
    MDIFrm_Apl_Principal.Menu_Clientes_Facturas.Visible = False
    MDIFrm_Apl_Principal.Menu_Ope_Cuentas_por_Pagar.Visible = False
    MDIFrm_Apl_Principal.Menu_Apl_Ventanas.Visible = False
End Sub

Private Sub Btn_Cancel_Click()
    Unload Frm_Apl_Login
End Sub

Private Sub Btn_OK_Click()
Dim Rs_Aceptar_Apl_Cat_Usuarios As rdoResultset 'Obtiene el login del usuario
Dim security As Integer                         '
Dim Siguiente As Integer
Set Conectar_Ayudante = New Ayudante

On Error GoTo handler:
'    Cuentas_Caducar 'Modifica las cuentas de los usuarios
    'Consulta que el usuario este dado de alta
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Login='" & UCase(Txt_Login) & "'"
    Set Rs_Aceptar_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si durante la consulta no encontro al usuario manda un mensaje
    If Rs_Aceptar_Apl_Cat_Usuarios.EOF Then
        MsgBox "Usuario Inválido, Verifique su usuario!", vbCritical, "Login"
        Txt_Login.SetFocus
        SendKeys "{Home}+{End}"
        Rs_Aceptar_Apl_Cat_Usuarios.Close
        Exit Sub
    End If
    '1. Valida primero que las cajas de texto no esten vacías y compara lo que tienen
    '2. Valida el login con la base de datos
    '3. Compara el password y manda mensaje de password incorrecto si no es igual
    '4. Si es exitoso guarda las variables globales en Usuario
    '5. Habilita loe menus de acuerdo a su seguridad
    With Rs_Aceptar_Apl_Cat_Usuarios
        If (Txt_Login.text <> "") Or (Txt_Password.text <> "") Then
            If UCase(Txt_Login) = UCase(.rdoColumns("Login")) Then
                Usuario_ID = UCase(.rdoColumns("Usuario_ID"))
                Nombre_Usuario = .rdoColumns("Nombre")
                Consulta_Parametros 'Consulta los parámetros del sistema
                If UCase(Txt_Password) = UCase(.rdoColumns("Password")) Then
                    'Valida la variable para entrar al sistema
                    If Tipo_Validacion = "Loguin" Then
''                        'DETERMINA SI EL USUARIO YA ESTA LOGEADO EN EL SISTEMA
''                        If Trim(.rdoColumns("Sesion_Abierta")) = "NO" Then
''                            'DETERMINA SI LA CUENTA SE ESTA UTILIZANDO
''                            If DateDiff("d", .rdoColumns("Fecha_Ultimo_Acceso"), Now) < Bloqueo_Por_No_Utilizar Then
                                'DETERMINA SI LA CADUSIDAD DEL PASSWORD ESTA VIGENTE
''                                If DateDiff("d", .rdoColumns("Fecha_Ultimo_Cambio_Password"), Now) < Bloqueo_Por_Expiración_Password Then
                                    If Trim(.rdoColumns("Estatus")) = "ACTIVO" Then
''                                        If CDate(Format(.rdoColumns("Fecha_Caduca"), "MM/dd/yyyy")) < CDate(Format(Now, "MM/dd/yyyy")) Then
''                                            Call Deshabilita_Usuario
''                                            MsgBox "No puede entrar al sistema ya que su cuenta caduco," & Chr(13) & Chr(13) & _
''                                                   "consulte a su administrador del sistema", vbCritical, "SISTEMA"
''                                            End
''                                        End If
                                        'ACTUALIZA LA ULTIMA FECHA DE ACCESO DEL USUARIO
                                        Call Actualiza_Ultimo_Acceso
                                        'MDIFrm_Apl_Principal.StatusBar.Panels(3).text = Nombre_Usuario
                                        Unload Frm_Apl_Login
                                        Rol_ID = .rdoColumns("Rol_ID")
                                        Call Habilita_Configuracion
                                        Consulta_Parametros_Facturacion
                                        Asignación_Datos_Cliente
                                        If Crear_ODBC = False Then
                                            MsgBox "No se ha podido crear el ODBC, favor de reportarlo al administrador", vbExclamation
                                        End If
                                    Else
                                        MsgBox "No puede entrar al sistema ya que su estatus es inactivo," & Chr(13) & Chr(13) & _
                                               "consulte a su administrador del sistema", vbCritical, "SISTEMA"
                                        End
                                    End If
''                                Else
''                                    Call Deshabilita_Usuario
''                                    MsgBox "Su Password ha caducado," & Chr(13) & Chr(13) & _
''                                               "consulte a su administrador del sistema", vbCritical, "SISTEMA"
''                                End If
''                            Else
''                                Call Deshabilita_Usuario
''                                MsgBox "Su cuenta supero el tiempo permitido de inactividad," & Chr(13) & Chr(13) & _
''                                           "consulte a su administrador del sistema", vbCritical, "SISTEMA"
''                            End If
''                        Else
''                                MsgBox "La cuenta con la que intenta registrarse en este equipo, ya ha sido utilizada para registrarse en otro equipo y no se ha cerrado la sesión, cierre la sesión en el equipo que utilizo con esta cuenta, e intente registrarse en este equipo nuevamente o" & Chr(13) & Chr(13) & _
''                                           "consulte a su administrador del sistema", vbCritical, "SISTEMA"
''                        End If
                    End If
                    'Valida la variable para desbloqueo de cuentas
                    If Tipo_Validacion = "Desbloqueo" And Rol_ID = "00001" Then
                        Unload Frm_Apl_Login
                        'Abre la ventana de desbloqueo
                        Load Frm_Apl_Desbloqueo_Cuentas
                    End If
                Else
                    Intentos = Intentos + 1
                    If Intentos_Fallidos = Intentos Then
                        Call Deshabilita_Usuario 'Deshabilita la cuenta del usuario para que no pueda acceder al sistema
                        End
                    Else
                        If Intentos_Fallidos = (Intentos + 1) Then
                            MsgBox "Invalido Password, Verifique su password!" & Chr(13) & Chr(13) & _
                                   "Le resta un intento para no inabilitar la cuenta", vbCritical
                        Else
                            MsgBox "Invalido Password, Verifique su password!", vbCritical, "Login"
                        End If
                    End If
                    Txt_Password.SetFocus
                    SendKeys "{Home}+{End}"
                    Exit Sub
                End If
            Else
                MsgBox "Invalido usuario, Veririque su usuario!", vbCritical, "Login"
                Txt_Login.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        Else
            MsgBox "Invalido Usuario, Verifique su usuario!", vbCritical, "Login"
            Txt_Login.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End With
    Rs_Aceptar_Apl_Cat_Usuarios.Close
    Exit Sub
'Declara el error en el proceso
handler:
    Debug.Print Err, Error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub

Private Sub Form_Load()
    Me.Height = 1830
    Me.Width = 3585
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = 1000
    Oculta_Menus
    Intentos = 0
End Sub

'*************************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cuentas_Caducar
    'DESCRIPCIÓN: Cambia el estatus de activo a inactivo a cuentas que caducaron en el
    '             día
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 15-Octubre-2007
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*************************************************************************************
Private Sub Cuentas_Caducar()
Dim Rs_Modifica_Apl_Cat_Usuarios As rdoResultset 'Modifica el estatus del usuario si su cuenta caduco

    'Consulta los usuarios que hayan vencido sus cuentas
    Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Estatus='ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Fecha_Caduca <" & Par_Fecha & Format(Now, "MM/dd/yyyy") & Par_Fecha
    Set Rs_Modifica_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Apl_Cat_Usuarios.EOF Then
        With Rs_Modifica_Apl_Cat_Usuarios
            While Not .EOF
                .Edit
                    .rdoColumns("Estatus") = "INACTIVO"
                .Update
                .MoveNext
            Wend
        End With
    End If
    Rs_Modifica_Apl_Cat_Usuarios.Close
End Sub
