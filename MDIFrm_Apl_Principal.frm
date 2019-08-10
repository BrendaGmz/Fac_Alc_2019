VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIFrm_Apl_Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SISTEMA INTEGRAL DE ADMINISTRACIÓN ALCOHOLERA"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13350
   Icon            =   "MDIFrm_Apl_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm_Apl_Principal.frx":076A
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   13290
      TabIndex        =   0
      Top             =   8685
      Width           =   13350
   End
   Begin VB.Menu Menu_Apl_Archivo 
      Caption         =   "&Archivo"
      WindowList      =   -1  'True
      Begin VB.Menu Submenu_Apl_Impresora 
         Caption         =   "Configurar Impresora"
      End
      Begin VB.Menu Submenu_Apl_Formato 
         Caption         =   "Configurar Formatos"
      End
      Begin VB.Menu Submenu_Apl_Calculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu Submenu_Apl_Registro 
         Caption         =   "Registro de Sistema"
      End
      Begin VB.Menu Submenu_Apl_Registrarse 
         Caption         =   "Registrarse"
      End
      Begin VB.Menu SubMenu_Cambio_Password 
         Caption         =   "Cambio de Password"
      End
      Begin VB.Menu SubMenu_Desbloqueo_Cuentas 
         Caption         =   "Desbloqueo de Cuentas"
      End
      Begin VB.Menu Submenu_Apl_Registra_Cambios_BD 
         Caption         =   "Registrar Cambios en BD"
      End
      Begin VB.Menu Raya_2 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Apl_Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Menu_Ope_Ventas 
      Caption         =   "&Ventas"
      Begin VB.Menu Submenu_Pedidos 
         Caption         =   "Pedidos"
      End
   End
   Begin VB.Menu Menu_Almacen 
      Caption         =   "&Almacén"
      Begin VB.Menu Submenu_Entradas 
         Caption         =   "Entradas"
      End
      Begin VB.Menu Submenu_Salidas 
         Caption         =   "Salidas de Almacen"
      End
   End
   Begin VB.Menu Menu_Clientes_Facturas 
      Caption         =   "Cuentas por &Cobrar"
      Begin VB.Menu Submenu_Clientes_Facturas 
         Caption         =   "Documento Cliente"
      End
      Begin VB.Menu Submenu_Clientes_Cobranza 
         Caption         =   "Cobranza"
      End
      Begin VB.Menu Raya_Cuentas_Por_Cobrar 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Notas_Credito 
         Caption         =   "Notas de Crédito"
      End
      Begin VB.Menu Linea_NC 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Clientes_Consulta_Facturas_Clientes 
         Caption         =   "Consulta Facturas Clientes"
      End
      Begin VB.Menu Submenu_Baja_Facturas 
         Caption         =   "Dar de baja Facturas"
      End
      Begin VB.Menu Submenu_Clientes_Consulta_Cobranza 
         Caption         =   "Consulta Cobranza"
      End
      Begin VB.Menu Submenu_Consulta_Cancelados 
         Caption         =   "Consulta Cancelados"
      End
   End
   Begin VB.Menu Menu_Ope_Cuentas_por_Pagar 
      Caption         =   "&Cuentas por pagar"
      Begin VB.Menu Submenu_Anticipos_Proveedores 
         Caption         =   "Anticipos Facturas"
      End
      Begin VB.Menu Submenu_Facturas_proveedores 
         Caption         =   "Facturas Proveedores"
      End
      Begin VB.Menu Raya_Pagos 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Pagos 
         Caption         =   "Pagos Facturas"
      End
      Begin VB.Menu Submenu_Otros_Pagos 
         Caption         =   "Otros Pagos"
      End
      Begin VB.Menu Raya_Cuentas_Por_PAgar 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Consulta_Factura_Proveedores 
         Caption         =   "Consulta Facturas Proveedor"
      End
      Begin VB.Menu Submenu_Consulta_Movimientos 
         Caption         =   "Consulta Pagos"
      End
   End
   Begin VB.Menu Menu_Rep_Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu Submenu_Reportes_Ventas 
         Caption         =   "Ventas"
         Begin VB.Menu Submenu_Reportes_Pedidos 
            Caption         =   "Pedidos"
         End
         Begin VB.Menu Submenu_Reportes_Pedidos_por_Documento 
            Caption         =   "Pedidos Por Documento"
         End
         Begin VB.Menu Submenu_Reporte_Productos_Documento 
            Caption         =   "Reporte de Ventas"
         End
      End
      Begin VB.Menu Submenu_Reportes_Almacen 
         Caption         =   "Almacen"
         Begin VB.Menu Submenu_Inventario_General 
            Caption         =   "Inventario General"
         End
         Begin VB.Menu Submenu_Reporte_Kardex_Producto 
            Caption         =   "Kardex  de producto"
         End
         Begin VB.Menu Submenu_Entradas_Almacen 
            Caption         =   "Entradas "
         End
         Begin VB.Menu Submenu_Reporte_Salidas 
            Caption         =   "Salidas"
         End
         Begin VB.Menu Submenu_Salidas_Almacen 
            Caption         =   "Salidas de Almacen Pendientes"
         End
      End
      Begin VB.Menu Submenu_Reportes_CXC 
         Caption         =   "Credito y Cobranza"
         Begin VB.Menu Submenu_Consulta_Cobranza 
            Caption         =   "Cobranza"
         End
         Begin VB.Menu Submenu_Consulta_Saldos_Clientes 
            Caption         =   "Saldos Clientes"
         End
         Begin VB.Menu Submenu_Rpt_Notas_Credito 
            Caption         =   "Notas de crédito"
         End
      End
      Begin VB.Menu Submenu_Facturacion 
         Caption         =   "Facturación"
         Begin VB.Menu Submenu_Reporte_Facturas_Canceladas 
            Caption         =   "Facturas canceldas"
         End
         Begin VB.Menu Submenu_Reporte_Facturas_Por_Cliente 
            Caption         =   "Facturas por cliente"
         End
         Begin VB.Menu Submenu_Reporte_General_Facturas 
            Caption         =   "Reporte General de Facuras"
         End
      End
      Begin VB.Menu Submenu_Reportes_Notas_Remision 
         Caption         =   "Notas de Remisión"
         Begin VB.Menu Submenu_Remisiones_Clientes 
            Caption         =   "Remisiones Clientes"
         End
         Begin VB.Menu Submenu_Reporte_General_Remisiones 
            Caption         =   "General de Remisiones"
         End
      End
      Begin VB.Menu Submenu_Reportes_CXP 
         Caption         =   "Cuentas por pagar"
         Begin VB.Menu Submenu_Reporte_Factuas_Vencidas 
            Caption         =   "Facturas Vencidas"
         End
         Begin VB.Menu Submenu_Reporte_Factuas_Por_Vencer 
            Caption         =   "Facturas por Vencer"
         End
         Begin VB.Menu Submenu_Facturas_Proveedor 
            Caption         =   "Facturas Proveedores"
         End
         Begin VB.Menu Submenu_Pagos_Cheques 
            Caption         =   "Pagos con cheques"
         End
         Begin VB.Menu Submenu_Saldos_Proveedor 
            Caption         =   "Reporte de Cuentas Por Pagar"
         End
      End
   End
   Begin VB.Menu Menu_Cat_Catalogos 
      Caption         =   "&Catálogos"
      Begin VB.Menu Submenu_Cat_Clasificacion_Clientes 
         Caption         =   "Clasificación de Clientes"
      End
      Begin VB.Menu Submenu_Cat_Clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Submenu_Cat_Clasificacion_Proveedor 
         Caption         =   "Clasificación de Proveedores"
      End
      Begin VB.Menu Submenu_Cat_Proveedores 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu Raya_Clientes 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Cat_Presentaciones 
         Caption         =   "Presentaciones"
      End
      Begin VB.Menu Submenu_Cat_Categorias 
         Caption         =   "Categorias"
      End
      Begin VB.Menu Submenu_Cat_Productos_Tipo 
         Caption         =   "Tipo Producto"
      End
      Begin VB.Menu Submenu_Cat_Productos 
         Caption         =   "Productos"
      End
      Begin VB.Menu Raya_Bancos 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Cat_Bancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu Raya_Presentaciones 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenu_Parametros 
         Caption         =   "Parametros"
      End
      Begin VB.Menu Submenu_Cat_Roles 
         Caption         =   "Roles"
      End
      Begin VB.Menu Submenu_Cat_Usuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu Submenu_Cat_Parametros_Facturacion 
         Caption         =   "Parametros Facturación"
      End
      Begin VB.Menu Submenu_Cat_Aduanas 
         Caption         =   "Aduanas"
      End
      Begin VB.Menu Submenu_Cat_Unidades 
         Caption         =   "Unidades"
      End
      Begin VB.Menu Submenu_Cat_Codigo_Postal 
         Caption         =   "Código Postal"
      End
      Begin VB.Menu Submenu_Cat_Impuestos 
         Caption         =   "Impuestos"
      End
      Begin VB.Menu Submenu_Cat_Metodo_Pago 
         Caption         =   "Métodos Pagos"
      End
      Begin VB.Menu Submenu_Cat_Tipo_Factor 
         Caption         =   "Tipo Factor"
      End
      Begin VB.Menu Submenu_Cat_Patentes_Aduanales 
         Caption         =   "Patentes Aduanales"
      End
      Begin VB.Menu Submenu_Cat_Monedas 
         Caption         =   "Monedas"
      End
      Begin VB.Menu Submenu_Cat_Tipos_Relacion 
         Caption         =   "Tipos Relación"
      End
      Begin VB.Menu Submenu_Cat_Tipos_Comprobantes 
         Caption         =   "Tipos Comprobantes"
      End
      Begin VB.Menu Submenu_Cat_Paises 
         Caption         =   "Países"
      End
      Begin VB.Menu Submenu_Cat_Uso_Comprobantes 
         Caption         =   "Uso Comprobantes"
      End
      Begin VB.Menu Submenu_Cat_Regimen_Fiscal 
         Caption         =   "Régimen Fiscal"
      End
      Begin VB.Menu Submenu_Cat_Formas_Pago 
         Caption         =   "Formas Pago"
      End
      Begin VB.Menu Submenu_Cat_Pedimentos_Operados 
         Caption         =   "Pedimentos Operados"
      End
      Begin VB.Menu Submenu_Cat_Tasa_Cuota 
         Caption         =   "Tasa Cuota"
      End
   End
   Begin VB.Menu Menu_Apl_Ventanas 
      Caption         =   "&Ventanas"
      Begin VB.Menu Submenu_Apl_Horizontal 
         Caption         =   "Horizontal"
      End
      Begin VB.Menu Submenu_Apl_Vertical 
         Caption         =   "Vertical"
      End
      Begin VB.Menu Submenu_Apl_Cascada 
         Caption         =   "Cascada"
      End
   End
   Begin VB.Menu Menu_Apl_Acerca_de 
      Caption         =   "Ac&erca de..."
      Begin VB.Menu Submenu_Apl_Sistema 
         Caption         =   "Sistema"
      End
   End
End
Attribute VB_Name = "MDIFrm_Apl_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Set Conectar_Ayudante = New Ayudante
    Conectar_Ayudante.Conexion 'Manda llamar a la conexión a la base desde Ayudante
End Sub

Private Sub MDIForm_Terminate()
'Cierra la sesion del susurio logeado
Call Cerrar_Sesion_Usuario
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'Cierra la sesion del susurio logeado
    Call Cerrar_Sesion_Usuario
End Sub

'*************************************************************************************
'NOMBRE DE LA FUNCIÓN: Cerrar_Sesion_Usuario
'DESCRIPCIÓN: Actualiza la fecha de ultimo acceso del usuario al sistema
'PARÁMETROS : Login
'CREO       : Miguel Segura Gonzalez
'FECHA_CREO : 29-Octubre-2007
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*************************************************************************************
Private Sub Cerrar_Sesion_Usuario()
Dim Mi_SQL  As String                                   'Almacena la consulta
Dim Rs_Modificar_Apl_Cat_Usuarios As rdoResultset         'Manejador de registro

On Error GoTo handler:
'Actaualiza la ultima fecha de acceso del usuario
    Mi_SQL = "SELECT Sesion_Abierta FROM Apl_Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & Usuario_ID & "'"
    Set Rs_Modificar_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Si durante la consulta no encontro al usuario manda un mensaje
    If Not Rs_Modificar_Apl_Cat_Usuarios.EOF Then
        With Rs_Modificar_Apl_Cat_Usuarios
            .Edit
                Rs_Modificar_Apl_Cat_Usuarios.rdoColumns("Sesion_Abierta") = "NO"
            .Update
        End With
    End If
    Rs_Modificar_Apl_Cat_Usuarios.Close  'Se cierra el recordset
    Exit Sub
handler:
    Debug.Print Err, Error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub
Private Sub Submenu_Anticipos_Proveedores_Click()
    Load Frm_Adm_Proveedores_Anticipos
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Sub Submenu_Anticipos_Proveedores", Frm_Adm_Proveedores_Anticipos)
End Sub

Private Sub Submenu_Apl_Cascada_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    MDIFrm_Apl_Principal.Arrange vbCascade
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub



Private Sub Submenu_Apl_Registra_Cambios_BD_Click()
    Unload Frm_Apl_Cambios_BD
    Load Frm_Apl_Cambios_BD
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Apl_Registra_Cambios_BD", Frm_Apl_Cambios_BD)
End Sub

Private Sub Submenu_Baja_Facturas_Click()
    Load Frm_Baja_Facturas
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Baja_Facturas", Frm_Baja_Facturas)
End Sub

Private Sub SubMenu_Cambio_Password_Click()
    'Carga la forma de cambio de password
    Load Frm_Apl_Cambio_Password
End Sub



Private Sub SubMenu_Cat_Almacenes_Click()
    Unload Frm_Cat_Generales
    Catalogo = "ALMACENES"
    Load Frm_Cat_Generales
    Call Conectar_Ayudante.Cargar_Picture(Frm_Cat_Generales.Pic_Cat_Almacenes, Frm_Cat_Generales)
    Frm_Cat_Generales.Caption = "ALMACENES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Almacenes", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Aduanas_Click()
    Frm_Cat_Aduanas.Show
End Sub

Private Sub Submenu_Cat_Bancos_Click()
Set Conectar_Ayudante = New Ayudante
    Unload Frm_Cat_Generales
    Catalogo = "BANCOS" 'Catalogo de Bancos
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Bancos, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE BANCOS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Bancos", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Categorias_Click()
    Unload Frm_Cat_Generales
    Catalogo = "CATEGORIAS" 'Catalogo de CATEGORIAS
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Categorias, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE CATEGORIAS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Categorias", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Clasificacion_Clientes_Click()
    Unload Frm_Cat_Generales
    Catalogo = "CLASIFICACION_CLIENTES"
    Load Frm_Cat_Generales
    Call Conectar_Ayudante.Cargar_Picture(Frm_Cat_Generales.Pic_Clasificacion_Clientes, Frm_Cat_Generales)
    Frm_Cat_Generales.Caption = "CLASIFICACIÓN DE CLIENTES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Clientes", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Clasificacion_Proveedor_Click()
    Unload Frm_Cat_Generales
    Catalogo = "CLASIFICACION_PROVEEDORES"
    Load Frm_Cat_Generales
    Call Conectar_Ayudante.Cargar_Picture(Frm_Cat_Generales.Pic_Clasificacion_Proveedores, Frm_Cat_Generales)
    Frm_Cat_Generales.Caption = "CLASIFICACIÓN DE PROVEEDORES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clasificacion_Proveedor", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Clientes_Click()
Set Conectar_Ayudante = New Ayudante
    Unload Frm_Cat_Clientes
    Catalogo = "CLIENTES" 'Catalogo de Clientes
    Load Frm_Cat_Clientes
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Clientes.Pic_Clientes, Frm_Cat_Clientes
    Frm_Cat_Clientes.Caption = "CATALOGO DE CLIENTES"
    Call Conectar_Ayudante.Llena_Combo_Item("Clasificacion_ID,Nombre", "Cat_Clientes_Clasificacion", Frm_Cat_Clientes.Cmb_Clasificacion_Clientes, 1, " Estatus='ACTIVO' AND Nombre")
    Frm_Cat_Clientes.Consulta_Clientes ("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Clientes", Frm_Cat_Clientes)
End Sub

Private Sub Submenu_Apl_Formato_Click()
    Frm_Cfg_Impresion.Show
End Sub

Private Sub Submenu_Apl_Horizontal_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    MDIFrm_Apl_Principal.Arrange vbTileHorizontal
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Submenu_Apl_Registrarse_Click()
    'Inicial la variable para entrar al sistema
    Tipo_Validacion = "Loguin"
    'Cierra la sesion del susurio logeado
    Call Cerrar_Sesion_Usuario
    'Muestra la ventana de registro
    Frm_Apl_Login.Show
End Sub

Private Sub Submenu_Apl_Registro_Click()
    Load Frm_Apl_Registro_Sistema
End Sub

Private Sub Submenu_Apl_Salir_Click()
    'Cierra la sesion del susurio logeado
    Call Cerrar_Sesion_Usuario
    End
End Sub

Private Sub Submenu_Apl_Calculadora_Click()
Dim szfilename As String
Dim nLength As Long
Dim retval As Variant
Const MAX_PATH = 255
    szfilename = Space(MAX_PATH)
    nLength = GetWindowsDirectory(szfilename, Len(szfilename))
    'Indica si no existe el Kernel 32
    If nLength = 0 Then
        MsgBox "Unable to Obtain the Windows Directory"
    End If
    szfilename = Left$(szfilename, nLength) & "\SYSTEM32\CALC.exe"
    retval = Shell(szfilename, 1)
End Sub

Private Sub Submenu_Apl_Impresora_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    CommonDialog1.ShowPrinter
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Submenu_Apl_Sistema_Click()
    Frm_Apl_Acerca_de.Show
End Sub

Private Sub Submenu_Cat_Laboratorios_Click()
    Unload Frm_Cat_Generales
    Catalogo = "LABORATORIOS" 'Catalogo de LABORARTORIOS
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Laboratorios, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE LABORATORIOS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Laboratorios", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Marcas_Click()
    Unload Frm_Cat_Generales
    Catalogo = "MARCAS" 'Catalogo de MARCAS
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Marcas, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE MARCAS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Marcas", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Codigo_Postal_Click()
    Frm_Cat_Codigo_Postal.Show
End Sub

Private Sub Submenu_Cat_Formas_Pago_Click()
    Frm_Cat_Formas_Pago.Show
End Sub

Private Sub Submenu_Cat_Impuestos_Click()
    Frm_Cat_Impuestos.Show
End Sub

Private Sub Submenu_Cat_Metodo_Pago_Click()
    Frm_Cat_Metodo_Pago.Show
End Sub

Private Sub Submenu_Cat_Monedas_Click()
    Frm_Cat_Monedas.Show
End Sub

Private Sub Submenu_Cat_Paises_Click()
    Frm_Cat_Paises.Show
End Sub

Private Sub Submenu_Cat_Parametros_Facturacion_Click()
    Unload Frm_Cat_Parametros_Factura_Electronica
    Load Frm_Cat_Parametros_Factura_Electronica
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Parametros_Facturacion", Frm_Cat_Parametros_Factura_Electronica)
    If Usuario_ID = "00006" Then
        Frm_Cat_Parametros_Factura_Electronica.Btn_Nuevo_Serie_Folios.Enabled = True
        Frm_Cat_Parametros_Factura_Electronica.Btn_Modificar_Serie_Folios.Enabled = True
        Frm_Cat_Parametros_Factura_Electronica.Btn_Eliminar_Serie_Folios.Enabled = True
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Nuevo_Serie_Folios.Enabled = True
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = True
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = True
    Else
        Frm_Cat_Parametros_Factura_Electronica.Btn_Nuevo_Serie_Folios.Enabled = False
        Frm_Cat_Parametros_Factura_Electronica.Btn_Modificar_Serie_Folios.Enabled = False
        Frm_Cat_Parametros_Factura_Electronica.Btn_Eliminar_Serie_Folios.Enabled = False
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Nuevo_Serie_Folios.Enabled = False
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Modificar_Serie_Folios.Enabled = False
        Frm_Cat_Parametros_Factura_Electronica.Btn_Notas_Credito_Eliminar_Serie_Folios.Enabled = False
    End If
End Sub

Private Sub Submenu_Cat_Patentes_Aduanales_Click()
    Frm_Cat_Patentes_Aduanales.Show
End Sub

Private Sub Submenu_Cat_Pedimentos_Operados_Click()
    Frm_Cat_Pedimentos_Operados.Show
End Sub

Private Sub Submenu_Cat_Presentaciones_Click()
    Unload Frm_Cat_Generales
    Catalogo = "PRESENTACIONES" 'Catalogo de Presentaciones
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Presentaciones, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE PRESENTACIONES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Presentaciones", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Productos_Click()
    Unload Frm_Cat_Clientes
    Catalogo = "PRODUCTOS" 'Catalogo de PRODUCTOS
    Load Frm_Cat_Clientes
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Clientes.Pic_Cat_Productos, Frm_Cat_Clientes
    Frm_Cat_Clientes.Caption = "CATALOGO DE PRODUCTOS"
    Call Conectar_Ayudante.Llena_Combo_Item("Presentacion_ID,Nombre", "Cat_Presentaciones", Frm_Cat_Clientes.Cmb_Presentaciones_Cat_Productos, 1, " Estatus='ACTIVO' AND Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Tipo_ID,Nombre", "Cat_Productos_Tipo", Frm_Cat_Clientes.Cmb_Cat_Producto_Tipo, 1, " Estatus='ACTIVO' AND Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Categoria_ID,Nombre", "Cat_Categorias", Frm_Cat_Clientes.Cmb_Cat_Productos_Categorias, 1, " Estatus='ACTIVO' AND Nombre")
    Call Frm_Cat_Clientes.Consulta_Cat_Productos("", "CLAVE")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos", Frm_Cat_Clientes)
End Sub

Private Sub Submenu_Cat_Productos_Tipo_Click()
    Unload Frm_Cat_Generales
    Catalogo = "PRODUCTOS_TIPO" 'Catalogo de CAT_PRODUCTOS_TIPO
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Cat_Productos_Tipo, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE TIPOS PRODUCTOS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Productos_Tipo", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Proveedores_Click()
    Set Conectar_Ayudante = New Ayudante
    Unload Frm_Cat_Clientes
    Catalogo = "PROVEEDORES" 'Catalogo de Proveedores
    Load Frm_Cat_Clientes
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Clientes.Pic_Proveedores, Frm_Cat_Clientes
    Frm_Cat_Clientes.Caption = "CATALOGO DE PROVEEDORES"
    Call Conectar_Ayudante.Llena_Combo_Item("Clasificacion_ID,Nombre", "Cat_Clasificacion_Proveedores WHERE Estatus='ACTIVO'  Order by Nombre", Frm_Cat_Clientes.Cmb_Clasificacion_Proveedor, 0, " Estatus='ACTIVO' AND Nombre")
    Frm_Cat_Clientes.Consulta_Proveedor ("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Proveedores", Frm_Cat_Clientes)
End Sub

Private Sub Submenu_Cat_Regimen_Fiscal_Click()
    Frm_Cat_Regimen_Fiscal.Show
End Sub

Private Sub Submenu_Cat_Roles_Click()
    Unload Frm_Cat_Generales
    Catalogo = "ROLES"
    Load Frm_Cat_Generales
    Call Conectar_Ayudante.Cargar_Picture(Frm_Cat_Generales.Pic_Apl_Cat_Roles, Frm_Cat_Generales)
    Frm_Cat_Generales.Caption = "CATALOGO DE ROLES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Roles", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Sustancia_Activa_Click()
    Unload Frm_Cat_Generales
    Catalogo = "SUSTANCIA_ACTIVA" 'Catalogo de sustancia activa
    Load Frm_Cat_Generales
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Cat_Sustancia_Activa, Frm_Cat_Generales
    Frm_Cat_Generales.Caption = "CATALOGO DE SUSTANCIA ACTIVA"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Sustancia_Activa", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Tasa_Cuota_Click()
    Frm_Cat_Tasa_Cuota.Show
End Sub

Private Sub Submenu_Cat_Tipo_Factor_Click()
    Frm_Cat_Tipo_Factor.Show
End Sub

Private Sub Submenu_Cat_Tipos_Comprobantes_Click()
    Frm_Cat_Tipos_Comprobantes.Show
End Sub

Private Sub Submenu_Cat_Tipos_Relacion_Click()
    Frm_Cat_Tipos_Relacion.Show
End Sub

Private Sub Submenu_Cat_Unidades_Click()
    Frm_Cat_Unidades.Show
End Sub

Private Sub Submenu_Cat_Uso_Comprobantes_Click()
    Frm_Cat_Uso_Comprobantes.Show
End Sub

Private Sub Submenu_Cat_Usuarios_Click()
    Unload Frm_Cat_Generales
    Catalogo = "USUARIOS"
    Load Frm_Cat_Generales
    Call Conectar_Ayudante.Cargar_Picture(Frm_Cat_Generales.Pic_Apl_Cat_Usuarios, Frm_Cat_Generales)
    Frm_Cat_Generales.Caption = "CATALOGO DE USUARIOS"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Usuarios", Frm_Cat_Generales)
End Sub
Private Sub Submenu_Apl_Vertical_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    MDIFrm_Apl_Principal.Arrange vbTileVertical
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub



Private Sub Submenu_Clientes_Cobranza_Click()
    Unload Frm_Adm_Cobranza
    Load Frm_Adm_Cobranza
    Frm_Adm_Cobranza.Caption = "COBRANZA"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cobranza", Frm_Adm_Cobranza)
End Sub

Private Sub Submenu_Clientes_Consulta_Cobranza_Click()
    Load Frm_Adm_Cobranza_Consulta
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Consulta_Cobranza", Frm_Adm_Cobranza_Consulta)
End Sub

Private Sub Submenu_Clientes_Consulta_Facturas_Clientes_Click()
    Load Frm_Adm_Clientes_Facturas_Consulta
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Consulta_Facturas_Clientes", Frm_Adm_Clientes_Facturas_Consulta)
End Sub


Private Sub Submenu_Clientes_Facturas_Click()
    Unload Frm_Adm_Clientes_Facturas
    Load Frm_Adm_Clientes_Facturas
    Frm_Adm_Clientes_Facturas.Caption = "FACTURAS CLIENTES"
End Sub

Private Sub Submenu_Consulta_Cancelados_Click()
    Consulta_Cancelados
    MsgBox "Actualización exitosa"
End Sub

Private Sub Submenu_Consulta_Cobranza_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Consulta_Cobranza.Visible = True
    Frm_Rpt_Reportes.Fra_Consulta_Cobranza.Caption = "Reporte de Cobranza"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Cobranza.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Cobranza.Value = Now
End Sub

Private Sub Submenu_Consulta_Factura_Proveedores_Click()
    Load Frm_Adm_Proveedores_Facturas
    Frm_Adm_Proveedores_Facturas.Pic_Busqueda_Facturas_Proveedores.Visible = True
    Frm_Adm_Proveedores_Facturas.Caption = "CONSULTA FACTURAS DE PROVEEDORES"
    Frm_Adm_Proveedores_Facturas.Btn_Nuevo.Enabled = False
    Frm_Adm_Proveedores_Facturas.Btn_Buscar.Enabled = False
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Consulta_Factura_Proveedores", Frm_Adm_Proveedores_Facturas)
End Sub

Private Sub Submenu_Consulta_Movimientos_Click()
    Load Frm_Adm_Movimientos_Consulta
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Consulta_Movimientos", Frm_Adm_Movimientos_Consulta)
End Sub




Private Sub Submenu_Consulta_Saldos_Clientes_Click()
    Catalogo = "Reporte de Saldo por Cliente"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Saldos.Visible = True
    Frm_Rpt_Reportes.Chk_Imprimir_Saldos_ceros.Visible = False
    Frm_Rpt_Reportes.Fra_Saldos.Caption = "Reporte de Saldo por Cliente"
    Frm_Rpt_Reportes.Lbl_Cliente_Saldo.Caption = "Cliente"
    Frm_Rpt_Reportes.Cmb_Cliente_Saldo.SetFocus
    Frm_Rpt_Reportes.Dtp_Fecha_Saldo_Inicio.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Saldo_Fin.Value = Now
End Sub

Private Sub SubMenu_Desbloqueo_Cuentas_Click()
    'Inicial la variable para entrar al sistema
    Tipo_Validacion = "Desbloqueo"
    Load Frm_Apl_Login
End Sub



Private Sub Submenu_Entradas_Almacen_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Entradas_Salidas.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Entradas_Salidas.Caption = "Reporte de Entradas de Productos"
    Frm_Rpt_Reportes.Cmb_Proveedor_Entrada.Visible = False
    Frm_Rpt_Reportes.Dtp_Fecha_Entrada_Salida_Inicio.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Entrada_Salida_Fin.Value = Now
    ''Frm_Rpt_Reportes.Cmb_Proveedor_Entrada.SetFocus
End Sub

Private Sub Submenu_Entradas_Click()
    Unload Frm_Alm_Entradas
    Load Frm_Alm_Entradas
    Frm_Alm_Entradas.Caption = "ENTRADAS ALMACEN"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Entradas", Frm_Alm_Entradas)
End Sub





Private Sub Submenu_Facturas_proveedor_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Facturas_Proveedor.Visible = True
    Frm_Rpt_Reportes.Fra_Facturas_Proveedor.Caption = "Reporte de Facturas Proveedores"
    Frm_Rpt_Reportes.Fra_Facturas_Proveedor.Caption = "Facturas por Proveedor"
    Frm_Rpt_Reportes.Cmb_Facturas_Proveedores.SetFocus
    Frm_Rpt_Reportes.DTP_Fecha_Inicio_Facturas_Proveedores.Value = Now
    Frm_Rpt_Reportes.DTP_Fecha_Final_Facturas_Proveedores.Value = Now
End Sub

Private Sub Submenu_Facturas_proveedores_Click()
    Unload Frm_Adm_Proveedores_Facturas
    Load Frm_Adm_Proveedores_Facturas
    Frm_Adm_Proveedores_Facturas.Caption = "FACTURAS PROVEEDORES"
   '' Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Facturas_Proveedores", Frm_Adm_Proveedores_Facturas)
End Sub





























Private Sub Submenu_Inventario_General_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Inventarios.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Inventarios.Caption = "Inventario General"
End Sub

Private Sub Submenu_Notas_Credito_Click()
    Unload Frm_Adm_Notas_Credito
    Load Frm_Adm_Notas_Credito
End Sub

Private Sub Submenu_Otros_Pagos_Click()
    Load Frm_Adm_Otros_Pagos
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Otros_Pagos", Frm_Adm_Otros_Pagos)
    Frm_Adm_Otros_Pagos.Caption = "Otros Pagos"
End Sub

Private Sub Submenu_Pagos_Cheques_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Pagos_Cheques.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Pagos_Cheques.Caption = "Reporte pagos con cheques"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Pagos_Cheques.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Pagos_Fechas.Value = Now
End Sub

Private Sub Submenu_Pagos_Click()
    Load Frm_Adm_Pagos
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Pagos", Frm_Adm_Pagos)
End Sub

Private Sub SubMenu_Parametros_Click()
    Unload Frm_Cat_Parametros
    Catalogo = "PARAMETROS"
    Load Frm_Cat_Parametros
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Parametros", Frm_Cat_Parametros)
End Sub






Private Sub Submenu_Pedidos_Click()
    Unload Frm_Ope_Pedidos
    Load Frm_Ope_Pedidos
    Frm_Ope_Pedidos.Caption = "PEDIDOS CLIENTES"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Pedidos", Frm_Ope_Pedidos)
    Frm_Ope_Pedidos.Btn_Imprimir.Enabled = False
    Frm_Ope_Pedidos.Btn_Modificar.Enabled = False
End Sub

Private Sub Submenu_Remisiones_Clientes_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Caption = "Remisiones Cliente"
    Frm_Rpt_Reportes.Cmb_Cliente_Factura.SetFocus
    Frm_Rpt_Reportes.DTP_Fecha_Inicio_Historico_Factura.Value = Now
    Frm_Rpt_Reportes.DTP_Fecha_Final_Historico_Factura.Value = Now
End Sub

Private Sub Submenu_Reporte_Factuas_Por_Vencer_Click()
    Catalogo = "Facturas_Por_Vencer"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    If Catalogo = "No_Facturas_Vencer" Then
        Catalogo = ""
        Unload Frm_Rpt_Reportes
    End If
End Sub

Private Sub Submenu_Reporte_Factuas_Vencidas_Click()
    Catalogo = ""
    Catalogo = "Facturas_Vencidas"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    If Catalogo = "No_Facturas_Vencidas" Then
        Catalogo = ""
        Unload Frm_Rpt_Reportes
    End If
End Sub

Private Sub Submenu_Reporte_Facturas_Canceladas_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Caption = "Facturas Canceladas"
    Frm_Rpt_Reportes.Cmb_Cliente_Factura.Visible = False
    Frm_Rpt_Reportes.Lbl_Reporte_hitorico_Facturas.Visible = False
    Frm_Rpt_Reportes.DTP_Fecha_Inicio_Historico_Factura.Value = Now
    Frm_Rpt_Reportes.DTP_Fecha_Final_Historico_Factura.Value = Now
End Sub

Private Sub Submenu_Reporte_Facturas_Por_Cliente_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Historico_Facturas.Caption = "Facturas Cliente"
    Frm_Rpt_Reportes.Cmb_Cliente_Factura.SetFocus
    Frm_Rpt_Reportes.DTP_Fecha_Inicio_Historico_Factura.Value = Now
    Frm_Rpt_Reportes.DTP_Fecha_Final_Historico_Factura.Value = Now
End Sub

Private Sub Submenu_Reporte_General_Facturas_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_General_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_General_Facturas.Caption = "Reporte General de Facturas"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Reporte_General_Facturas.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Reporte_General_Facturas.Value = Now
End Sub


Private Sub Submenu_Reporte_General_Remisiones_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_General_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_General_Facturas.Caption = "Reporte General de Remisiones"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Reporte_General_Facturas.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Reporte_General_Facturas.Value = Now
End Sub


Private Sub Submenu_Reporte_Kardex_Producto_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Width = 6000
    Frm_Rpt_Reportes.Fra_Kardex.Visible = True
    Frm_Rpt_Reportes.Fra_Kardex.Caption = "Kardex de Materia Prima"
    Frm_Rpt_Reportes.Cmb_Producto.SetFocus
    Frm_Rpt_Reportes.Dtp_Kardex_Fecha_Inicial.Value = Now
    Frm_Rpt_Reportes.Dtp_Kardex_Fecha_Final.Value = Now
End Sub

Private Sub Submenu_Reporte_Productos_Documento_Click()
    Catalogo = "Reporte de Ventas"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Producto_Por_Documento.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Producto_Por_Documento.Caption = "Reporte de Ventas"
    Frm_Rpt_Reportes.Cmb_Producto_Reporte_Producto_Por_Documento.SetFocus
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Reporte_Producto_Por_Documento.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Reporte_Producto_Por_Documento.Value = Now
    Catalogo = ""
End Sub

Private Sub Submenu_Reporte_Salidas_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Entradas_Salidas.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Entradas_Salidas.Caption = "Reporte de Salidas de Productos"
    Frm_Rpt_Reportes.Cmb_Proveedor_Entrada.Visible = False
    Frm_Rpt_Reportes.Dtp_Fecha_Entrada_Salida_Inicio.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Entrada_Salida_Fin.Value = Now
    ''Frm_Rpt_Reportes.Cmb_Proveedor_Entrada.SetFocus
End Sub

Private Sub Submenu_Reportes_Pedidos_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Pedidos_Clientes.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Pedidos_Clientes.Caption = "Reporte de Pedidos"
    Frm_Rpt_Reportes.Lbl_Reporte_Pedido_Producto.Caption = "Clientes"
    Frm_Rpt_Reportes.Cmb_Reporte_Pedido_Producto.SetFocus
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Pedido_Producto.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Pedidos_Productos.Value = Now
End Sub

Private Sub Submenu_Reportes_Pedidos_por_Documento_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Pedidos_Facturas.Visible = True
    Frm_Rpt_Reportes.Fra_Pedidos_Facturas.Caption = "Reporte de Documentos Por Pedidos"
    Frm_Rpt_Reportes.DTP_Fecha_Inicial_Reporte_documentos_Pedidos.Value = Now
    Frm_Rpt_Reportes.DTP_Fecha_Fin_Reporte_documentos_Pedidos.Value = Now
End Sub

Private Sub Submenu_Rpt_Notas_Credito_Click()
    Catalogo = "Reporte_Notas_Credito"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Notas_Credito.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Notas_Credito.Caption = "Reporte de Notas de Crédito Generadas"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Nota_Credito.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Nota_Credito.Value = Now
    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID, Nombre", "Cat_Clientes ORDER BY Nombre", Frm_Rpt_Reportes.Cmb_Cliente_Nota_Credito, 0, "Nombre")
End Sub

Private Sub Submenu_Saldos_Proveedor_Click()
    Catalogo = "Reporte de cuentas por pagar"
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Saldos.Caption = "Reporte de cuentas por pagar"
    Frm_Rpt_Reportes.Fra_Saldos.Visible = True
    Frm_Rpt_Reportes.Chk_Imprimir_Saldos_ceros.Visible = True
    Frm_Rpt_Reportes.Lbl_Cliente_Saldo.Caption = "Proveedor"
    ''Frm_Rpt_Reportes.Cmb_Remision_Factura.Visible = False
    ''Frm_Rpt_Reportes.Lbl_Remision_Factura.Visible = False
    Frm_Rpt_Reportes.Cmb_Cliente_Saldo.SetFocus
    Frm_Rpt_Reportes.Dtp_Fecha_Saldo_Inicio.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Saldo_Fin.Value = Now
End Sub

Private Sub Submenu_Salidas_Almacen_Click()
    Unload Frm_Rpt_Reportes
    Load Frm_Rpt_Reportes
    Frm_Rpt_Reportes.Fra_Reporte_Salidas.Visible = True
    Frm_Rpt_Reportes.Fra_Reporte_Salidas.Caption = "Reporte de Salidas de Almacen"
    Frm_Rpt_Reportes.Lbl_Reporte_Pedido_Producto.Caption = "Salidas"
    Frm_Rpt_Reportes.Dtp_Fecha_Inicio_Pedido_Factura.Value = Now
    Frm_Rpt_Reportes.Dtp_Fecha_Fin_Pedido_Factura.Value = Now
End Sub

Private Sub Submenu_Salidas_Click()
    Unload Frm_Alm_Salidas_de_Producto
    Load Frm_Alm_Salidas_de_Producto
    Frm_Alm_Salidas_de_Producto.Caption = "SALIDAS DE ALMACEN"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Salidas", Frm_Alm_Salidas_de_Producto)
End Sub


