Attribute VB_Name = "Module1"

Public Mi_SQL As String                             'Obtiene los valores de la consulta
Public Catalogo As String                           'Indicar que formulario se va a abrir
Public Er As rdoError                               'Especifica que tipo de error se esta cometiendo
Public Conexion_Base As New rdoConnection           'Se utiliza para la conexion a la base de datos
Public Nombre_Usuario As String                     'Obtiene el nombre del usuario
Public Ciclos As Integer                            'Se utiliza para válidar el tiempo de espera de la pantalla de presentación
Public Conectar_Ayudante As Ayudante                'Es utilizada para ligar a la ayuda
Public Par_Fecha As String
Public Base_Datos As String                         'Indica el nombre de la base de datos a conectarse
Public Server As String                             'Indica el nombre del servidor en donde se encuentra la base de datos
Public Rol_ID As String                             'Obtiene el rol que tiene asignado el usuario
Public Usuario_ID As String                         'Alamacena el ID del usuario registrado
Public Intentos_Fallidos As Integer                 'Indica el número de intentos fallidos que puede tener un usuario para deshabilitar la cuenta
Public Bloqueo_Por_No_Utilizar As Integer           'Almacena el parametyro de diferencia de dias
Public Bloqueo_Por_Expiración_Password As Integer   'Almacena el parametyro de diferencia de dias
Public Tipo_Validacion As String                    'Almacena el tipo de utilidad que tenda la ventana de loguin
Public PG_Retencion_IVA As String                   '#  Almacena la Retencion del IVA
Public PG_Retencion_ISR As String                   '#  Almacena la Retencion del ISR
Public PG_Retencion_Flete As String                 '#  Almacena la Retencion de Fletes
Public PG_Impuesto_Cedular As String                '#  Almacena el Impuesto Cedular
Public Porcentaje_IVA As String                     '#  Almacena el Impuesto IVA

''*********************************************VARIABLES PARA LA FACTURA ELECTRONICA*************************************
Public Nombre_Emisor As String
Public Calle_Emisor As String
Public No_Exterior_Emisor As String
Public No_Interior_Emisor As String
Public Colonia_Emisor As String
Public Codigo_Postal_Emisor As String
Public Localidad_Emisor As String
Public Municipio_Emisor As String
Public Estado_Emisor As String
Public RFC_Emisor As String
Public Pais_Emisor As String
Public Regimen_Emisor As String
Public Expedida_En As String
Public Ruta_Certificado As String                   'Almacena la ruta del certificado
Public Ruta_Llave_Privada As String                 'Almacena la ruta de la llave privada
Public Password_Llave As String                     'Almacena el password de la llave privada
Public Ruta_Pdfs As String                          'Almacena la ruta de los pdfs
Public Ruta_Xmls As String                          'Almacena la ruta de los xmls
Public Ruta_NC As String                            'Almacena la ruta de las notas de crédito
Public Ruta_Remisiones As String                    'Almacena la ruta de las remisiones
Public Tasa_Interes_Pagare As Double                'Almacena la tasa del pagare
Public Texto_Factura As String                      'Almacena el texto a imprimir en la factura
Public Continuar As Boolean                         'Bandera que se habilita si los parametros necesrios para generar una factura estan registrados
Public Timbrado_Version As String                   'Almacena la versión del timbrado
Public Timbrado_Codigo_Usuario As String            'Almacena el codigo de usuario para el timbrado
Public Timbrado_Codigo_Usuario_Proveedor As String  'Almacena el código de usuario ante el proveedor
Public Timbrado_ID_Sucursal As String               'Almacena el ID de la sucursal para el timbrado
Public Timbrado_VersionSat As String                 'Almacena la version del timbrado
Public Timbrado_UUID As String                      'Almacena el codigo del timbrado
Public Timbrado_FechaTimbrado As String             'Almacena la fecha del timbrado
Public Timbrado_selloCFD As String                  'Almacena el sello cfd del timbrado
Public Timbrado_noCertificadoSAT As String          'Almacena el certificado sat del timbrado
Public Timbrado_selloSAT As String                  'Almacena el sello sat del timbrado
Public Timbrado_Ambiente As String                  'Almacena el ambiente en que se trabajará el timbrado(Pruebas o Produccion)
Public Folios_Terminados As Boolean                 'Bandera que se habilita cuando ya no hay folios electrónicos disponibles
Public Nombre_BD As String
Public Nombre_Servidor As String
Public Usuario_BD As String
Public Password_BD As String
Public Timbrado_RFCTimbra As String

'Obtiene el directorio ya sea windows o winnt
Public Declare Function GetWindowsDirectory Lib "kernel32" _
      Alias "GetWindowsDirectoryA" ( _
      ByVal lpBuffer As String, _
      ByVal nSize As Long) As Long

'Api para saber si una carpeta existe
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'===========Funciones utilizadas en el proceso de generacion del archivo .Pak=================
Public Declare Function OpenProcess Lib "kernel32.dll" _
   (ByVal dwDA As Long, ByVal bIH As Integer, ByVal dwPID As Long) As Long

Public Declare Sub WaitForSingleObject Lib "kernel32.dll" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long)
   
Public Declare Sub CloseHandle Lib "kernel32.dll" (ByVal hObject As Long)
'=============================================================================================

'===========Para desplegar lista de combo automaticamente=================
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'=============================================================================================

'Funciones para generar el codigo bidimensional
Public Declare Sub GenerateFile Lib "BarCodeLibrary.dll" (ByVal text As String, ByVal fileName As String)

'Abrir archivos
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3

'Manejo de archivos
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

'Funciones para esperar a que termine un proceso de otro programa
Public Const INFINITE = &HFFFF
Public Const SYNCHRONIZE = &H100000

'Variables para manejo de la ventana de seleccion de directorio
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

' Funcion Api CoTaskMemFree
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
' Funcion Api SHBrowseForFolder
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
' Funcion Api SHGetPathFromIDList
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long

' Constantes
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260 ' Para Buffer de caracteres del path
' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
'Const SWP_NOZORDER = &H4
'Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
'Const SWP_DRAWFRAME = &H20
Public Const SWP_SHOWWINDOW = &H40
'Const SWP_HIDEWINDOW = &H80
'Const SWP_NOCOPYBITS = &H100
'Const SWP_NOREPOSITION = &H200
Public Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

'Crear ODBC
Private Const ODBC_ADD_DSN = 1
Private Const ODBC_CONFIG_SYS_DSN = 5
Private Const vbAPINull As Long = 0&

Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hwndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) _
    As Long

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Parametros
'DESCRIPCIÓN: Consulta los parámetros que tiene el sistema asignados
'PARÁMETROS :
'CREO       : Yazmin Delgado Gómez
'FECHA_CREO : 15-Octubre-2007
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Parametros()
Dim Rs_Consulta_Cat_Parametros As rdoResultset 'Consulta los parámetros que tiene asignado el sistema
PG_Retencion_IVA = ""

    'Consulta los parámetros del sistema
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Parametros.EOF Then
        With Rs_Consulta_Cat_Parametros
            Intentos_Fallidos = .rdoColumns("Intentos_Fallidos")
            Bloqueo_Por_No_Utilizar = Val(.rdoColumns("Vencimiento_Cuenta_Usuario"))
            Bloqueo_Por_Expiración_Password = Val(.rdoColumns("Limite_Cambio_Password"))
            If Not IsNull(.rdoColumns("Retencion_IVA")) Then PG_Retencion_IVA = Val(.rdoColumns("Retencion_IVA"))
            If Not IsNull(.rdoColumns("Retencion_ISR")) Then PG_Retencion_ISR = Val(.rdoColumns("Retencion_ISR"))
            If Not IsNull(.rdoColumns("Retencion_Fletes")) Then PG_Retencion_Flete = Val(.rdoColumns("Retencion_Fletes"))
            If Not IsNull(.rdoColumns("Impuesto_Cedular")) Then PG_Impuesto_Cedular = Val(.rdoColumns("Impuesto_Cedular"))
            If Not IsNull(.rdoColumns("Impuesto_IVA")) Then Porcentaje_IVA = Val(.rdoColumns("Impuesto_IVA"))
        End With
    End If
    Rs_Consulta_Cat_Parametros.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Parametros_Facturación
'DESCRIPCIÓN: Consulta los parámetros que tiene el sistema asignados para la facturación electrónica
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 14-Agosto-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Parametros_Facturacion()
Dim Rs_Consulta_Cat_Parametros As rdoResultset 'Consulta los parámetros que tiene asignado el sistema

    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Cat_Parametros
            If Not .EOF Then
                If Not IsNull(.rdoColumns("Ruta_Certificado")) Then Ruta_Certificado = .rdoColumns("Ruta_Certificado")
                If Not IsNull(.rdoColumns("Ruta_Llave_Privada")) Then Ruta_Llave_Privada = .rdoColumns("Ruta_Llave_Privada")
                If Not IsNull(.rdoColumns("Password_Llave_Privada")) Then Password_Llave = Trim(.rdoColumns("Password_Llave_Privada"))
                If Not IsNull(.rdoColumns("Ruta_Pdfs")) Then Ruta_Pdfs = .rdoColumns("Ruta_Pdfs")
                If Not IsNull(.rdoColumns("Ruta_Xmls")) Then Ruta_Xmls = .rdoColumns("Ruta_Xmls")
                If Not IsNull(.rdoColumns("Ruta_NC")) Then Ruta_NC = .rdoColumns("Ruta_NC")
                Ruta_Remisiones = App.Path & "\Remisiones"
                If Not IsNull(.rdoColumns("Porcentaje_Interes_Pagare")) Then Tasa_Interes_Pagare = .rdoColumns("Porcentaje_Interes_Pagare")
                If Not IsNull(.rdoColumns("Mensaje_Factura")) Then Texto_Factura = .rdoColumns("Mensaje_Factura")
                If Not IsNull(.rdoColumns("Version_Timbrado")) Then Timbrado_Version = .rdoColumns("Version_Timbrado")
                If Not IsNull(.rdoColumns("Codigo_Usuario")) Then Timbrado_Codigo_Usuario = .rdoColumns("Codigo_Usuario")
                If Not IsNull(.rdoColumns("Codigo_Usuario_Proveedor")) Then Timbrado_Codigo_Usuario_Proveedor = .rdoColumns("Codigo_Usuario_Proveedor")
                If Not IsNull(.rdoColumns("ID_Sucursal_Timbrado")) Then Timbrado_ID_Sucursal = .rdoColumns("ID_Sucursal_Timbrado")
                If Not IsNull(.rdoColumns("Ambiente_Timbrado")) Then Timbrado_Ambiente = .rdoColumns("Ambiente_Timbrado")
            End If
        End With
    Rs_Cat_Parametros.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Alinea_Derecha
'DESCRIPCIÓN: Alinea_Derecha los número a la izquierda del documento
'PARÁMETROS:
'             1. Numero:
'             2. Longitud:
'CREO: Joel G. Romero Cervantes
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Alinea_Derecha(Numero As String, Longitud As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim Caracteres_Ciclo As Integer     'Cuenta el numero de caracteres de la cadena

    Nuevo = Numero
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For Caracteres_Ciclo = 1 To Longitud - Len(Numero)
        Nuevo = " " & Nuevo
    Next Caracteres_Ciclo
    Alinea_Derecha_Derecha = Nuevo
End Function

Public Sub Mover_Control_Grid_TextBox(Flex_Grid As MSFlexGrid, Control_Mover As TextBox)      'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = Flex_Grid.Left + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            .Height = Flex_Grid.CellHeight
            .text = Flex_Grid
            .Visible = True
            .SetFocus
            'SendKeys "{Home}+{End}"
        End If
    End With
End Sub

Function Imprime_Varias_Lineas(Real As String, Tamaño As Integer, Contador_Renglon, Salto_Linea) As Double
Dim Ultima_Posicion As Integer
Dim Aux_Espacio As Integer
Dim Espacio As Integer
Dim Cadena As String
Dim Cortada As String

    Ultima_Posicion = 1
    Espacio = 1
    Aux_Espacio = 1
    Real = Real & Chr(13)
    Cadena = Mid(Real, Ultima_Posicion, Tamaño)
    While Cadena <> ""
      ' Debug.Print Cadena
        Espacio = 0
        Aux_Espacio = 1
        While Aux_Espacio > 0
            Espacio = Aux_Espacio
            Aux_Espacio = InStr(Espacio + 1, Cadena, Chr(13), vbTextCompare)
            If Aux_Espacio = 0 Then
                Aux_Espacio = InStr(Espacio + 1, Cadena, " ", vbTextCompare)
            Else
                Espacio = Aux_Espacio + 1
                Aux_Espacio = 0
                Cadena = Mid(Cadena, 1, Espacio - 2)
            End If
        Wend
        If Espacio > 0 Then
            Printer.Print Spc(3); Mid(Cadena, 1, Espacio)
            Contador_Renglon = Contador_Renglon + Salto_Linea
        End If
        Ultima_Posicion = Ultima_Posicion + Espacio
        Cadena = Mid(Real, Ultima_Posicion, Tamaño)
    Wend
    Imprime_Varias_Lineas = Contador_Renglon
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Selecciona_Ruta_Directorio
'DESCRIPCIÓN: Asigna la ruta seleccionada a nivel directorio
'PARÁMETROS : Frm, pasa la forma requerida
'             Caption_Asignado, asigna el titulo a la ventana
'CREO       : Ismael Prieto Sánchez
'FECHA_CREO : 29/Ago/2009 10:05am
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Selecciona_Ruta_Directorio(Frm As Form, Caption_Asignado As String) As String
Dim Nulo As Integer
Dim Identificador As Long
Dim Ruta As String
Dim Directorios As BrowseInfo

    With Directorios
        .hWndOwner = Frm.hwnd    'Formulario
        .lpszTitle = lstrcat(Caption_Asignado, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Mostrar el cuadro de dialogo Buscar carpeta
    Identificador = SHBrowseForFolder(Directorios)
    If Identificador Then
        Ruta = String$(MAX_PATH, 0)
        'Llamamos a la APi y recuperamos el id del path y en _
         El_Path obtenemos el path seleccionado
        Call SHGetPathFromIDList(Identificador, Ruta)
        'Liberamos el bloque de memoria
        Call CoTaskMemFree(Identificador)
        'Busca la posición del primer caracter nulo
        Nulo = InStr(Ruta, vbNullChar)
        If Nulo Then
            'Formateamos la cadena anterior eliminado los espacios nulos del path
            Ruta = Left$(Ruta, Nulo - 1)
        End If
        Selecciona_Ruta_Directorio = Ruta
    End If
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Asignación_Datos_Cliente
'DESCRIPCIÓN: Se asignan la información a las variables globales con los datos de la empresa
'             sin el uso del archivo config.ini
'PARÁMETROS:
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  26-Agosto-2010
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Asignación_Datos_Cliente()
    'Asignacion de las variables globales con los datos de la empresa cliente
    Nombre_Emisor = "ALCOHOLERA DEL CENTRO S.A. DE C.V."
    'RFC_Emisor = "ACE781116T52"
    RFC_Emisor = "AAA010101AAA"
    Calle_Emisor = "CARRETERA IRAPUATO-SILAO"
    No_Exterior_Emisor = "KM 13.5"
    No_Interior_Emisor = ""
    Colonia_Emisor = "TARETAN"
    Codigo_Postal_Emisor = "36815"
    Municipio_Emisor = "IRAPUATO"
    Estado_Emisor = "GUANAJUATO"
    Pais_Emisor = "MEXICO"
    Expedida_En = "IRAPUATO,GUANAJUATO"
    Regimen_Emisor = "601 REGIMEN GENERAL DE LEY PERSONAS MORALES"
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Aviso_Termino_Folios
'DESCRIPCIÓN: Consulta si aun hay folios disponible
'PARÁMETROS : Tipo_Folio - indica si el folio si es FACTURA O NOTA CREDITO
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 14-Agosto-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Aviso_Termino_Folios(Tipo_Folio As String)
Dim Rs_Consulta_Parametro_Folio As rdoResultset 'Variable para el manejo de la tabla
Dim Rs_Consulta_Parametros As rdoResultset      'Variable para el manejo de la tabla
Dim Rs_Consulta_Factura_Folio As rdoResultset   'Variable para el manejo de la tabla
Dim Termino_Folios As Double                    'Almacena el parametro de los folios
Dim Folio_Final As Double                       'Almacena el folio final activo
Dim Folio_Factura As Double                     'Almacena el folio de la factura

On Error GoTo errorHandler
    
    MDIFrm_Apl_Principal.MousePointer = 11
    Folios_Terminados = False
    'Realiza la consulta del parametro
    Mi_SQL = "SELECT Dias_Aviso_Termina_Folios FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametros
            If Not .EOF Then
                If Not IsNull(.rdoColumns("Dias_Aviso_Termina_Folios")) Then
                    Termino_Folios = Val(.rdoColumns("Dias_Aviso_Termina_Folios"))
                Else
                    Termino_Folios = 0
                End If
            Else
                Termino_Folios = 0
            End If
        End With
    Rs_Consulta_Parametros.Close
    
    If Termino_Folios > 0 Then
        If Tipo_Folio = "FACTURA" Then
            'Consulta los parametros del folio final
            Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
            Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
            Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
            Set Rs_Consulta_Parametro_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Parametro_Folio
                    If Not .EOF Then
                        Folio_Final = Val(.rdoColumns("Folio_Final"))
                    Else
                        Folio_Final = 0
                    End If
                End With
            Rs_Consulta_Parametro_Folio.Close
        Else
            If Tipo_Folio = "NOTA CARGO" Then
                'Consulta los parametros del folio final
                Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
                Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
                Mi_SQL = Mi_SQL & " AND Tipo = 'FACTURA'"
                Mi_SQL = Mi_SQL & " AND Serie = 'NCA'"
                Set Rs_Consulta_Parametro_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    With Rs_Consulta_Parametro_Folio
                        If Not .EOF Then
                            Folio_Final = Val(.rdoColumns("Folio_Final"))
                        Else
                            Folio_Final = 0
                        End If
                    End With
                Rs_Consulta_Parametro_Folio.Close
                
            ElseIf Tipo_Folio = "PAGOS" Then
            'Consulta los parametros del folio final
            Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
            Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
            Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
            Set Rs_Consulta_Parametro_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Parametro_Folio
                    If Not .EOF Then
                        Folio_Final = Val(.rdoColumns("Folio_Final"))
                    Else
                        Folio_Final = 0
                    End If
                End With
            Rs_Consulta_Parametro_Folio.Close
            
            Else
                If Tipo_Folio = "NOTA_CREDITO" Then
                    'Consulta los parametros del folio final
                    Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
                    Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
                    Mi_SQL = Mi_SQL & " AND Tipo = '" & Tipo_Folio & "'"
                    Set Rs_Consulta_Parametro_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        With Rs_Consulta_Parametro_Folio
                            If Not .EOF Then
                                Folio_Final = Val(.rdoColumns("Folio_Final"))
                            Else
                                Folio_Final = 0
                            End If
                        End With
                    Rs_Consulta_Parametro_Folio.Close
                End If
        End If
    End If
        
        'Valida el folio final electronico
        If Tipo_Folio = "FACTURA" Then
            'Consulta el folio final de la factura
'            Mi_SQL = "SELECT ISNULL(MAX(No_Factura),0) AS No_Factura FROM Ope_Facturas_Electronicas"
            Mi_SQL = "SELECT MAX(No_Factura_Electronica) AS Factura FROM Adm_Clientes_Facturas WHERE Forma_Factura = 'E'"
            Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Factura_Folio
                    If Not .EOF Then
                        If Not IsNull(.rdoColumns("Factura")) Then
                            Folio_Factura = Val(.rdoColumns("Factura"))
                        Else
                            Folio_Factura = 0
                        End If
                    Else
                        Folio_Factura = 0
                    End If
                End With
            Rs_Consulta_Factura_Folio.Close
        Else
            If Tipo_Folio = "NOTA CARGO" Then
            'Consulta el folio final de la nota de cargo
            Mi_SQL = "SELECT MAX(No_Factura_Electronica) AS Factura FROM Adm_Clientes_Facturas WHERE Forma_Factura = 'E' AND  Serie = 'NCA'"
            
            Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Factura_Folio
                    If Not .EOF Then
                        If Not IsNull(.rdoColumns("Factura")) Then
                            Folio_Factura = Val(.rdoColumns("Factura"))
                        Else
                            Folio_Factura = 0
                        End If
                    Else
                        Folio_Factura = 0
                    End If
                End With
            Rs_Consulta_Factura_Folio.Close
            
            ElseIf Tipo_Folio = "PAGOS" Then
            'Consulta el folio final de la nota
'            Mi_SQL = "SELECT ISNULL(MAX(No_Nota_Credito),0) AS No_Nota_Credito FROM Ope_Notas_Credito_Clientes"
            Mi_SQL = "SELECT MAX(No_Factura) AS Pago FROM Complemento_Pago"
            Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Factura_Folio
                    If Not .EOF Then
                        If Not IsNull(.rdoColumns("Pago")) Then
                            Folio_Factura = Val(.rdoColumns("Pago"))
                        Else
                            Folio_Factura = 0
                        End If
                    Else
                        Folio_Factura = 0
                    End If
                End With
            Rs_Consulta_Factura_Folio.Close
            
            Else
            'Consulta el folio final de la nota
'            Mi_SQL = "SELECT ISNULL(MAX(No_Nota_Credito),0) AS No_Nota_Credito FROM Ope_Notas_Credito_Clientes"
            Mi_SQL = "SELECT MAX(No_Nota_Credito) AS Nota_Credito FROM Adm_Notas_Credito"
            Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Factura_Folio
                    If Not .EOF Then
                        If Not IsNull(.rdoColumns("Nota_Credito")) Then
                            Folio_Factura = Val(.rdoColumns("Nota_Credito"))
                        Else
                            Folio_Factura = 0
                        End If
                    Else
                        Folio_Factura = 0
                    End If
                End With
            Rs_Consulta_Factura_Folio.Close
            End If
            
        End If
        
        'Realiza la validacion de los folios
        If (Folio_Final = Folio_Factura) Or (Folio_Final = 0) Then
            'Habilita la bandera para indicar que ya no hay folios disponibles para utilizar
            Folios_Terminados = True
        Else
            If (Folio_Final - Folio_Factura) <= Termino_Folios Then
                MDIFrm_Apl_Principal.MousePointer = 0
                If Tipo_Folio = "FACTURA" Then
                    MsgBox "Quedan " & Folio_Final - Folio_Factura & " folios disponibles para las facturas electrónicas", vbExclamation
                Else
                    If Tipo_Folio = "NOTA CARGO" Then
                        MsgBox "Quedan " & Folio_Final - Folio_Factura & " folios disponibles para las notas de cargo electrónicas", vbExclamation
                    ElseIf Tipo_Folio = "PAGOS" Then
                        MsgBox "Quedan " & Folio_Final - Folio_Factura & " folios disponibles para los pagos", vbExclamation
                    Else
                        MsgBox "Quedan " & Folio_Final - Folio_Factura & " folios disponibles para las notas de crédito electrónicas", vbExclamation
                    End If
                End If
            End If
        End If
    Else
        MDIFrm_Apl_Principal.MousePointer = 0
        If Tipo_Folio = "FACTURA" Then
            MsgBox "No hay folios disponibles para poder generar las facturas electrónicas", vbCritical
        Else
            If Tipo_Folio = "NOTA CARGO" Then
                MsgBox "No hay folios para poder generar las notas de cargo electrónicas", vbCritical
            ElseIf Tipo_Folio = "PAGOS" Then
                MsgBox "No hay folios para poder generar los pagos", vbCritical
            Else
                MsgBox "No hay folios para poder generar las notas de crédito electrónicas", vbCritical
            End If
        End If
        'Habilita la bandera para indicar que ya no hay folios disponibles para utilizar
        Folios_Terminados = True
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN:  Enviar_Correo_Documentos
'DESCRIPCIÓN:           Envia el correo con los parametros establecido
'PARÁMETROS :           From_Email: correo de quien envia
'                       Nombre_From: Nombre quien envia
'                       To_Email:correo a quien se envia
'                       Nombre_To: nombre a quien se envia
'                       Asunto: asunto del correo
'                       Mensaje_Email: mensaje del correo
'CREO       :           Sergio Godínez Banda
'FECHA_CREO :           19 Mayo 2009
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Enviar_Correo_Documentos(From_Email As String, Nombre_From As String, To_Email As String, Asunto As String, Mensaje_Email As String, Mostrar_Mensaje As Boolean, Adjunto As Boolean, Optional Ruta_Adjunto As String, Optional CC As String) As Boolean
Dim Mensaje_Outlook As Object
Dim Correo As Object
On Error GoTo handler
Enviar_Correo = False
    Dim colAttach As Object
    Dim l_Attach As Object
    Dim l_Msg As Object
    Dim oSession As Object
    Dim oMsg As Object
    Dim oAttachs As Object
    Dim oAttach As Object
    Dim colFields As Object
    Dim oField As Object
    If Mensaje_Outlook Is Nothing Then
        Set Mensaje_Outlook = CreateObject("Outlook.Application")
    End If
    Set Correo = Mensaje_Outlook.CreateItem(0)
    
    Dim version_windows As String
    Dim SigString As String
    Dim Firma As String
    version_windows = GetVersion
   
    Dim strEntryID As String
    Dim ImgString As String

    With Correo
        .To = To_Email
        If CC <> "" Then
            .CC = CC
        End If
        .Subject = Asunto
        .BodyFormat = 2
    'uBICA LA RUTA DE LA FIRMA
''        If version_windows = "6" Or Trim(version_windows) = "" Then
''       '        WINDOWS XP
''            SigString = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Firmas\WAP Institucional.htm"
''           'SigString = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Firmas\WAP Institucional.htm"
''           ImgString = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Firmas\WAP Institucional_archivos\image001.jpg"
''        Else
''       'WINDOWS 7
''            SigString = "C:\Documents and Settings\" & Environ("username") & _
''                "\Datos de programa\Microsoft\Firmas\WAP Institucional.htm"
''                'C:\Documents and Settings\diego\Datos de programa\Microsoft\Signatures
''            ImgString = "C:\Documents and Settings\Diego\Datos de programa\Microsoft\Signatures\WAP Institucional_archivos\image001.jpg"
''        End If
''        'SigString = "C:\Documents and Settings\Diego\Datos de programa\Microsoft\Firmas\WAP Institucional.htm"
''        'Agrega la imagen del la firma
''        If Dir(SigString) <> "" Then
''            Firma = GetBoiler(SigString)
''        Else
''            Firma = ""
''        End If
''        'ImgString = "C:\Documents and Settings\Diego\Datos de programa\Microsoft\Firmas\WAP Institucional_archivos\image001.jpg"
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(ImgString, "ARCHIVO") = True Then
            Set colAttach = Correo.Attachments
            Set l_Attach = colAttach.Add(ImgString)
            Correo.Close olSave
            strEntryID = Correo.EntryID
            Set Correo = Nothing
            Set colAttach = Nothing
            Set l_Attach = Nothing
            
            On Error Resume Next
            ' initialize CDO session
            Set oSession = CreateObject("MAPI.Session")
            oSession.Logon "", "", False, False
            ' get the message created earlier
            Set oMsg = oSession.GetMessage(strEntryID)
            ' set properties of the attached graphic that make
            ' it embedded and give it an ID for use in an <IMG> tag
            Set oAttachs = oMsg.Attachments
            Set oAttach = oAttachs.Item(1)
            Set colFields = oAttach.Fields
            Set oField = colFields.Add(CdoPR_ATTACH_MIME_TAG, "image/jpeg")
            Set oField = colFields.Add(&H3712001E, "wap_sign")
            oMsg.Fields.Add "{0820060000000000C000000000000046}0x8514", 11, True
            oMsg.Update
            Firma = Replace(Firma, "WAP%20Institucional_archivos/image002.jpg", "cid:image001.jpg")
            Firma = Replace(Firma, "WAP%20Institucional_archivos/image001.jpg", "cid:image001.jpg")
            ' get the Outlook MailItem again
            Set Correo = Mensaje_Outlook.GetNamespace("MAPI").GetItemFromID(strEntryID)
        End If
        '.Body = Mensaje_Email
        .HTMLBody = Mensaje_Email & "<br><br>" & Firma
        If Adjunto = True Then
            Dim Cont_Fila As Integer
            Dim Archivos_Adjuntos() As String
            Archivos_Adjuntos = Split(Ruta_Adjunto, "|")
            For Cont_Fila = 0 To UBound(Archivos_Adjuntos)
                If Trim(Archivos_Adjuntos(Cont_Fila)) <> "" Then
                    'valida que exista el archivo para adjuntarlo
                    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Archivos_Adjuntos(Cont_Fila), "ARCHIVO") = True Then
                        .Attachments.Add Trim(Archivos_Adjuntos(Cont_Fila))
                    End If
                End If
            Next
        End If
        Correo.Close (olSave)
        .send
    End With
    Set oField = Nothing
    Set colFields = Nothing
    Set oMsg = Nothing
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(ImgString, "ARCHIVO") = True Then
        oSession.Logoff
    End If
    Set oSession = Nothing
    Set Mensaje_Outlook = Nothing
    Set l_Msg = Nothing
    Set Correo = Nothing
    Set Mensaje_Outlook = Nothing
    Enviar_Correo_Documentos = True
Exit Function

handler:
    Enviar_Correo = False
    If Err.Number = 429 Then
        MsgBox "Error conectando a outlook"
    End If
End Function

'*************************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Fechas
'DESCRIPCIÓN: Valida que las fechas que proporciono el usuario sean validas para el
'             sistema mandando un estatus de verdadero pero si no son validas
'             entonces manda un valor de falso
'PARÁMETROS :
'CREO       : Yazmin A. Delgado Gómez
'FECHA_CREO : 13-Diciembre-2007
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*************************************************************************************
Function Valida_Fechas(Fecha_Inicio As Date, Fecha_Final As Date) As Boolean
    If Year(Format(Fecha_Inicio, "yyyy/MM/dd")) < 1900 Or Year(Format(Fecha_Final, "yyyy/MM/dd")) > Year(Format(Now, "yyyy/MM/dd")) Then
        Valida_Fechas = False
    Else
        Valida_Fechas = True
    End If
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Crear_ODBC
'DESCRIPCIÓN: Se crea en tiempo de ejecución el ODBC requerido para la generación de los PDF mediante
'             los reportes de crystal
'PARÁMETROS:  Devuelve TRUE si el ODBC fue creado, de lo contrario devuelve FALSE
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  12-Marzo-2013
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Crear_ODBC() As Boolean
Dim DName As String
Dim DSName As String
Dim DServer As String
Dim DDatabase As String
Dim DDesc As String
Dim RetVar As Long

    'Asigna los parametros
    DName = "SQL Server" & Chr$(0)
    DSName = "DSN=Alcoholera" & Chr$(0)
    DServer = "Server=" & Nombre_Servidor & Chr$(0)
    DDatabase = "Database=" & Nombre_BD & Chr$(0)
    DDesc = "Description=ODBC de facturación electrónica" & Chr$(0)

    'Ejecuta la configuracion
    RetVar = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, DName, DSName & DServer & DDatabase & DDesc)

    If RetVar = 1 Then
        Crear_ODBC = True
    Else
        Crear_ODBC = False
    End If
End Function

Public Function Limpia_Variables()
    CFD_Generales.Fecha = ""
    CFD_Generales.Fecha_Timbrado = ""
    CFD_Generales.Folio = ""
    CFD_Generales.Forma_Pago = ""
    CFD_Generales.Metodo_Pago = ""
'    CFD_Generales.No_Cuenta_Pago = ""
    CFD_Generales.Condiciones_Pago = ""
    CFD_Generales.Descuento = 0
    CFD_Generales.SubTotal = 0
    CFD_Generales.Total = 0
    CFD_Generales.Impuestos = 0
'    CFD_Generales.Metodo_Pago_Completo = ""
    CFD_Generales.Tipo_Factura = ""
    CFD_Emisor.RFC = ""
    CFD_Emisor.cp = ""
    CFD_Emisor.Regimen_Fiscal = ""
    CFD_Relacionados.Existe = False
End Function

Public Function Consulta_Cancelados()
    Dim Consultar As rdoResultset
    Dim Modificar As rdoResultset
    Dim Rs_MiComplemento As rdoResultset
    Dim Mensaje As String
    Dim Modificar2 As rdoResultset
    Dim Factura() As String
    On Error GoTo errorHandler
    MDIFrm_Apl_Principal.MousePointer = 11
    Conexion_Base.BeginTrans
    'Consulta facturas en proceso de cancelación
    If Timbrado_Ambiente = "" Then Consulta_Parametros_Facturacion
    Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas WHERE Cancelada='PC'"
    Set Consultar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Consultar
        While Not .EOF
            Mensaje = CFD_Cancela_Xml(.rdoColumns("Timbre_UUID"), True)
            Mi_SQL = "SELECT Cancelada,Mensaje_Cancelado FROM Adm_Clientes_Facturas WHERE Timbre_UUID='" & .rdoColumns("Timbre_UUID") & "'"
            Set Modificar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            With Modificar
                .Edit
                    .rdoColumns("Mensaje_Cancelado") = Mensaje
                    If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                        .rdoColumns("Cancelada") = "S"
                    End If
                    If Mensaje Like "*Rechazada*" Then
                        .rdoColumns("Cancelada") = "N"
                    End If
                .Update
            End With
            Modificar.Close
            .MoveNext
        Wend
    End With
    Consultar.Close
    'Consulta notas de crédito en proceso de cancelación
    Mi_SQL = "SELECT * FROM Adm_Notas_Credito WHERE Cancelada='PC'"
    Set Consultar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Consultar
        While Not .EOF
                Mensaje = CFD_Cancela_Xml(.rdoColumns("Timbre_UUID"), True)
                Mi_SQL = "SELECT Cancelada,Mensaje_Cancelado FROM Adm_Notas_Credito WHERE Timbre_UUID='" & .rdoColumns("Timbre_UUID") & "'"
                Set Modificar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                With Modificar
                    .Edit
                        .rdoColumns("Mensaje_Cancelado") = Mensaje
                        If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                            .rdoColumns("Cancelada") = "S"
                        End If
                        If Mensaje Like "*Rechazada*" Then
                            .rdoColumns("Cancelada") = "N"
                        End If
                    .Update
                End With
                Modificar.Close
                If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                    Factura = Split(.rdoColumns("Factura_Referencia"), " ")
                    Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas WHERE No_Factura_Electronica='" & Format(Factura(1), "0000000000") & "' AND Serie='" & Trim(Factura(0)) & "'"
                    Set Modificar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    Modificar.Edit
                        Modificar.rdoColumns("Saldo") = Modificar.rdoColumns("Saldo") + Val(.rdoColumns("Total"))
                        Modificar.rdoColumns("Pagada") = "N"
                    Modificar.Update
                    Modificar.Close
                End If
            .MoveNext
        Wend
    End With
    Consultar.Close
    'Consulta pagos en proceso de cancelación
    Mi_SQL = "SELECT * FROM Adm_Movimientos WHERE Estatus='P'"
    Set Consultar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    With Consultar
        While Not .EOF
                .Edit
                    Mi_SQL = "SELECT Timbre_UUID FROM Complemento_Pago WHERE No_Factura='" & .rdoColumns("No_Complemento_Pago") & "'"
                    Set Modificar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    With Modificar
                            Mensaje = CFD_Cancela_Xml(.rdoColumns("Timbre_UUID"), True)
                    End With
                    Modificar.Close
                
'                    Mi_SQL = "SELECT Estatus,Mensaje_Cancelado FROM Adm_Movimientos WHERE No_Complemento_Pago='" & .rdoColumns("Timbre_UUID") & "'"
'                    Set Modificar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
'                    With Modificar
'                        .Edit
                    .rdoColumns("Mensaje_Cancelado") = Mensaje
                    If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                        Mi_SQL = "SELECT *"
                        Mi_SQL = Mi_SQL & " FROM Complemento_Pago"
                        Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & .rdoColumns("No_Complemento_Pago") & "'"
                        Set Rs_MiComplemento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        If Not Rs_MiComplemento.EOF Then
                            Rs_MiComplemento.Edit
                                Rs_MiComplemento.rdoColumns("Estatus") = "CANCELADA"
                            Rs_MiComplemento.Update
                        End If
                        Rs_MiComplemento.Close
                        .rdoColumns("Estatus") = "C"
                    End If
                    If Mensaje Like "*Rechazada*" Then
                        .rdoColumns("Estatus") = "A"
                            End If
'                        .Update
'                    End With
'                    Modificar.Close
                If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
'                    Mi_SQL = "SELECT * FROM Ope_Complemento_Pago_Detalles WHERE No_Pago='" & .rdoColumns("No_Pago") & "' AND Serie='" & Trim(.rdoColumns("Serie")) & "'"
'                    Set Modificar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'                        While Not Modificar.EOF
                            Mi_SQL = "SELECT * FROM Adm_Clientes_Facturas WHERE No_Factura='" & .rdoColumns("No_Factura") & "'"
                            Set Modificar2 = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                            Modificar2.Edit
                                Modificar2.rdoColumns("Saldo") = Modificar2.rdoColumns("Saldo") + Val(.rdoColumns("Cantidad"))
                                Modificar2.rdoColumns("Pagada") = "N"
                                Modificar2.rdoColumns("No_Parcialidad") = Modificar2.rdoColumns("No_Parcialidad") - 1
                            Modificar2.Update
                            Modificar2.Close
'                            Modificar.MoveNext
'                        Wend
'                    Modificar.Close
                End If
                .Update
            .MoveNext
        Wend
    End With
    Consultar.Close
    Conexion_Base.CommitTrans
    MDIFrm_Apl_Principal.MousePointer = 0
'    MsgBox "Actualización Exitosa"
    Exit Function
errorHandler:
    Conexion_Base.RollbackTrans
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox Err.Description, vbExclamation, “Error”
End Function
