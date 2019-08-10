VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ayudante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Option Compare Text
Private Mensage_Error As String          'Propiedad de mensaje de error, Vacía si no existe error
Private Redim_Forma As Form                    'Forma a la que se va a redimencionar
Private Cont_Controles_No_Redim As Integer 'Cuenta la cantidad de controles que no se pueden redimencionar
Private Alto_Original As Long              'Guarda el Alto original
Private Ancho_Original As Long               'Guarda el Ancho original
Private Escala_Alto_Original As Long         'Guarda el Scaleheight original

Private Type Info_Control                       'Tipo de Dato definido por el usuario
    Numero_Control As Integer                      'Numero de Control
    Borde_Superior As Long                              'Manejo del propiedad Top de control
    Borde_Izquierdo As Long                              'Manejo de la propiedad Left del control
    Alto As Long                              'Manejo de la propiedad Height del control
    Ancho As Long                              'Manejo de la propiedad Width del control
    f As Long
End Type
Private Arr_Controles() As Info_Control                'Variable de Usuario

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Solo_Caracteres_Rango
    'DESCRIPCIÓN:          Permite caracteres en el rango establecido: rango[A-B,a-b,0-9,(,),Backspace,Tab,Spacebar,Supr]
    '                      Regresa el caracter si esta dentro del rango o caracter 0
    '                      se utiliza en en keypress
    'PARÁMETROS:           KeyAscii:entrada del teclado
    'CREO:                 Yañez Rodriguez Diego Neftali
    'FECHA_CREO:           13-Diciembre-2007
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Solo_Caracteres_Alfa_Numericos(KeyAscii As Integer)
   
   'Regresa el valor en mayuscula para cacteres alfabeticos
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    'Rango de caracteres validos para la funcion
    Select Case (KeyAscii)
        'Numeros
        Case vbKey0 To vbKey9: Exit Function
        'Letras
        Case vbKeyA To vbKeyZ: Exit Function
        'Barra Espaciadora
        Case vbKeySpace: Exit Function
        'Regresa Espacio,
        Case vbKeyBack: Exit Function
        'Parentesis ( )
        Case 40 To 41: Exit Function
        'Supr
        Case vbKeyDelete: Exit Function
        'Tab
        Case vbKeyTab: Exit Function
        'Ñ
        Case 209: Exit Function
        'Tecla Enter
        Case vbKeyReturn: Exit Function
    End Select
    'Fuera de rango
    KeyAscii = 0
End Function
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Correccion_Ortografica
    'DESCRIPCIÓN: Se instancia un objeto document para pasarle el texto contenido
    '             en un control TextBox el cual será analizado por la herramienta ortográfica
    '             de Ms. Word.
    'PARÁMETROS : Txt_Cadena - Control TextBox que contiene el texto a ser revisado.
    'CREO       : Jose Antonio Lopez Hernandez
    'FECHA_CREO : 17/Junio/2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Correccion_Ortografica(Txt_Cadena As TextBox)
'Realiza la correccion ortografica
Dim objWord As New Word.Application
Dim objDocument As New Document

objWord.WindowState = wdWindowStateMinimize
Set objDocument = objWord.Documents.Add()
With objDocument
    .Select
    .Range.text = Txt_Cadena.text
    .CheckSpelling , True, True
    .CheckGrammar
    Txt_Cadena.text = Replace(.Range.text, vbCr, vbCrLf)
    .Saved = True
    .Close
End With
objWord.Quit
Set objDocument = Nothing
Set objWord = Nothing
End Function
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Quitar_Caracter_Raro
    'DESCRIPCIÓN:          Permite caracteres en el rango establecido: rango[A-B,a-b,0-9,(,),Backspace,Tab,Spacebar,Supr]
    '                      Regresa el caracter si esta dentro del rango, ó caracter 0 caso contrario
    '                      se utiliza en el keypress del control
    'PARÁMETROS:           Caracter_KeyAscii: entrada del teclado
    'CREO:                 Yañez Rodriguez Diego Neftali
    'FECHA_CREO:           13-Diciembre-2007
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Quitar_Caracter_Raro(Caracter_KeyAscii As Integer)
    
    'Regresa el valor en mayuscula para cacteres alfabeticos
    Caracter_KeyAscii = Asc(UCase(Chr(Caracter_KeyAscii)))
    
    'Rango de caracteres validos para la funcion
    Select Case (Caracter_KeyAscii)
        'Numeros
        Case vbKey0 To vbKey9: Exit Function
        'Letras
        Case vbKeyA To vbKeyZ: Exit Function
        'Barra Espaciadora
        Case vbKeySpace: Exit Function
        'Regresa Espacio,
        Case vbKeyBack: Exit Function
        'Parentesis ( )
        Case 40 To 41: Exit Function
        'Supr
        Case vbKeyDelete: Exit Function
        'Tab
        Case vbKeyTab: Exit Function
        'Ñ
        Case 209: Exit Function
        'Tecla Enter
        Case vbKeyReturn: Exit Function
    End Select
    
    'Fuera de rango
    Caracter_KeyAscii = 0
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Quitar_Caracter
'DESCRIPCIÓN: Quita los caracteres raros de los textos para poder ser guardados
'             en la base de datos
'PARÁMETROS:
'             1. Cadena1: Texto al que se le van a quitar los caracteres raros
'CREO: Joel G. Romero Cervantes
'FECHA_CREO:6/JUN/05
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Quitar_Caracter(Texto As String, Caracter As String) As String
    Dim posicion As Integer
    posicion = 1
    While posicion <> 0
        posicion = InStr(1, Texto, Caracter, vbTextCompare)
        If posicion <> 0 Then
            Texto = Mid(Texto, 1, posicion - 1) & Mid(Texto, posicion + 1, Len(Texto))
        End If
    Wend
    Quitar_Caracter = Texto
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Solo_Letras
'DESCRIPCIÓN: 'Solo acepta letras, espacios y puntos
'PARÁMETROS:
'             1. Caracter_KeyAscii: Regresa el valor de la letra
'CREO: Yazmin Abigail Delgado Gómez
'FECHA_CREO:
'MODIFICO: Antonio Salvador Benavides Guardado
'FECHA_MODIFICO: 13/junio/2008
'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Solo_Letras(Caracter_KeyAscii As Integer)
    Caracter_KeyAscii = Asc(UCase(Chr(Caracter_KeyAscii)))
    'Acepta las letras de la A a la Z, incluyendo Ñ o ñ, espacios  y Puntos
    If (Chr(Caracter_KeyAscii) < "A" Or Chr(Caracter_KeyAscii) > "Z") Or Caracter_KeyAscii = 46 Then
        If Caracter_KeyAscii = 8 Or Caracter_KeyAscii = 32 Or Caracter_KeyAscii = 209 Then
            Exit Function
        End If
        Caracter_KeyAscii = 0
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Es_Alfanumerico
    'DESCRIPCIÓN: Indica si el texto tiene numeros y letras
    'PARÁMETROS : Texto: Texto a compara para saber si tiene letras y numeros
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 15-Octubre-2007
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Es_Alfanumerico(Texto As String) As Boolean
Dim Temp_Texto As String     'Obtiene el texto que se desea validar
Dim Long_Texto As Integer     'Obtiene la longitud del texto que se esta validando
Dim Pos_Caracter As Integer     'Indica la posición del caracter que se esta consultando dentro de la cadena
Dim Caracter As String     'Obtiene el caracter que se esta comparando del texto que se desea validar
Dim Hay_Numeros As Boolean   'Indica si se encontraron numeros
Dim Hay_Caracteres As Boolean 'Indica si se encontraron letras
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    Hay_Numeros = False
    Hay_Caracteres = False
    Temp_Texto = Texto
    Long_Texto = Len(Temp_Texto)
    If Long_Texto > 0 Then
        For Pos_Caracter = 1 To Long_Texto
            Caracter = Mid(Temp_Texto, Pos_Caracter, 1)
            If Caracter Like "[0-9]" Then Hay_Numeros = True
            If Caracter Like "[A-Za-z]" Then Hay_Caracteres = True
        Next
        If Hay_Caracteres = True And Hay_Numeros = True Then
            Es_Alfanumerico = True
        Else
            Es_Alfanumerico = False
        End If
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Solo_Numeros
    'DESCRIPCIÓN: 'Solo acepta números y el punto
    'PARÁMETROS:
    '             1. Caracter_KeyAscii: Regresa el valor del numero
    '             2. Cadena  : Contiene la cadena del text a validar
    'CREO: Yazmin Abigail Delgado Gómez
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Solo_Numeros(KeyAscii As Integer, Cadena As String, Permitir_Punto As Boolean)
    'Solo acepta numeros, el espacio y un punto decimal
    If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) Or KeyAscii = 46 Then
        'Valida que solo haya un punto en la cadena
        If KeyAscii = 46 Then
            'Si se permite escribir un punto
            If Permitir_Punto = True Then
                'Busca el punto en la cadena (1-Existe/0-No Existe)
                If InStr(1, Cadena, ".", vbBinaryCompare) = 0 Then
                    Exit Function
                End If
            End If
        Else 'Si no es el punto
            If KeyAscii = 8 Then
                Exit Function
            End If
        End If
        
        KeyAscii = 0
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Centrar_Texto
    'DESCRIPCIÓN: Centra el texto en la línea de impresón
    'PARÁMETROS: Texto     : Texto a centrar
    '            Longitud_Editor: Numero de caracteres a centrar
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Centrar_Texto(Texto As String, Longitud_Editor As Integer)
Dim Txt_Sin_Espacios As String
Dim Txt_Longitud, Espacio_Izquierda, I As Integer
    Txt_Sin_Espacios = Trim(Texto)
    Txt_Longitud = Len(Txt_Sin_Espacios)
    Espacio_Izquierda = (Longitud_Editor - Txt_Longitud) / 2
    For I = 1 To Int(Espacio_Izquierda)
        Txt_Sin_Espacios = " " & Txt_Sin_Espacios
    Next I
    Centrar_Texto = Txt_Sin_Espacios
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Alinea_Derecha
'DESCRIPCIÓN: Alinea los número a la derecha del documento
'PARÁMETROS:
'             1. Numero: Es la cadena que deseo alinear
'             2. Longitud: Es la longitud que va a contener mi cadena
'CREO: Joel G. Romero Cervantes
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Alinea_Derecha(Numero As String, Longitud As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim I As Integer     'Cuenta el numero de caracteres de la cadena
    Nuevo = Numero
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For I = 1 To Longitud - Len(Numero)
        Nuevo = " " & Nuevo
    Next I
    Alinea_Derecha = Nuevo
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Formulario_Cargado
'DESCRIPCIÓN: Indica si el formulario esta cargado o no dentro del sistema
'PARÁMETROS : Nombre_Formulario: Nombre del formulario el cual se pretende saber si
'             esta cargado o no en el sistema, si esta cargada la forma
'             manda un valor verdadero
'CREO       : Yazmin Delgado Gómez
'FECHA_CREO : 19-Julio-2007
'MODIFICO: Antonio Salvador Benavides Guardado
'FECHA_MODIFICO: 13/junio/2008
'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Function Formulario_Cargado(Nombre_Formulario As String) As Boolean
Dim Frm_Control As Form

  On Error Resume Next
  'Recorre sólo los controles de tipo forma que esten cargados en el MDI para validar
  'que el nombre del control de la forma a buscar este cargado
  For Each Frm_Control In Forms
    'Si esta cargado entonces retorna un valor verdadero
    If StrComp(Frm_Control.Name, Nombre_Formulario, vbTextCompare) = 0 Then
      Formulario_Cargado = True
      Exit For
    End If
  Next
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Buscar_Cadena_Combo
    'DESCRIPCIÓN: Busca alguna cadena de caracteres dentro de los elementos de un combo
    '             agregados previamente desde diseño o con un Recordset y me
    '             regresa el indice donde lo encuentra
    'PARÁMETROS:
    '             1. Cadena: Es la cadena que se va a buscar dentro de los elementos del Combo
    '             2. ID: Clave del Cliente (Cliente_ID)de la Tabla Cat_Clientes
    '             3. Valor: paso del valor contenido en la caja de texto Txt_
    '             4. Caracteres: Manda el numero de caracteres que se quieren tomar de la
    '                izquierda del elemento del combo para compraralos contra la cadena
    '                que se busca, si caracteres=0 se comprara la cadena completa del elemento
    '                del combo
    'CREO: Susana Ledesma Ramirez
    'FECHA_CREO:2/Mayo/2006
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Buscar_Cadena_Combo(Cadena As String, Combo As ComboBox, Optional Caracteres As Integer) As Integer
Dim I As Integer
Dim Comparar_Cadena As String       'Cadena que se usa para tomar solo los caracetres a comparar de los
                                    'elementos del combo
'El combo debe tener algunos elementos agregados previamente,si no; se deben agregar
'uno por uno sin usar esta funcion

    For I = 0 To Combo.ListCount - 1
        'Asigno todos los caracteres del elemento del combo,
        'cuando no necesito tomar solamente algunos caracteres
        Comparar_Cadena = Combo.List(I)
        'Si necesito solo algunos caracetres corto la cadena
        'utilizando el numero de caracteres que se mandaron en la variable OPCIONAL caracteres
        If Caracteres > 0 Then Comparar_Cadena = Left(Combo.List(I), Caracteres)
        If Comparar_Cadena = Trim(Cadena) Then
            Buscar_Cadena_Combo = I
            Exit For
        End If
    Next I
End Function

'*******************************************************************************
    'NOMBRE DE LA PROPIEDAD: Forma
    'DESCRIPCIÓN: Ajusta al ayudante la Forma a utilizar en el redimensionado
    'PARÁMETROS:               1. Forma_Ajustar La forma que se va a aplicar  el redimensionado
    'CREO: Ruben Garcia
    'FECHA_CREO:21 Nov 2005
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Property Set Forma(Forma_Ajustar As Object)
    Mensage_Error = ""            'inicializamos los errores a Ninguno
    On Error GoTo ErrorRoutine
    Dim Cont_Controles As Integer   'Lleva al cuenta del control
    Dim Control_Emulado As Control  'Emula al control actual
    Dim Altura_Control As Long           'Toma el Alto del  control
    Dim I As Integer                'Contador
    Set Emular_Forma = Forma_Ajustar           'variable que emula la forma afectada
    Set Redim_Forma = Forma_Ajustar
    '-----------------------------------------------
    'guardamos el Alto y Ancho originales de la forma.
    '-----------------------------------------------
    Alto_Original = Emular_Forma.Height
    Ancho_Original = Emular_Forma.Width
    Escala_Alto_Original = Emular_Forma.ScaleHeight
    '-------------------------------------------------------
    'Examinamos la coleccion de controles, determinando el numero
    'de controles que tienen la propiedad Height. Controles como
    'los menus no tienen tal propiedad.
    '-------------------------------------------------------
    Cont_Controles_No_Redim = -1
    On Error Resume Next
    For I = 0 To (Emular_Forma.Controls.Count - 1)
        Altura_Control = Emular_Forma.Controls(I).Height
        If Err = 0 Then
            Cont_Controles_No_Redim = Cont_Controles_No_Redim + 1
        Else
            Err = 0
        End If
    Next
    On Error GoTo ErrorRoutine
    
    If Cont_Controles_No_Redim = -1 Then
        Exit Property
    End If
    '-----------------------------------------------
    'Redimensionamos el arreglo de Controles igual al tamaño
    'de la cantidad de estos que si tienen la propiedad Height
    '-----------------------------------------------
    ReDim Arr_Controles(Cont_Controles_No_Redim)
    '------------------------------------------------------
    'Alamcenamos el numero de colecciones y las dimensiones originales
    'de cada control que tiene la propiedad Height
    '------------------------------------------------------
    I = -1
    For Cont_Controles = 0 To (Emular_Forma.Controls.Count - 1)
        Set Control_Emulado = Emular_Forma.Controls(Cont_Controles)
        With Control_Emulado
            Debug.Print .Name
            'MsgBox lv_Ctl.Name
            On Error Resume Next
            Altura_Control = .Height
            If Err = 0 Then  'El control si tiene la propiedad Height
                I = I + 1
                If Not (I > Cont_Controles_No_Redim) Then 'para que sea seguro
                    Arr_Controles(I).Numero_Control = Cont_Controles
                    Arr_Controles(I).Borde_Superior = .Top
                    Arr_Controles(I).Borde_Izquierdo = .Left
                    Arr_Controles(I).Alto = .Height
                    Arr_Controles(I).Ancho = .Width
                    Arr_Controles(I).f = .Font.Size
                End If
            End If
            On Error GoTo ErrorRoutine
        End With
    Next
    Exit Property
ErrorRoutine: 'Guardamos el error que ocurrio para tomarlo de la propiedad Mensaje_Error
    Mensage_Error = _
        "Eror en " & App.EXEName & ".clsControlResizer.Form.PropertySet" & _
        Chr(13) & Chr(13) & "Error" & Str(Err) & " " & Error
    Exit Property
End Property

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Redimensionar_Controles
    'DESCRIPCIÓN: Redimesiona los controles de la Forma contenida en la propiedad
    'Forma.
    'USO: '         1.  En la seccion de declaraciones de la forma:
    '                     Dim Resizer as Ayudante
    '               2.  En el evento Form_Load , escribe:
    '                     Set Resizer = new ayudante
    '                     Set Resizer.Form = Me
    '                     If Len(Resizer.Mensaje_Error) > 0 then
    '                         MsgBox Resizer.Mensaje_Error
    '                     End If
    '               3.  En el evento  Form_Resize event, escribe:
    '                     Resizer.RedimensionarControles
    '                     If Len(Resizer.Mensaje_Error) > 0 then
    '                         MsgBox Resizer.Mensaje_Error
    '                     End If
    '               4.  Antes de terminar la ejecucion de tu forma, escribe:
    '                     Set Resizer = Nothing
    'PARÁMETROS:
    'CREO: Ruben Garcia
    'FECHA_CREO: 21-Nov-2005
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
    'NOTA: No funciona bien con el contro SSTab
    '      Si se usa este control se debera manejar en su evento Clic la aparicion
    '      Y posicion de los controles contenidos en el
'*******************************************************************************
Public Sub Redimensionar_Controles()
    On Error GoTo ErrorRoutine
    Dim Control_Emulado As Control  'Emula el control actual
    Dim Cadena_Tag As String    'Almacena la cadena destinada a cubrir el Tag
    Dim Posicion_Tag As Integer    'Guarda la posicion del Tag que se va editando
    Dim Caracteres_Tag As String
    Dim No_Redimensionar As Boolean   'indica si algun control no se va a redimensionar
    Dim Redimensionar_Pos_Superior As Boolean
    Dim Redimensionar_Pos_Izquierda As Boolean
    Dim Redimensionar_Alto As Boolean
    Dim Redimensionar_Ancho As Boolean
    Dim Redimensionar_Fuente As Boolean
    Dim Redimensionar_Margen_Forma As Boolean 'indica si se va respetar el margen de la forma
    Dim Nueva_Pos_Superior As Long                     'Almacen ala nueva posicion de altura del control
    Dim Nueva_Pos_Izquierda As Long                     'Almacena la nuev aposicion izquierda del control
    Dim Nueva_Altura As Long                     'Almacena la nueva altura del control
    Dim Nuevo_Ancho As Long                     'Almacena la Nueva Anchura del control
    Dim lv_NewF As Long
    Dim Factor_Altura As Single                'guarda el Factor de altura para el redimensionado
    Dim Factor_Anchura As Single                'guarda el Factor de anchura para el redimensionado
    Dim I As Integer                        'Iterador de bucle
    '-----------------------------------------------------
    'Salida del procedimiento si no hay controles que s epuedan redimensionar.
    '-----------------------------------------------------
    If Cont_Controles_No_Redim = -1 Then
        Exit Sub
    End If
    '---------------------------------------------------
    'Se calcula el factor horizontal y vertical de redimensionado.
    '---------------------------------------------------
    Factor_Altura = Redim_Forma.Height / Alto_Original
    Factor_Anchura = Redim_Forma.Width / Ancho_Original
    '----------------------------------------------
    'recorrido por todos los controles,se redmensiona si es necesario
    '----------------------------------------------
    For I = LBound(Arr_Controles) To UBound(Arr_Controles)
        Set Control_Emulado = Redim_Forma.Controls(Arr_Controles(I).Numero_Control)
        '----------------------------------------------------------
        'Para cambiar el comportamiento habitual del redimensionado, el tag de un control
        'puede contener una cadena como "Resize=TLW" el cual en este caso
        'especifica que el ajuste de tamaño aplica a la propiedaa Top, Left y Width
        'pero no para height.El codigo siguinte se encarga de ello
        '----------------------------------------------------------
        No_Redimensionar = False
        Redimensionar_Pos_Superior = True
        Redimensionar_Pos_Izquierda = True
        Redimensionar_Alto = True
        Redimensionar_Ancho = True
        Redimensionar_Fuente = True
        Redimensionar_Margen_Forma = False
        Cadena_Tag = Control_Emulado.Tag
        Posicion_Tag = InStr(Cadena_Tag, "Resize=")
        If Posicion_Tag <> 0 Then
            Posicion_Tag = Posicion_Tag + 6
            Redimensionar_Pos_Superior = False
            Redimensionar_Pos_Izquierda = False
            Redimensionar_Alto = False
            Redimensionar_Ancho = False
            Redimensionar_Fuente = False
            Do
                Posicion_Tag = Posicion_Tag + 1
                If Posicion_Tag > Len(Cadena_Tag) Then
                    Exit Do
                End If
                Caracteres_Tag = Mid(Cadena_Tag, Posicion_Tag, 1)
                If Caracteres_Tag = Space(1) Then
                    Exit Do
                End If
                Select Case Caracteres_Tag
                    Case "N"
                        No_Redimensionar = True
                    Case "T"
                        Redimensionar_Pos_Superior = True
                    Case "L"
                        Redimensionar_Pos_Izquierda = True
                    Case "H"
                        Redimensionar_Alto = True
                    Case "W"
                        Redimensionar_Ancho = True
                    Case "G"
                        Redimensionar_Margen_Forma = True
                    Case "F"
                End Select
            Loop
        End If
        If No_Redimensionar Then
            'no hace nada
        Else
            '-----------------------------------------
            'Calcular las nuevas dimensiones para el control
            '-----------------------------------------
            With Arr_Controles(I)
                Nueva_Pos_Superior = .Borde_Superior * IIf(Redimensionar_Pos_Superior, Factor_Altura, 1)
                Nueva_Pos_Izquierda = .Borde_Izquierdo * IIf(Redimensionar_Pos_Izquierda, Factor_Anchura, 1)
                Nuevo_Ancho = .Ancho * IIf(Redimensionar_Ancho, Factor_Anchura, 1)
                If Factor_Anchura < Factor_Altura Then
                    lv_NewF = .f * IIf(Redimensionar_Ancho, Factor_Anchura, 1)
                Else
                    lv_NewF = .f * IIf(Redimensionar_Ancho, Factor_Altura, 1)
                End If
                If Redimensionar_Margen_Forma Then
                   '-------------------------------------------------
                   'Ajusta el alto para mantener una distancia constante
                   'desde el bajo de la forma
                   '-------------------------------------------------
                   Nueva_Altura = _
                    (Redim_Forma.ScaleHeight - Nueva_Pos_Superior) - (Escala_Alto_Original - (.Borde_Superior + .Alto))
                Else
                   '------------------------------------------------
                   'calcular la nueva altura basado en el factor de altura
                   'Calculate new height based on new height factor.
                   '------------------------------------------------
                   Nueva_Altura = .Alto * IIf(Redimensionar_Alto, Factor_Altura, 1)
                End If
            End With
            '------------------------------
            'Posición y tamaño del control
            '------------------------------
            On Error Resume Next
            Control_Emulado.Move Nueva_Pos_Izquierda, Nueva_Pos_Superior
            Control_Emulado.Width = Nuevo_Ancho
            Control_Emulado.Height = Nueva_Altura
            Control_Emulado.Font.Size = lv_NewF
            On Error GoTo ErrorRoutine
            
        End If 'No redimensiona
    Next I  'procesar el siguiente control
    Exit Sub
ErrorRoutine:
    Mensage_Error = _
        "Error en " & App.EXEName & ".clsControlResizer.ResizeControls" & _
        Chr(13) & Chr(13) & "Error" & Str(Err) & " " & Error
    Exit Sub
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cargar_Picture
    'DESCRIPCIÓN: Este proceso hace visible la Pic_Control que necesitamos ocultando
    '             todo las demas a la vez ajusta el tamaño de la forma
    'PARÁMETROS:
    '             1. Pic_Control: Nombre del Picture el cual vamos hacer visible
    '             2. Formulario: Nombre del Formulario al cual nos estamos refiriendo
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
'
Public Sub Cargar_Picture(Pic_Control As PictureBox, Formulario As Form)
'Dim Pic_Ocultar As PictureBox    'Ocultar los PictureBox
Dim Control As Control          'Toma la forma del objeto al que esta apuntando en ese momento
    Pic_Control.Top = 0
    Pic_Control.Left = 0
    'Oculta los Pic_Control contenidos en la forma
    For Each Control In Formulario.Controls
        If TypeOf Control Is PictureBox Then
            Control.Visible = False
        End If
    Next
    Pic_Control.Visible = True
    'Se alinea el formulario al centro de la forma
    'Formulario.Left = (Screen.Width - Formulario.Width) \ 2
    'Formulario.Top = (Screen.Height - Formulario.Height) \ 2
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Validacion_Accesos_Sistema
    'DESCRIPCIÓN: 1. Consulta el menu o submenu que se requiere realizar alguna
    '                operación para consultar las operaciones que puede realizar
    '                el rol al cual pertenece el usuario
    '             2. De acuerdo al nombre que tenga asignado el Tag del boton se
    '                habilita o deshabilita el mismo
    'PARÁMETROS : Nombre_Menu_submenu: Indica el nombre del menu o submenu que se
    '                                  desea realizar las operaciones
    '             Forma: Indica el nombre de la forma en el cual se va a validar
    '                    las operaciones
    'CREO       : 28-Mayo-2007
    'FECHA_CREO : Yazmin Delgado Gómez
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Validacion_Accesos_Sistema(Nombre_Menu_Submenu As String, Forma As Form)
Dim Ctrl As Control                             'Indica que control de la forma se esta consultando
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta las operaciones que puede realizar

'Consulta las operaciones que tiene habilitada
Mi_SQL = "SELECT Alta, Cambio, Eliminar, Consultar"
Mi_SQL = Mi_SQL & " FROM Apl_Cat_Accesos"
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Rol_ID & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = '" & Nombre_Menu_Submenu & "'"
Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos.EOF Then
    With Rs_Consulta_Apl_Cat_Accesos
        'Recorre la forma para consulta los botones que tiene
        For Each Ctrl In Forma.Controls
            If TypeOf Ctrl Is CommandButton Then
                Select Case Ctrl.Tag
                    Case "A":
                        If .rdoColumns("Alta") = "S" Then
                            Ctrl.Enabled = True
                        Else
                            Ctrl.Enabled = False
                        End If
                    Case "M"
                        If .rdoColumns("Cambio") = "S" Then
                            Ctrl.Enabled = True
                        Else
                            Ctrl.Enabled = False
                        End If
                    Case "B"
                        If .rdoColumns("Eliminar") = "S" Then
                            Ctrl.Enabled = True
                        Else
                            Ctrl.Enabled = False
                        End If
                    Case "C"
                        If .rdoColumns("Consultar") = "S" Then
                            Ctrl.Enabled = True
                        Else
                            Ctrl.Enabled = False
                        End If
                    Case ""
                        Ctrl.Enabled = True
                End Select
            End If
        Next
    End With
End If
Rs_Consulta_Apl_Cat_Accesos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Conexion
    'DESCRIPCIÓN: Se hace la conexión a la Base de datos verificando si es en SQL
    '             o Acces de acuerdo a los datos proporcionados por el Usuario
    'PARÁMETROS:
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Conexion()
Dim I As Integer        'Cuenta cada una de las lineas del documento Confing.ini
Dim linea As String     'Guarda lo que contenga la linea a la cual se esta apuntando
Dim Servidor As String  'Guarda el nombre del servidor
Dim Base As String      'Guarda el nombre de la base de datos
Dim Tipo  As String     'Identifica en que esta hecha la BD Access o SQL
Dim Password As String  'Guarda la contraseña para el acceso al servidor
Dim Usuario_Conexion As String   'Guarda el nombre válido del usuario para el acceso a la BD
Dim Empresa As String
Dim RFC As String
Dim Direccion As String
Dim cp As String
Dim Ciudad_Edo As String
Dim BaseD As String
    
On Error GoTo Etiqueta
    I = 0
    'LEE EL ARCHIVO DE INICIO DEL SISTEMA
    Open App.Path & "\Config.ini" For Input As #1
        Do While Not EOF(1)
            Line Input #1, linea
            If I = 0 Then Empresa = Trim(linea)
            If I = 1 Then RFC = Trim(linea)
            If I = 2 Then Direccion = Trim(linea)
            If I = 3 Then Direccion = Direccion & " " & Trim(linea)
            If I = 4 Then cp = Trim(linea)
            If I = 5 Then Ciudad_Edo = Trim(linea)
            If I = 6 Then Ciudad_Edo = Ciudad_Edo & " " & Trim(linea)
            If I = 7 Then Tipo = Trim(linea)
            If I = 8 Then Servidor = Trim(linea)
            If I = 9 Then Usuario_Conexion = Trim(linea)
            If I = 10 Then Base = Trim(linea)
            If I = 11 Then Password = Trim(linea)
            If I = 12 Then Nombre_Emisor = Trim(linea)
            If I = 13 Then Calle_Emisor = Trim(linea)
            If I = 14 Then No_Exterior_Emisor = Trim(linea)
            If I = 15 Then No_Interior_Emisor = Trim(linea)
            If I = 16 Then Colonia_Emisor = Trim(linea)
            If I = 17 Then Codigo_Postal_Emisor = Trim(linea)
            If I = 18 Then Localidad_Emisor = Trim(linea)
            If I = 19 Then Municipio_Emisor = Trim(linea)
            If I = 20 Then Estado_Emisor = Trim(linea)
            If I = 21 Then RFC_Emisor = Trim(linea)
            I = I + 1
        Loop
    Close #1
    BaseD = Base
    Conexion_Base.CursorDriver = rdUseServer
    
    ''Conexion_Base.CursorDriver = rdUseOdbc
    
      
    '1. Verifica si la base de datos es en Access o SQL
    '2. Establece la conexión a la base
    If Tipo = "ACCESS" Then
        Conexion_Base.Connect = "Driver={Microsoft Access Driver (*.mdb)};" & _
                                "Dbq=" & App.Path & "\" & Base & ".MDB;" & _
                                "Uid=;Pwd="
                                Par_Fecha = "#"
                                Base_Datos = "A"
    Else
        Conexion_Base.Connect = "uid=" & Usuario_Conexion & ";pwd=" & Password & ";server=" & Servidor & ";" _
                                & "driver={SQL Server};protocol=6.3;readonly=false;database=" & Base & ";dsn=;"
                                Par_Fecha = "'"
                                Server = Servidor
                                Base_Datos = "S"
                                
    End If
    Conexion_Base.EstablishConnection
    Conexion_Base.QueryTimeout = 80
    Nombre_Servidor = Servidor
    Nombre_BD = Base
    Usuario_BD = Usuario_Conexion
    Password_BD = Password
    Load Frm_Apl_Login
    Exit Sub
Etiqueta:
    Load Frm_Apl_Registro_Sistema
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Cantidad_Letras
    'DESCRIPCIÓN: Convierte de cantidad de número a letras
    'PARÁMETROS:
    '             1. Cantidad_Convertir: Es el número al convertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Convierte_Cantidad_Letras(Cantidad_Convertir As String)
Dim I As Integer
Dim Longitud As Integer   'Obtiene la longitud de la cadena (número)
Dim Posicion_Punto As Integer
Dim Analisis As Integer
Dim Pesos As String
Dim Letras As String      'Asigna el nombre al número
Dim Centavos As String
Dim Temp_Letras As String
Dim Uno As Boolean
Centavos = ""
Pesos = ""
Letras = ""
Uno = False
' Conversion  de cantidad de numero a letras
Longitud = Len(Cantidad_Convertir)
Posicion_Punto = InStr(1, Cantidad_Convertir, ".")
    '
    If (Longitud > 0) And (Posicion_Punto <= 13) And Convierte_Valida_Cantidad(Cantidad_Convertir) Then
        If Posicion_Punto > 0 Then
            Centavos = Mid(Cantidad_Convertir, Posicion_Punto + 1, 2)
            Pesos = Mid(Cantidad_Convertir, 1, Posicion_Punto - 1)
            Longitud = Posicion_Punto - 1
        Else
            Centavos = "00"
            Pesos = Cantidad_Convertir
        End If
        I = 1
        If Longitud >= 10 Then            ' Miles de Millones }
            Select Case Longitud
                Case 10                   ' Miles de Millones }
                    If Mid(Pesos, I, 1) <> "1" Then
                        Temp_Letras = Mid(Pesos, I, 1)
                        Letras = Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                        I = I + 1
                    End If
                Case 11                   ' Decenas de Miles de Millones }
                    Letras = Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
    
                Case 12                   ' Centenas de Miles de Millones }
                    Letras = Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
            End Select
            Letras = Letras & " MIL"      ' Agrega la Terminaci¢n MIL }
            Longitud = Longitud - I + 1   ' Actualiza para continuar an lisis }
        End If
        If Longitud >= 7 Then             ' Millones }
            If Letras <> "" Then          ' Hay miles de millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 7)
                    I = I + 1
                    Longitud = Longitud - 1
                Loop
            End If
            Select Case Longitud           ' Millones }
                Case 7
                    If (Letras = "") And (Mid(Pesos, I, 1) = "1") Then Uno = True
                    Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    I = I + 1
                    Longitud = Longitud - 1     ' Actualiza Para continuar an lisis }
                Case 8                          ' Decenas de Millones }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                    Longitud = Longitud - 2
                Case 9                          ' Centenas de Millones }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
                    Longitud = Longitud - 3     ' Actualiza Para continuar an lisis }
            End Select
            If Uno Then
                    ' Si no hay miles de millones y es 1 entonces la
                    ' Terminación debe de ser MILLON y no MILLONES }
                Letras = Letras & " MILLON"
                Uno = False
            Else
                Letras = Letras & " MILLONES"   ' Agrega la terminaci¢n MILLONES }
            End If
        End If
        If Longitud >= 4 Then             ' Miles }
            If Letras <> "" Then          ' Hay millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 1)
                    I = I + 1
                    Longitud = Longitud - 1
                Loop
                If Longitud = 0 Then Uno = True
            End If
            Select Case Longitud           ' Miles
                Case 4
                    If (Mid(Pesos, I, 1) <> "1") Then
                        Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    End If
                    I = I + 1
                    Analisis = 1
                Case 5                      ' Decenas de Miles }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                    Analisis = 2
                Case 6                      ' Centenas de Miles }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
                    Analisis = 3
            End Select
            If Uno Then
                        ' Si no hay miles ni cientos pero hubo millones entonces
                        ' se debe de agregar la preposici¢n de Antes de PESOS }
                Letras = Letras & " DE"
                Uno = False
            Else
                If Longitud >= 4 Then
                    Letras = Letras & " MIL"    ' Agrega la terminaci¢n MIL }
                   End If
                If Longitud >= 4 Then Longitud = Longitud - Analisis
                                                ' Actualiza Para continuar an lisis }
            End If
        End If
        If Longitud >= 1 Then                   ' Cientos }
            If Letras <> "" Then                ' Hay Miles o Millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 1)
                        I = I + 1
                        Longitud = Longitud - 1
                Loop
            End If
            Select Case Longitud                 ' Unidades }
                Case 1
                    If (Letras = "") And (Mid(Pesos, I, 1) = "1") Then Uno = True
                    Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    I = I + 1
                Case 2                           ' Decenas }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                Case 3                           ' Centenas }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
            End Select
            Longitud = 0
        End If
        If Uno Then
            ' La cantidad es un Peso, ya que esta bandera ( Uno ) se prendió
            ' Cuando no Hubo Millones, ni miles y la cantidad era uno, por lo que la terminación es PESO }
            Letras = Letras & " PESO"
        Else
            Letras = Letras & " PESOS"
        End If
        Letras = Letras & " " & Centavos & "/100 M.N."
    End If
    Convierte_Cantidad_Letras = Letras
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Maximo_Catalogo
'DESCRIPCIÓN: Obtine el número del último registro existente en la tabla
'             especificada
'PARÁMETROS:
'             1. Tabla: Nombre de la tabla de la Base de datos
'             2. Campo: Llave primaria de la tabla
'CREO: Jorge Razo
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Maximo_Catalogo(Tabla As String, Campo As String) As Double
Dim Rs_Tabla As rdoResultset
Dim Mi_SQL As New rdoQuery      'Obtiene el valor máximo del registro
    'Consulta el último registro de la tabla
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campo & " FROM " & Tabla & " ORDER BY " & Campo & " DESC"
        .LockType = rdConcurReadOnly
        .RowsetSize = 1
        .CursorType = rdUseOdbc
        Set Rs_Tabla = .OpenResultset(rdConcurReadOnly)
    End With
    'Obtiene el último registro de la tabla a consultar
    If Not Rs_Tabla.EOF Then
        Maximo_Catalogo = Val(Rs_Tabla(0)) + 1
    Else
        Maximo_Catalogo = 1
    End If
    Rs_Tabla.Close
    Mi_SQL.Close
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recordset_Eliminar
    'DESCRIPCIÓN: Crea un recordset para eliminar el registro y lo devuelve
    '             para su uso posterior, basandose en la consulta SQl
    'PARÁMETROS:
    '               1. Mi_SQl: Sentencia del registro a eliminar
    'CREO: Yazmin A Delgado Gómez
    'FECHA_CREO: 06-Junio-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Recordset_Eliminar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurRowVer
        .CursorType = rdOpenDynamic
        .RowsetSize = 1
        Set Recordset_Eliminar = .OpenResultset
    End With
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Recordset_Editar
'DESCRIPCIÓN: Crea un recordset para editar los registros y lo devuelve
'             para su uso posterior, basandose en la consulta SQl
'PARÁMETROS:
'               1. Mi_SQl: Sentencia del registro a modificar
'CREO: Yazmin A Delgado Gómez
'FECHA_CREO: 06-Junio-2005
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Editar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurRowVer
        .CursorType = rdOpenDynamic
        .RowsetSize = 1
        Set Recordset_Editar = .OpenResultset
    End With
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Recordset_Consultar
'DESCRIPCIÓN: Crea un recordset para consultar de solo lectura y lo devuelve
'             para su uso posterior, basandose en la consulta SQl
'PARÁMETROS:
'               1. Mi_SQL: Sentencia a consultar
'CREO: Yazmin A Delgado Gómez
'FECHA_CREO: 06-Junio-2005
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Consultar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurReadOnly
        .CursorType = rdOpenDynamic
        .RowsetSize = 20
        Set Recordset_Consultar = .OpenResultset
    End With
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Convierte_Decenas_Letras
'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
'PARÁMETROS:
'             1. Ciento: Es el número al convertir a letra
'             2. Decenas: este valor es obtenido de la funcion Decenas
'             3. Digito: este valor eds obtenido de la función
'CREO: Joel G. Romero Cervantes
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Convierte_Centenas_Letras(Ciento, Decena, Digito As String)
Dim Letras As String 'para asignarle la cadena a letras
Letras = ""
    Select Case Ciento
        Case "1" 'Para los número del 100 al 199
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "CIEN"
                Else
                    Letras = "CIENTO " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "CIENTO " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "2" 'Para los número del 200 al 299
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "DOSCIENTOS"
                Else
                    Letras = "DOSCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "DOSCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "3" 'Para los número del 300 al 399
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "TRESCIENTOS"
                Else
                    Letras = "TRESCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "TRESCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "4" 'Para los número del 400 al 499
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "CUATROCIENTOS"
                Else
                    Letras = "CUATROCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "CUATROCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "5" 'Para los número del 500 al 599
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "QUINIENTOS"
                Else
                    Letras = "QUINIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "QUINIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "6" 'Para los número del 600 al 699
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "SEISCIENTOS"
                Else
                    Letras = "SEISCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "SEISCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "7" 'Para los número del 700 al 799
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "SETECIENTOS"
                Else
                    Letras = "SETECIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
              Else
                Letras = "SETECIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "8" 'Para los número del 800 al 899
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "OCHOCIENTOS"
                Else
                    Letras = "OCHOCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
              Else
                Letras = "OCHOCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "9" 'Para los número del 900 al 999
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "NOVECIENTOS"
                Else
                    Letras = "NOVECIENTOS " & Convierte_Unidades_Letras(Digito)
              End If
              Else
                Letras = "NOVECIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
    End Select
    Convierte_Centenas_Letras = Letras
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Decenas_Letras
    'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
    'PARÁMETROS:
    '             1. Decena: Es el número al convertir a letra
    '             2. Digito: es el nombre de los números obtenidos de la funcion Unidades
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Decenas_Letras(Decena, Digito As String)
Dim Letras As String 'Para asignarle la cadena a letras
Letras = ""
    'Escribe el nombre de acuerdo al número
    Select Case Decena
        Case "1" 'Para número del 10 al 19
            Select Case Digito
                Case "0"
                    Letras = "DIEZ"
                Case "1"
                    Letras = "ONCE"
                Case "2"
                    Letras = "DOCE"
                Case "3"
                    Letras = "TRECE"
                Case "4"
                    Letras = "CATORCE"
                Case "5"
                    Letras = "QUINCE"
                Case "6", "7", "8", "9"
                    Letras = "DIECI" & Convierte_Unidades_Letras(Digito)
              End Select
        Case "2" 'Para número del 20 al 29
            If Digito = "0" Then
                Letras = "VEINTE"
            Else
                Letras = "VEINTI" & Convierte_Unidades_Letras(Digito)
            End If
        Case "3" 'Para número del 30 al 39
            If Digito = "0" Then
                Letras = "TREINTA"
            Else
                Letras = "TREINTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "4" 'Para número del 40 al 49
            If Digito = "0" Then
                Letras = "CUARENTA"
            Else
                Letras = "CUARENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "5" 'Para número del 50 al 59
            If Digito = "0" Then
                Letras = "CINCUENTA"
            Else
                Letras = "CINCUENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "6" 'Para número del 60 al 69
            If Digito = "0" Then
                Letras = "SESENTA"
            Else
                Letras = "SESENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "7" 'Para número del 70 al 79
            If Digito = "0" Then
                Letras = "SETENTA"
            Else
                Letras = "SETENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "8" 'Para número del 80 al 89
            If Digito = "0" Then
                Letras = "OCHENTA"
            Else
                Letras = "OCHENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "9" 'Para número del 90 al 99
            If Digito = "0" Then
                Letras = "NOVENTA"
            Else
                Letras = "NOVENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
    End Select
    Convierte_Decenas_Letras = Letras
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Unidades_Letras
    'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
    'PARÁMETROS:
    '             1. Digito Es el número al convertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Unidades_Letras(Digito As String)
    'Escribe el nombre de los número del 1 al 9
    Select Case Digito
        Case "0"
            Convierte_Unidades_Letras = "CERO"
        Case "1"
            Convierte_Unidades_Letras = "UN"
        Case "2"
            Convierte_Unidades_Letras = "DOS"
        Case "3"
            Convierte_Unidades_Letras = "TRES"
        Case "4"
            Convierte_Unidades_Letras = "CUATRO"
        Case "5"
            Convierte_Unidades_Letras = "CINCO"
        Case "6"
            Convierte_Unidades_Letras = "SEIS"
        Case "7"
            Convierte_Unidades_Letras = "SIETE"
        Case "8"
            Convierte_Unidades_Letras = "OCHO"
        Case "9"
            Convierte_Unidades_Letras = "NUEVE"
    End Select
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Valida_Cantidad
    'DESCRIPCIÓN: Sirve para validar la cantidad
    'PARÁMETROS:
    '             1. Cantidad: Es el número a combertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Valida_Cantidad(Cantidad As String)
Dim I As Integer
Dim Lista As String
Lista = "[0-9]"
I = 1
    Do While (I <= Len(Cantidad)) And (Mid(Cantidad, I, 1) Like Lista)
        I = I + 1
    Loop
    If I < Len(Cantidad) Then
        If Mid(Cantidad, I, 1) = "." Then
          If (Mid(Cantidad, I + 1, 1) Like Lista) And (Mid(Cantidad, I + 2, 1) Like Lista) Then
            Convierte_Valida_Cantidad = True
          Else
            Convierte_Valida_Cantidad = False
          End If
        Else
            Convierte_Valida_Cantidad = False
        End If
    Else
        Convierte_Valida_Cantidad = True
    End If
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Recordset_Agregar
'DESCRIPCIÓN: Crea un recordset para agregar registros y lo devuelve
'             para su uso posterior, basandose en la consulta SQL
'PARÁMETROS:
'               1. Catalogo: Nombre de la tabla a en la que se agregaran los registros
'CREO: Yazmin A Delgado Gómez
'FECHA_CREO: 06-Junio-2005
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Agregar(Catalogo As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    'Se crea el recorset de Alta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT *"
        .SQL = .SQL & " FROM " & Catalogo & ""
        .LockType = rdConcurRowVer
        .CursorType = rdOpenDynamic
        Set Recordset_Agregar = .OpenResultset
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Elimina_Catalogo
    'DESCRIPCIÓN: Elimina a un registro especifico
    'PARÁMETROS:
    '             1. Catalogo: Nombre de la Tabla de la Base de Datos
    '             2. ID: Clave del Cliente (Cliente_ID)de la Tabla Cat_Clientes
    '             3. Valor: paso del valor contenido en la caja de texto Txt_Cliente_ID
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Elimina_Catalogo(Catalogo As String, ID As String, Valor As String) As Boolean
Dim Rs_Elimina As rdoResultset     'Manejo de registro
Dim Contador As Integer
On Error GoTo handler

    Mi_SQL = "SELECT * FROM " & Catalogo & " WHERE " & ID & " = '" & Valor & "'"
    'Selecciona el registro
    Set Rs_Elimina = Recordset_Eliminar(Mi_SQL)
    'Elimina todos los datos
    If Not Rs_Elimina.EOF Then
        'Sirva para eliminar todos los datos que pueda contener con el mismo ID
        While Not Rs_Elimina.EOF
            Rs_Elimina.Delete
            Rs_Elimina.MoveNext
        Wend
        Rs_Elimina.Close
        Elimina_Catalogo = True
    Else
        Elimina_Catalogo = False
    End If
    Exit Function
    
handler:
    Conexion_Base.RollbackTrans
    
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
    
    Elimina_Catalogo = False
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Limpiar_Textos
'DESCRIPCIÓN: 'Limpia las cajas de texto de un catalogo en espacífico
'PARÁMETROS:
'             1. Forma: Nombre de la Forma a la que queremos borrar el contenido
'                       de las cajas de texto
'CREO: Ruben García
'FECHA_CREO:
'MODIFICO: Antonio Salvador Benavides Guardado
'FECHA_MODIFICO: 14/junio/2008
'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Limpiar_Textos(Forma As Form)
Dim Mi_Text As Object    'Indica que objeto se esta consultando
    '1. Consulta todos los objetos contienidos en la forma
    '2. Limpia las cajas de texto
    For Each Mi_Text In Forma.Controls
        If TypeOf Mi_Text Is TextBox Then Mi_Text.text = ""
        If TypeOf Mi_Text Is ComboBox Then
            Mi_Text.ListIndex = -1
            'Mi_Text.text = ""
        End If
    Next
End Sub

''*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo_Item
    'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
    'PARÁMETROS:
    '             1. Campos: Campo a consultar para el llenado del ComboBox
    '             2. Tabla: Nombre de la tabla a consultar
    '             3. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             4. Tipo: Para saber si esta consultando
    '             5. Dato_Campo: Para consultar y llenar el campo con las palabras introducidas
    '                           por el usuario
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Sub Llena_Combo_Item(Campos As String, Tabla As String, Combo_Control As ComboBox, Tipo As Integer, Dato_Campo As String)
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
'Dim Cont_Campos As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & Dato_Campo & " LIKE '%" & Combo_Control.text & "%'"
        End If
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    
    'Llena el ComboBox de la forma
    If Not Rs_Combo.EOF Then
        Combo_Control.Clear
        While Not Rs_Combo.EOF
            Combo_Control.AddItem Rs_Combo(1)
            Combo_Control.ItemData(Combo_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Llena_Combo
'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
'PARÁMETROS:
'             1. Campos: Campo a consultar para el llenado del ComboBox
'             2. Tabla: Nombre de la tabla a consultar
'             3. Combo_Control: Nombre del ComboBox de la forma el cual se
'                               va a llenar con los valores
'             4. Tipo: Para saber si esta consultando
'             5. Campo_Con: Para consultar y llenar el campo con las palabras introducidas
'                           por el usuario
'CREO: Jorge Razo
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Llena_Combo(Campos As String, Tabla As String, Combo_Control As ComboBox, Tipo As Integer, campo_con As String) As String
Dim Grupo() As String           'Guarda lo contenido en el campo
Dim Cadena As String            'Asignar el valor del campo en el ComboBox
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer

    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & campo_con & " LIKE '%" & Combo_Control.text & "%'"
        End If
        .SQL = .SQL & " ORDER BY Clave"
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    Grupo2 = ""
    If Not Rs_Combo.EOF Then
        Combo_Control.Clear
        Combo_Control.AddItem ""
        Grupo() = Split(Campos, ",")
        campos_cont = UBound(Grupo())
        While Not Rs_Combo.EOF
            Cadena = UCase(Rs_Combo(0))
            Grupo2 = Grupo2 & Rs_Combo(1) & ","
            'For i = 1 To campos_cont - 1
            '    Cadena = Cadena & "  " & Rs_Combo(2)
           ' Next i
            Combo_Control.AddItem Cadena
            'Combo_Control.ItemData(Combo_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Llena_Combo = Grupo2
    Rs_Combo.Close
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Asigna_Item_Combo
'DESCRIPCIÓN: 'Selecciona un registro especifico de un combo
'PARÁMETROS:
'             1. Valor: Clave oculta del registro en el combo
'             2. Cmb_Dato: Combo que contiene la lista de registros
'CREO: Ruben García
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Asigna_Item_Combo(Valor As String, Cmb_Dato As ComboBox)
Dim I As Integer

For I = 0 To Cmb_Dato.ListCount - 1
    If Cmb_Dato.List(I) = Valor Or _
    Format(Cmb_Dato.ItemData(I), "00000") = Valor Or _
    Format(Cmb_Dato.ItemData(I), "0000000000") = Valor Then
        Cmb_Dato.ListIndex = I
        Exit For
    End If
Next I
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Mover_Control_Grid_CheckBox
'DESCRIPCIÓN: Permite colocar un control en un grid en las celdas que se especifiquen
'PARÁMETROS: Flex grid – es el nombre del grid al que se pondra el control
'            Control_Mover es el nombre del control que se quiere poner en el grid
'CREO: Jorge Razo
'FECHA_CREO: 24 Agosto 2006
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Mover_Control_Grid_CheckBox(Flex_Grid As MSFlexGrid, Optional Control_Mover As CheckBox) 'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = (Flex_Grid.Left) + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            .Height = Flex_Grid.CellHeight
            '.Text = Flex_Grid
            .Visible = True
            .SetFocus
            SendKeys "{Home}+{End}"
        End If
    End With
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Mover_Control_Grid_TextBox
'DESCRIPCIÓN: Permite colocar un control en un grid en las celdas que se especifiquen
'PARÁMETROS: Flex grid – es el nombre del grid al que se pondra el control
'            Control_Mover es el nombre del control que se quiere poner en el grid
'CREO: Jorge Razo
'FECHA_CREO: 24 Agosto 2006
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
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

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Valida_Rango_Fechas
    'DESCRIPCIÓN:           Verifica que la fecha inicial no sea mayor a la final, por rango de dias
    '                       Regresa un valor booleano, true si es mayor la fecha inicial, false si es menor
    'PARÁMETROS:            Fecha_Inicial, Fecha inicial de la comparacion
    '                       Fecha_Final, Fecha final de la comparacion
    'CREO       :           Yañez Rodrguez Diego Neftali
    'FECHA_CREO :           09-Enero-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Function Valida_Rango_Fechas(Fecha_Inicial As Date, Fecha_Final As Date) As Boolean
    'Asigna a false la funcion
    Valida_Rango_Fechas = False
    'Utiliza la funcion DateDiff que regresa un valor entero
    If DateDiff("d", Format(Fecha_Final, "MM/dd/yyyy"), Format(Fecha_Inicial, "MM/dd/yyyy")) > 0 Then
        Valida_Rango_Fechas = True
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Generar_Rectangulo
    'DESCRIPCIÓN: De acurdo a las coordenadas que el usuario le pasa a la funcion
    '             dibuja un rectangulo con esas dimensiones
    'PARÁMETROS:
    '             X=El primer valor de donde va a empesar a dibujar el rectangulo
    '               con respecto a lo horizontal
    '             Y=El primero valor en donde va a empesara a dibujar el rectangulo
    '               conrespecto a lo alto
    '            dx=Hasta donde va a dibujar la linea horizontal
    '            dy=Lo Alto del rectangulo
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 14/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
'
Public Sub Generar_Rectangulo(X!, Y!, dx!, dy!, Radio!, op%)
'se genera un rectángulo
'si op es true es de puntas redondeadas
    Printer.DrawWidth = 2
    Select Case op%
        Case True 'esquinas redondeadas
            'ajuste para el radio de las esquinas
            X! = X! + Radio
            Y! = Y! + Radio
            dx! = dx! - 2 * Radio
            dy! = dy! - 2 * Radio
            'esquina izquierda superior
            Printer.Circle (X!, Y!), Radio!, , PI / 2, PI
            'Esquina derecha superior.
            Printer.Circle (X! + dx!, Y!), Radio!, , 0, PI / 2
            'Esquina izquierda inferior.
             Printer.Circle (X!, Y! + dy!), Radio!, , PI, 3 * PI / 2
            'Esquina derecha inferior.
            Printer.Circle (X! + dx!, Y! + dy!), Radio!, , 3 * PI / 2, 2 * PI
            'Ahora las lineas.
            'Linea horizontal de arriba.
            Printer.Line (X!, Y! - Radio!)-(X! + dx!, Y! - Radio!)
            'Linea izquierda vertical.
            Printer.Line (X! - Radio!, Y!)-(X! - Radio!, Y! + dy!)
             'Linea horizontal de abajo.
             Printer.Line (X!, Y! + dy! + Radio!)-(X! + dx!, Y! + dy! + Radio!)
             'Linea derecha vertical.
             Printer.Line (X! + dx! + Radio!, Y!)-(X! + dx! + Radio!, Y! + dy!)
        Case False  'Rectangulo normal.
             Printer.Line (X!, Y!)-(X! + dx!, Y! + dy!), , B
    End Select
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Enfocar
    'DESCRIPCIÓN: Manda el enfoque a la forma con el Caption correspondiente, ya que no se
    '             puede enviar el Setfocus directamente porque se estan utilizando objetos
    '             locales de cada formulario y el enfoque lo envia al formulario original
    '             y no al del objeto que se creo (Menus de Catalogos)
    'PARÁMETROS : Titulo: Es el Caption de la forma que se esta buscando para enviarle el enfoque
    'CREO       : Susana Ledesma Ramírez
    'FECHA_CREO : 13-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Enfocar(Titulo As String)
Dim Forma As Form 'Formulario al que se va a enviar el enfoque
    For Each Forma In Forms
        If Forma.Caption = Titulo Then
            Forma.SetFocus
            Exit For
        End If
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Validar_Email
'DESCRIPCIÓN: Función que comprueba si una dirección de email es válida en estructura
'PARÁMETROS : Email: Contiene la cadena a validar
'CREO       : Jose Antonio Lopez Hernandez
'FECHA_CREO : 04/Junio/2008 06:42 p.m.
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Validar_Email(Email As String) As Boolean
    Dim Fila As Integer
    Dim Longitud As Integer
    Dim Caracter As String
    Dim posicion As Integer
    Dim Bandera_Posicion As Boolean
    Dim Posicion_Arroba As Integer
    Dim Posicion_Punto As Integer

    On Local Error GoTo handler

    'Quita espacios en la cadena
    Email = Trim$(Email)
    'Falida que no este nula o vacia la variable
    If Email = vbNullString Then
        Exit Function
    End If
    'Convierte a minusculas la cadena
    Email = LCase$(Email)
    'Obtiene la longitud de la cadena
    Longitud = Len(Email)
    'Comienza a recorrer la cadena caracter por carecter
    For Fila = 1 To Longitud
        'Asigna a la variable un caracter de la cadena completa para ser validado
        Caracter = Mid(Email, Fila, 1)
        'Valida si el caracter en turno no es una letra o numero
        If (Not (Caracter Like "[a-z]")) And (Not (Caracter Like "[0-9]")) Then
            'Si no es letra o numero valida si dentro del caracter hay un guien bajo, medio o arroba
            If InStr(1, "_-" & "." & "@", Caracter) > 0 Then
                If Bandera_Posicion = True Then
                   Exit Function
                Else
                    Bandera_Posicion = True
                    'Si el guien bajo la arroba o el punto estan al principio o al final de la cadena sale de la funcion
                    If Fila = 1 Or Fila = Longitud Then
                        Exit Function
                    End If
                    'Si es una arroba guarda la posicion donde fue encontrada
                    If Caracter = "@" Then
                        If Posicion_Arroba = 0 Then
                            Posicion_Arroba = Fila
                        Else 'Si la variable ya tenia una posicion diferente de 0 indica que ya habia sido encontrada una arroba
                            'anteriormente
                            Exit Function
                        End If
                    End If
                    'Si es un punto guarda la posicion donde fue encontrado
                    If Caracter = "." Then
                        Posicion_Punto = Fila
                    End If
                End If
            Else
                Exit Function
            End If
        Else
            Bandera_Posicion = False
        End If
    Next Fila
    'Si la la arroba o el punto no fueron encontrados en la cadena sale de la funcion
    If Posicion_Arroba = 0 Or Posicion_Punto = 0 Then
        Exit Function
    End If
    'Si el ultimo punto encontrado esta antes de la arroba sale de la funcion
    If Posicion_Punto < Posicion_Arroba Then
        Exit Function
    End If
    'Devuelve un TRUE indicando que la cadena es valida
    Validar_Email = True
    Exit Function

handler:
    On Local Error Resume Next
    'Devuelve un FALSE si ocurrio algun error o la cadena no es valida
    Validar_Email = False
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Ordenar_Grid
    'DESCRIPCIÓN: Al dar un clic en uno de los encabezados de las columnas, los datos
    '             son ordenados por el criterio de la columna
    'PARÁMETROS : Grid - es el control al que se le aplicará el ordenamiento
    'CREO       : Jose Antonio Lopez Hernandez
    'FECHA_CREO : 17/Junio/2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Ordenar_Grid(Grid As MSFlexGrid)
    'Oculta el grid
    Grid.Visible = False
    Grid.Refresh
    
    'Ordena usando la columna a la que se le dio clic
    Grid.Col = Grid.MouseCol
    Grid.ColSel = Grid.MouseCol
    Grid.Row = 0
    Grid.RowSel = 0
    
    'Indica si se debe ordernar ascendentemente o descendentemente
    If Columna_Orden_Seleccionada <> Grid.MouseCol Then
        Modo_Orden_Grid = flexSortGenericAscending
    ElseIf Modo_Orden_Grid = flexSortGenericAscending Then
        Modo_Orden_Grid = flexSortGenericDescending
    Else
        Modo_Orden_Grid = flexSortGenericAscending
    End If
    Grid.Sort = Modo_Orden_Grid
    
    'Restaura el dato de la otra columna
    If Grid.MouseCol >= 0 Then
        Grid.TextMatrix(0, Grid.MouseCol) = Mid$(Grid.TextMatrix(0, Grid.MouseCol), 1)
    End If
    
    '
    Columna_Orden_Seleccionada = Grid.MouseCol
    If Modo_Orden_Grid = flexSortGenericAscending Then
        Grid.TextMatrix(0, Grid.MouseCol) = Grid.TextMatrix(0, Grid.MouseCol)
    Else
        Grid.TextMatrix(0, Grid.MouseCol) = Grid.TextMatrix(0, Grid.MouseCol)
    End If
    
    '
    Debug.Print Grid.MouseCol; Grid.Col; Columna_Orden_Seleccionada '@
    
    'Muestra el grid
    Grid.Visible = True
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Existe_Archivo_Carpeta
'DESCRIPCIÓN: Función para validar si existe un Archivo o Carpeta en la ruta especificada.
'PARÁMETROS : Ruta_Archivo - cadena con la ruta del directorio a validar.
'             Tipo - Indica si se validará una Carpeta o un Archivo.
'CREO       : Jose Antonio Lopez Hernandez
'FECHA_CREO : 17/Junio/2008
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Valida_Existe_Archivo_Carpeta(Ruta_Archivo As String, Tipo As String) As Boolean
Dim AttrDev%

    'Se valida si existe una carpeta en la ruta especificada.
    If Tipo = "CARPETA" Then
        If CBool(PathIsDirectory(Ruta_Archivo)) Then
            Valida_Existe_Archivo_Carpeta = True
        Else
            Valida_Existe_Archivo_Carpeta = False
        End If
    Else 'Si no es carpeta valida que sea un archivo y que exista.
        On Error Resume Next
        AttrDev = GetAttr(Ruta_Archivo)
        'Condicion para capturar un error y regersar un valor.
        'Si encuentra un error lo limpia y regersa un Valor false, eso es que no encontro el archivo.
        If Err.Number Then
            Err.Clear
            Valida_Existe_Archivo_Carpeta = False
        Else 'Si no encuentra el un error quiere decir que existe un archivo y regresa un valor True
            Valida_Existe_Archivo_Carpeta = True
        End If
    End If
End Function

Public Sub Actualiza_Saldo(Fecha_S As String, Cuenta_S As String)
Dim Rs_Maximo_Movimientos As rdoResultset
Dim Fecha As String
Dim Cuenta, Mi_SQL As String
Dim Bandera As Integer
Dim Consecutivo As Double
Dim Rs_Actualiza_Movimientos As rdoResultset
Dim Rs_Saldo_Movimientos As rdoResultset
Dim Fecha_Saldo As String
Dim Banco As String
Dim Saldo As Double

On Error GoTo MuestraError
    'Obtiene el ultimo movimiento de acuerdo a la fecha
    Fecha = Format(DateAdd("d", -1, Fecha_S), "MM/dd/yyyy")
    Cuenta = Cuenta_S
    Bandera = 0
    Saldo = 0
    'Mi_SQL = "SELECT MAX(Fecha) FROM Adm_Movimientos"
    Mi_SQL = "SELECT TOP 1 Fecha FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Fecha<='" & Fecha & "'"
    Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
    Mi_SQL = Mi_SQL & " ORDER BY 1 DESC"
    Set Rs_Maximo_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Maximo_Movimientos.EOF Then
        If Not IsNull(Rs_Maximo_Movimientos.rdoColumns(0)) Then
            Fecha_Saldo = Format(Rs_Maximo_Movimientos.rdoColumns(0), "MM/dd/yyyy")
            Mi_SQL = "SELECT Saldo FROM Adm_Movimientos"
            Mi_SQL = Mi_SQL & " WHERE Fecha='" & Fecha_Saldo & "'"
            Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha ASC,No_Movimiento"
            Set Rs_Saldo_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Saldo_Movimientos.EOF And Not IsNull(Rs_Saldo_Movimientos.rdoColumns("Saldo")) Then
                Saldo = Rs_Saldo_Movimientos.rdoColumns("Saldo")
            End If
        Else
            Saldo = 0
            Bandera = 1
            Fecha_Saldo = Fecha
        End If
    Else
        Saldo = 0
        Bandera = 1
        Fecha_Saldo = Fecha
    End If
    Rs_Maximo_Movimientos.Close
    'Actualiza saldos
    Mi_SQL = "SELECT * FROM Adm_Movimientos "
    Mi_SQL = Mi_SQL & " WHERE Fecha>='" & Format(Fecha_Saldo, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Movimiento"
    Set Rs_Actualiza_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Saldo = 0 And Bandera = 1 And Not Rs_Actualiza_Movimientos.EOF Then
        Saldo = Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
        If Rs_Actualiza_Movimientos.rdoColumns("Tipo") = "E" Then
            Saldo = Saldo * (-1)
        End If
    End If
    While Not Rs_Actualiza_Movimientos.EOF
        Rs_Actualiza_Movimientos.Edit
            Rs_Actualiza_Movimientos.rdoColumns("Saldo") = Saldo
        Rs_Actualiza_Movimientos.Update
        Rs_Actualiza_Movimientos.MoveNext
        If Not Rs_Actualiza_Movimientos.EOF Then
            If Rs_Actualiza_Movimientos.rdoColumns("Tipo") = "E" Then
                Saldo = Saldo - Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
            Else
                Saldo = Saldo + Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
            End If
        End If
    Wend
    Rs_Actualiza_Movimientos.Close
    Exit Sub
MuestraError:
    MsgBox Err.Description
End Sub


Public Function Limpia_Variables()
    CFD_Generales.Fecha = ""
    CFD_Generales.Fecha_Timbrado = ""
    CFD_Generales.Folio = ""
    CFD_Generales.Forma_Pago = ""
    CFD_Generales.Metodo_Pago = ""
    CFD_Generales.Condiciones_Pago = ""
    CFD_Generales.Descuento = 0
    CFD_Generales.SubTotal = 0
    CFD_Generales.Total = 0
    CFD_Generales.Impuestos = 0
    CFD_Generales.Tipo_Factura = ""
    CFD_Emisor.RFC = ""
    CFD_Emisor.cp = ""
    CFD_Emisor.Regimen_Fiscal = ""
    CFD_Relacionados.Existe = False
    CFD_Generales.Forma_Pago = ""
End Function

