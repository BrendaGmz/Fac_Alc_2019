VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_Cat_Codigo_Postal 
   Caption         =   "Catalogo Código Postal"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1680
      Picture         =   "Facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "M"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6120
      Picture         =   "Facturas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3120
      Picture         =   "Facturas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "B"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4560
      Picture         =   "Facturas.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "C"
      Top             =   5640
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Facturas.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "A"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.TextBox Txt_Clave_Codigo_Postal 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Txt_Clave_Estado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1290
      Width           =   1695
   End
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
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid_Postal 
         Height          =   3135
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         BackColorBkg    =   16777215
         ScrollBars      =   2
      End
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
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   7335
      Begin VB.TextBox Txt_Clave_Localidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox Txt_Clave_Municipio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Clave de Localidad"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Clave del Municipio"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Código Postal"
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Clave del Estado"
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CÓDIGO POSTAL"
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
      Width           =   7305
   End
End
Attribute VB_Name = "Frm_Cat_Codigo_Postal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Txt_Clave_Codigo_Postal_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Clave_Estado_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Clave_Municipio_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Clave_Localidad_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub Consulta_Postal(Cadena As String)
Dim Rs_Consulta_Cat_Postal As rdoResultset   'Manejo de registro
    Txt_Clave_Estado.Enabled = True
    Txt_Clave_Municipio.Enabled = True
    Txt_Clave_Localidad.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Postal.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Codigo_Postal"
    Mi_SQL = Mi_SQL & " WHERE (Clave_Codigo_Postal ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Postal.EOF Then
        'Pone un encabezado en el grid
        Grid_Postal.AddItem "Código Postal" & Chr(9) & "Estado" & Chr(9) & "Municipio" & Chr(9) & "Localidad" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Postal
                Txt_Clave_Codigo_Postal.text = .rdoColumns("Clave_Codigo_Postal")
                Txt_Clave_Estado.text = .rdoColumns("Clave_Estado")
                Txt_Clave_Municipio.text = .rdoColumns("Municipio")
                Txt_Clave_Localidad.text = .rdoColumns("Localidad")
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Postal.EOF
            Grid_Postal.AddItem Rs_Consulta_Cat_Postal.rdoColumns("Clave_Codigo_Postal") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Clave_Estado") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Municipio") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Localidad") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Estado")
            Grid_Postal.FixedRows = 1
            Rs_Consulta_Cat_Postal.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Postal.FixedCols = 1
        Grid_Postal.ColWidth(0) = 2000
        Grid_Postal.ColAlignment(0) = flexAlignCenterCenter
        Grid_Postal.ColWidth(1) = 1000
        Grid_Postal.ColAlignment(1) = flexAlignLeftCenter
        Grid_Postal.ColWidth(2) = 1000
        Grid_Postal.ColAlignment(2) = flexAlignLeftCenter
        Grid_Postal.ColWidth(3) = 1000
        Grid_Postal.ColAlignment(3) = flexAlignLeftCenter
        Grid_Postal.ColWidth(4) = 2000
        Grid_Postal.ColAlignment(4) = flexAlignLeftCenter
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Postal.Close
End Sub

Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Postal As rdoResultset   'Manejo de registro
    
    Grid_Postal.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Codigo_Postal WHERE (Clave_Codigo_Postal ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Postal.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Postal.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Codigo_Postal SET Estado='INACTIVO' WHERE (Clave_Codigo_Postal ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Codigo_Postal SET Estado='ACTIVO' WHERE (Clave_Codigo_Postal ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Postal.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub
Public Sub Modifica_Postal()
Dim Rs_Modificacion_Cat_Postal As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Codigo_Postal"
    Mi_SQL = Mi_SQL & " WHERE Clave_Codigo_Postal='" & Txt_Clave_Codigo_Postal.text & "'"
    Set Rs_Modificacion_Cat_Postal = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Postal.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Postal
            .Edit
                .rdoColumns("Clave_Estado") = UCase(Txt_Clave_Estado.text)
                .rdoColumns("Municipio") = UCase(Txt_Clave_Municipio.text)
                .rdoColumns("Localidad") = UCase(Txt_Clave_Localidad.text)
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Postal.Close
    MsgBox "El código postal ha sido modificado", vbInformation
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Clave_Codigo_Postal.text <> "" Then
        Consulta_Postal (Txt_Clave_Codigo_Postal.text)
        
    Else
        MsgBox "Ingrese el código del producto", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Clave_Codigo_Postal <> "" Then
        Cambiar_Estado (Txt_Clave_Codigo_Postal.text)
        
    Else
        MsgBox "Ingrese el código postal", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
 If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Clave_Codigo_Postal.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Clave_Estado.Enabled = True
                Txt_Clave_Municipio.Enabled = True
                Txt_Clave_Localidad.Enabled = True
                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Clave_Codigo_Postal.text) <> "" And Trim(Txt_Clave_Estado.text) <> "" Then
                    Modifica_Postal
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Postal.Rows <> 0 Then
                    Grid_Postal.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Clave_Estado.Enabled = True
                Txt_Clave_Municipio.Enabled = True
                Txt_Clave_Localidad.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Codigo_Postal", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Clave_Codigo_Postal.text) <> "" And Trim(Txt_Clave_Estado.text) <> "" Then
                Alta_Postal
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub

Public Sub Alta_Postal()
Dim Rs_Alta_Cat_Postal As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Postal As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Clave_Codigo_Postal FROM Cat_Codigo_Postal where Clave_Codigo_Postal ='" & Trim(Txt_Clave_Codigo_Postal.text) & "'"
    Set Rs_Consulta_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Postal.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Postal.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Postal.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Postal = Conectar_Ayudante.Recordset_Agregar("Cat_Codigo_Postal")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Postal
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Codigo_Postal", "Clave"), "00")
            .rdoColumns("Clave_Codigo_Postal") = Txt_Clave_Codigo_Postal.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Clave_Estado") = UCase(Txt_Clave_Estado.text)
            .rdoColumns("Municipio") = UCase(Txt_Clave_Municipio.text)
            .rdoColumns("Localidad") = UCase(Txt_Clave_Localidad.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Postal.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Postal.Rows = 0 Then
        Grid_Postal.AddItem "Código Postal" & Chr(9) & "Estado" & Chr(9) & "Municipio" & Chr(9) & "Localidad" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    Grid_Postal.AddItem UCase(Txt_Clave_Codigo_Postal.text) & Chr(9) & UCase(Trim(Txt_Clave_Estado.text)) & Chr(9) & UCase(Txt_Clave_Municipio.text) & Chr(9) & UCase(Txt_Clave_Localidad.text) & Chr(9) & UCase("ACTIVO")
  
    Grid_Postal.FixedCols = 1
    Grid_Postal.ColWidth(0) = 2000
    Grid_Postal.ColAlignment(0) = flexAlignCenterCenter
    Grid_Postal.ColWidth(1) = 1000
    Grid_Postal.ColAlignment(1) = flexAlignLeftCenter
    Grid_Postal.ColWidth(2) = 1000
    Grid_Postal.ColAlignment(2) = flexAlignLeftCenter
    Grid_Postal.ColWidth(3) = 1000
    Grid_Postal.ColAlignment(3) = flexAlignLeftCenter
    Grid_Postal.ColWidth(4) = 2000
    Grid_Postal.ColAlignment(4) = flexAlignLeftCenter
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
        Txt_Clave_Estado.Enabled = False
        Txt_Clave_Municipio.Enabled = False
        Txt_Clave_Localidad.Enabled = False
                
        If Grid_Postal.Rows <> 0 Then
            Grid_Postal.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Form_Load()
'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
End Sub
Private Sub Txt_Codigo_Producto_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Public Sub Consulta()
Dim Rs_Consulta_Cat_Postal As rdoResultset   'Manejo de registro
    
    Grid_Postal.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Codigo_Postal"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Postal.EOF Then
        'Pone un encabezado en el grid
        Grid_Postal.AddItem "Código Postal" & Chr(9) & "Estado" & Chr(9) & "Municipio" & Chr(9) & "Localidad" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Postal.EOF
            Grid_Postal.AddItem Rs_Consulta_Cat_Postal.rdoColumns("Clave_Codigo_Postal") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Clave_Estado") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Municipio") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Localidad") & Chr(9) & Rs_Consulta_Cat_Postal.rdoColumns("Estado")
            Grid_Postal.FixedRows = 1
            Rs_Consulta_Cat_Postal.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Postal.FixedCols = 1
        Grid_Postal.ColWidth(0) = 2000
        Grid_Postal.ColAlignment(0) = flexAlignCenterCenter
        Grid_Postal.ColWidth(1) = 1000
        Grid_Postal.ColAlignment(1) = flexAlignLeftCenter
        Grid_Postal.ColWidth(2) = 1000
        Grid_Postal.ColAlignment(2) = flexAlignLeftCenter
        Grid_Postal.ColWidth(3) = 1000
        Grid_Postal.ColAlignment(3) = flexAlignLeftCenter
        Grid_Postal.ColWidth(4) = 2000
        Grid_Postal.ColAlignment(4) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Postal.Close
End Sub

Private Sub Grid_Postal_Click()
Dim Rs_Consulta_Cat_Postal As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Postal.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Clave_Estado.Enabled = True
        Txt_Clave_Municipio.Enabled = True
        Txt_Clave_Localidad.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Codigo_Postal"
        Mi_SQL = Mi_SQL & " WHERE Clave_Codigo_Postal='" & Grid_Postal.TextMatrix(Grid_Postal.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Postal = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Postal.EOF Then
            With Rs_Consulta_Cat_Postal
                Txt_Clave_Codigo_Postal.text = .rdoColumns("Clave_Codigo_Postal")
                Txt_Clave_Estado.text = .rdoColumns("Clave_Estado")
                Txt_Clave_Municipio.text = .rdoColumns("Municipio")
                Txt_Clave_Localidad.text = .rdoColumns("Localidad")
                              
            End With
        End If
        Rs_Consulta_Cat_Postal.Close
    End If
End Sub
