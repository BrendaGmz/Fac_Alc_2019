VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Paises 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Países"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2400
      Picture         =   "Cat_Paises.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "M"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7440
      Picture         =   "Cat_Paises.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4080
      Picture         =   "Cat_Paises.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "B"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5760
      Picture         =   "Cat_Paises.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "C"
      Top             =   4680
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   720
      Picture         =   "Cat_Paises.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "A"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Formatos 
      Caption         =   "Formatos"
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
      Height          =   1815
      Left            =   4800
      TabIndex        =   9
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Txt_Validacion 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Txt_Formato_Registro 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Txt_Codigo_Postal 
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Validación"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Registro de Identidad Tributaria"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Código Postal"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Txt_Codigo_Pais 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Pais 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
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
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Paises 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   7
         FixedRows       =   0
         BackColorBkg    =   16777215
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
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Txt_Agrupaciones 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Agrupaciones"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Código del País"
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción "
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
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_Paises 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PAISES"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Frm_Cat_Paises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Pais.text <> "" Then
        Consulta_Pais (Txt_Codigo_Pais.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Consulta_Pais(Cadena As String)
Dim Rs_Consulta_Cat_Paises As rdoResultset   'Manejo de registro
    Txt_Descripcion_Pais.Enabled = True
    Txt_Agrupaciones.Enabled = True
    Fra_Formatos.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Paises.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Pais"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Pais ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Paises = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Paises.EOF Then
        'Pone un encabezado en el grid
        Grid_Paises.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Formato código postal" & Chr(9) & "Registro de identidad tributaria" & Chr(9) & "Validación del registro" & Chr(9) & "Agrupaciones" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Paises
                Txt_Descripcion_Pais.text = .rdoColumns("Descripcion")
                Txt_Agrupaciones.text = .rdoColumns("Agrupaciones")
                Txt_Codigo_Postal.text = Trim(.rdoColumns("Formato_Codigo_Postal"))
                Txt_Formato_Registro.text = Trim(.rdoColumns("Registro_Identidad"))
                Txt_Validacion.text = Trim(.rdoColumns("Validacion"))
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Paises.EOF
            Grid_Paises.AddItem Rs_Consulta_Cat_Paises.rdoColumns("Codigo_Pais") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Formato_Codigo_Postal") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Registro_Identidad") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Validacion") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Agrupaciones") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Estado")
            Grid_Paises.FixedRows = 1
            Rs_Consulta_Cat_Paises.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Paises.FixedCols = 1
        Grid_Paises.ColWidth(0) = 1000
        Grid_Paises.ColAlignment(0) = flexAlignCenterCenter
        Grid_Paises.ColWidth(1) = 2800
        Grid_Paises.ColAlignment(1) = flexAlignLeftCenter
        Grid_Paises.ColWidth(2) = 1700
        Grid_Paises.ColAlignment(2) = flexAlignLeftCenter
        Grid_Paises.ColWidth(3) = 2300
        Grid_Paises.ColAlignment(3) = flexAlignLeftCenter
        Grid_Paises.ColWidth(4) = 1700
        Grid_Paises.ColAlignment(4) = flexAlignLeftCenter
        Grid_Paises.ColWidth(5) = 1050
        Grid_Paises.ColAlignment(5) = flexAlignLeftCenter
        Grid_Paises.ColWidth(6) = 1000
        Grid_Paises.ColAlignment(6) = flexAlignLeftCenter
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Paises.Close
End Sub


Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Pais <> "" Then
        Cambiar_Estado (Txt_Codigo_Pais.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Paises As rdoResultset   'Manejo de registro
    
    Grid_Paises.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Pais WHERE (Codigo_Pais ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Paises = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Paises.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Paises.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Pais SET Estado='INACTIVO' WHERE (Codigo_Pais ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Pais SET Estado='ACTIVO' WHERE (Codigo_Pais ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Paises = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Paises.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Pais.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Pais.Enabled = True
                Txt_Agrupaciones.Enabled = True
                Fra_Formatos.Enabled = True
                    
            Else
                MsgBox "Ingrese código de país", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Codigo_Pais.text) <> "" And Trim(Txt_Descripcion_Pais.text) <> "" Then
                    Modifica_Pais
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Pais()
Dim Rs_Modificacion_Cat_Paises As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Pais"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Pais='" & Txt_Codigo_Pais.text & "'"
    Set Rs_Modificacion_Cat_Paises = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Paises.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Paises
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Pais.text)
                .rdoColumns("Formato_Codigo_Postal") = UCase(Txt_Codigo_Postal.text)
                .rdoColumns("Registro_Identidad") = UCase(Txt_Formato_Registro.text)
                .rdoColumns("Validacion") = UCase(Txt_Validacion.text)
                .rdoColumns("Agrupaciones") = UCase(Txt_Agrupaciones.text)
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Paises.Close
    MsgBox "La información ha sido actualizada", vbInformation
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Paises.Rows <> 0 Then
                    Grid_Paises.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Pais.Enabled = True
                Txt_Agrupaciones.Enabled = True
                Fra_Formatos.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Pais", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Pais.text) <> "" And Trim(Txt_Codigo_Pais.text) <> "" Then
                Alta_Paises
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Paises()
Dim Rs_Alta_Cat_Paises As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Pais As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Pais FROM Cat_Pais where Codigo_Pais ='" & Trim(Txt_Codigo_Pais.text) & "'"
    Set Rs_Consulta_Pais = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Pais.EOF Then
        MsgBox "El código que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Pais.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Pais.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Paises = Conectar_Ayudante.Recordset_Agregar("Cat_Pais")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Paises
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Codigo_Postal", "Clave"), "00")
            .rdoColumns("Codigo_Pais") = Txt_Codigo_Pais.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Pais.text)
            .rdoColumns("Formato_Codigo_Postal") = UCase(Txt_Codigo_Postal.text)
            .rdoColumns("Registro_Identidad") = UCase(Txt_Formato_Registro.text)
            .rdoColumns("Validacion") = UCase(Txt_Validacion.text)
            .rdoColumns("Agrupaciones") = UCase(Txt_Agrupaciones.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Paises.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Paises.Rows = 0 Then
        Grid_Paises.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Formato código postal" & Chr(9) & "Registro de identidad tributaria" & Chr(9) & "Validación del registro" & Chr(9) & "Agrupaciones" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    Grid_Paises.AddItem UCase(Txt_Codigo_Pais.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Pais.text)) & Chr(9) & UCase(Txt_Codigo_Postal.text) & Chr(9) & UCase(Txt_Formato_Registro.text) & Chr(9) & UCase(Txt_Validacion.text) & Chr(9) & UCase(Txt_Agrupaciones.text) & Chr(9) & UCase("ACTIVO")
  
        Grid_Paises.FixedCols = 1
        Grid_Paises.ColWidth(0) = 1000
        Grid_Paises.ColAlignment(0) = flexAlignCenterCenter
        Grid_Paises.ColWidth(1) = 2800
        Grid_Paises.ColAlignment(1) = flexAlignLeftCenter
        Grid_Paises.ColWidth(2) = 1700
        Grid_Paises.ColAlignment(2) = flexAlignLeftCenter
        Grid_Paises.ColWidth(3) = 2300
        Grid_Paises.ColAlignment(3) = flexAlignLeftCenter
        Grid_Paises.ColWidth(4) = 1700
        Grid_Paises.ColAlignment(4) = flexAlignLeftCenter
        Grid_Paises.ColWidth(5) = 1050
        Grid_Paises.ColAlignment(5) = flexAlignLeftCenter
        Grid_Paises.ColWidth(6) = 1000
        Grid_Paises.ColAlignment(6) = flexAlignLeftCenter
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
        If Btn_Nuevo.Caption = "Nuevo" Then
            Txt_Descripcion_Pais.Enabled = False
            Txt_Agrupaciones.Enabled = False
            Fra_Formatos.Enabled = False
        End If
        Grid_Paises.Rows = 0
        Consulta
    End If
End Sub

Private Sub Form_Load()
'Set Conexion = New Conectar
   ' Conexion.ConectarBD
    Consulta
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Paises As rdoResultset   'Manejo de registro
    
    Grid_Paises.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Pais"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Paises = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Paises.EOF Then
        'Pone un encabezado en el grid
        Grid_Paises.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Formato código postal" & Chr(9) & "Registro de identidad tributaria" & Chr(9) & "Validación del registro" & Chr(9) & "Agrupaciones" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Paises.EOF
            Grid_Paises.AddItem Rs_Consulta_Cat_Paises.rdoColumns("Codigo_Pais") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Formato_Codigo_Postal") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Registro_Identidad") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Validacion") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Agrupaciones") & Chr(9) & Rs_Consulta_Cat_Paises.rdoColumns("Estado")
            Grid_Paises.FixedRows = 1
            Rs_Consulta_Cat_Paises.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Paises.FixedCols = 1
        Grid_Paises.ColWidth(0) = 1000
        Grid_Paises.ColAlignment(0) = flexAlignCenterCenter
        Grid_Paises.ColWidth(1) = 2800
        Grid_Paises.ColAlignment(1) = flexAlignLeftCenter
        Grid_Paises.ColWidth(2) = 1700
        Grid_Paises.ColAlignment(2) = flexAlignLeftCenter
        Grid_Paises.ColWidth(3) = 2300
        Grid_Paises.ColAlignment(3) = flexAlignLeftCenter
        Grid_Paises.ColWidth(4) = 1700
        Grid_Paises.ColAlignment(4) = flexAlignLeftCenter
        Grid_Paises.ColWidth(5) = 1050
        Grid_Paises.ColAlignment(5) = flexAlignLeftCenter
        Grid_Paises.ColWidth(6) = 1000
        Grid_Paises.ColAlignment(6) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Paises.Close
End Sub

Private Sub Grid_Paises_Click()
Dim Rs_Consulta_Cat_Paises As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Paises.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Pais.Enabled = True
        Txt_Agrupaciones.Enabled = True
        Fra_Formatos.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Pais"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Pais='" & Grid_Paises.TextMatrix(Grid_Paises.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Paises = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Paises.EOF Then
            With Rs_Consulta_Cat_Paises
                Txt_Codigo_Pais.text = .rdoColumns("Codigo_Pais")
                Txt_Descripcion_Pais.text = .rdoColumns("Descripcion")
                Txt_Agrupaciones.text = .rdoColumns("Agrupaciones")
                Txt_Codigo_Postal.text = Trim(.rdoColumns("Formato_Codigo_Postal"))
                Txt_Formato_Registro.text = Trim(.rdoColumns("Registro_Identidad"))
                Txt_Validacion.text = Trim(.rdoColumns("Validacion"))
            End With
        End If
        Rs_Consulta_Cat_Paises.Close
    End If
End Sub
Private Sub Txt_Codigo_Pais_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
