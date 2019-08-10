VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Aduanas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de Aduanas"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   LinkTopic       =   "AduanasForm"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8445
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Cat_Aduanas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "A"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5235
      Picture         =   "Cat_Aduanas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "C"
      Top             =   5640
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3525
      Picture         =   "Cat_Aduanas.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "B"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6945
      Picture         =   "Cat_Aduanas.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1830
      Picture         =   "Cat_Aduanas.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "M"
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.TextBox Txt_Codigo_Aduana 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Aduana 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid Grid_Aduanas 
         Height          =   2775
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   0
         Cols            =   3
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8175
      Begin VB.Label Label2 
         Caption         =   "Código de aduana"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
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
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "ADUANAS"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8145
   End
End
Attribute VB_Name = "Frm_Cat_Aduanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Clave As String

Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Aduanas As rdoResultset   'Manejo de registro
    
    Grid_Aduanas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Aduanas WHERE (Aduana_ID ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Aduanas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Aduanas.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Aduanas.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Aduanas SET Estado='INACTIVO' WHERE (Aduana_ID ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Aduanas SET Estado='ACTIVO' WHERE (Aduana_ID ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Aduanas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Aduanas.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Aduanas As rdoResultset   'Manejo de registro
    
    Grid_Aduanas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Aduanas"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Aduanas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Aduanas.EOF Then
        'Pone un encabezado en el grid
        Grid_Aduanas.AddItem "Código Aduana" & Chr(9) & "Descripcion" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Aduanas.EOF
            Grid_Aduanas.AddItem Rs_Consulta_Cat_Aduanas.rdoColumns("Aduana_ID") & Chr(9) & Rs_Consulta_Cat_Aduanas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Aduanas.rdoColumns("Estado")
            Grid_Aduanas.FixedRows = 1
            Rs_Consulta_Cat_Aduanas.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Aduanas.FixedCols = 1
        Grid_Aduanas.ColWidth(0) = 1200
        Grid_Aduanas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Aduanas.ColWidth(1) = 4815
        Grid_Aduanas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Aduanas.ColWidth(2) = 1800
        Grid_Aduanas.ColAlignment(2) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Aduanas.Close
End Sub

Public Sub Consulta_Aduanas(Cadena As String)
Dim Rs_Consulta_Cat_Aduanas As rdoResultset   'Manejo de registro
    Txt_Descripcion_Aduana.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Aduanas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Aduanas"
    Mi_SQL = Mi_SQL & " WHERE (Aduana_ID ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Aduanas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Aduanas.EOF Then
        'Pone un encabezado en el grid
        Grid_Aduanas.AddItem "Código Aduana" & Chr(9) & "Descripcion" & Chr(9) & "Estado"
        'Cargar valores
        With Rs_Consulta_Cat_Aduanas
                Txt_Codigo_Aduana.text = .rdoColumns("Aduana_ID")
                Txt_Descripcion_Aduana.text = .rdoColumns("Descripcion")
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Aduanas.EOF
            Grid_Aduanas.AddItem Rs_Consulta_Cat_Aduanas.rdoColumns("Aduana_ID") & Chr(9) & Rs_Consulta_Cat_Aduanas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Aduanas.rdoColumns("Estado")
            Grid_Aduanas.FixedRows = 1
            Rs_Consulta_Cat_Aduanas.MoveNext
        Wend
        
        'Tamaño de las columnas en el grid
        Grid_Aduanas.FixedCols = 1
        Grid_Aduanas.ColWidth(0) = 1200
        Grid_Aduanas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Aduanas.ColWidth(1) = 4815
        Grid_Aduanas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Aduanas.ColWidth(2) = 1800
        Grid_Aduanas.ColAlignment(2) = flexAlignLeftCenter
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Aduanas.Close
End Sub


Public Sub Modifica_Productos()
Dim Rs_Modificacion_Cat_Aduanas As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Aduana As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Aduanas"
    Mi_SQL = Mi_SQL & " WHERE Aduana_ID='" & Txt_Codigo_Aduana.text & "'"
    Set Rs_Modificacion_Cat_Aduanas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Aduanas.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Aduanas
            .Edit
                .rdoColumns("Descripcion") = Trim(UCase(Txt_Descripcion_Aduana.text))
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Aduanas.Close
    MsgBox "El producto ha sido modificado", vbInformation
    
    'Configura el grid
    Grid_Aduanas.TextMatrix(1, 0) = UCase(Trim(Txt_Codigo_Aduana.text))
    Grid_Aduanas.TextMatrix(1, 1) = UCase(Trim(Txt_Descripcion_Aduana.text))
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Alta_Productos()
Dim Rs_Alta_Cat_Aduana As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Aduana_ID FROM Cat_Aduanas where Aduana_ID ='" & Trim(Txt_Codigo_Aduana.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Aduana.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Aduana = Conectar_Ayudante.Recordset_Agregar("Cat_Aduanas")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Aduana
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Aduanas", "Clave"), "00")
            .rdoColumns("Aduana_ID") = Txt_Codigo_Aduana.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Aduana.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Aduana.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    'Btn_Salir.Caption = "Salir"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Coloca un encabezado en la primera fila del grid
    If Grid_Aduanas.Rows = 0 Then
        Grid_Aduanas.AddItem "Código Aduana" & Chr(9) & "Descripción" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    Grid_Aduanas.AddItem Txt_Codigo_Aduana.text & Chr(9) & UCase(Trim(Txt_Descripcion_Aduana.text)) & Chr(9) & UCase("ACTIVO")
      Grid_Aduanas.FixedCols = 1
        Grid_Aduanas.ColWidth(0) = 1200
        Grid_Aduanas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Aduanas.ColWidth(1) = 4815
        Grid_Aduanas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Aduanas.ColWidth(2) = 1800
        Grid_Aduanas.ColAlignment(2) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Aduana.text <> "" Then
        Consulta_Aduanas (Txt_Codigo_Aduana.text)
        
    Else
        MsgBox "Ingrese el código de aduana", vbInformation
    End If
        
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Aduana.text <> "" Then
        Cambiar_Estado (Txt_Codigo_Aduana.text)
        
    Else
        MsgBox "Ingrese el código de aduana", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Aduana.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Txt_Descripcion_Aduana.Enabled = True
                    
            Else
                MsgBox "Ingrese código de Aduana", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion_Aduana.text) <> "" Then
                    Modifica_Productos
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
        
    
End Sub

Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Aduanas.Rows <> 0 Then
                    Grid_Aduanas.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Aduana.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Codigo_Aduana.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Aduanas", "Aduana_ID"), "00")
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Aduanas", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Aduana.text) <> "" Then
                Alta_Productos
                Btn_Nuevo.Caption = "Nuevo"
                Btn_Modificar.Enabled = True
                Btn_Eliminar.Enabled = True
                Btn_Modificar.Caption = "Actualizar"
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
        'Txt_Codigo_Aduana.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Aduanas", "Aduana_ID"), "00")
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
        If Grid_Aduanas.Rows <> 0 Then
            Grid_Aduanas.Rows = 0
            End If
        If Btn_Nuevo.Caption = "Nuevo" Then
            Txt_Descripcion_Aduana.Enabled = False
        End If
        Grid_Aduanas.Rows = 0
        Consulta
    End If

End Sub

Private Sub Grid_Aduanas_Click()
Dim Rs_Consulta_Cat_Aduanas As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Aduanas.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Aduana.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Aduanas"
        Mi_SQL = Mi_SQL & " WHERE Aduana_ID='" & Grid_Aduanas.TextMatrix(Grid_Aduanas.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Aduanas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Aduanas.EOF Then
            With Rs_Consulta_Cat_Aduanas
                Txt_Codigo_Aduana.text = .rdoColumns("Aduana_ID")
                Txt_Descripcion_Aduana.text = .rdoColumns("Descripcion")
                
            End With
        End If
        Rs_Consulta_Cat_Aduanas.Close
    End If
End Sub
