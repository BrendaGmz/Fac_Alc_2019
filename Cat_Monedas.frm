VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_Cat_Monedas 
   Caption         =   "Catalogo de Monedas"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   360
      Picture         =   "Cat_Monedas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "A"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4680
      Picture         =   "Cat_Monedas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "C"
      Top             =   4080
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3240
      Picture         =   "Cat_Monedas.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "B"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6120
      Picture         =   "Cat_Monedas.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1800
      Picture         =   "Cat_Monedas.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "M"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.TextBox Txt_Codigo_Moneda 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Moneda 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
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
      Top             =   1920
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid_Monedas 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   5
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
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   7335
      Begin VB.TextBox Txt_Variacion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Txt_Decimal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   6
         Left            =   6960
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Decimales"
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
         Index           =   2
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Porcentaje de variación"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Moneda"
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
         Width           =   2055
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MONEDAS"
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
      TabIndex        =   11
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Frm_Cat_Monedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Moneda.text <> "" Then
        Consulta_Moneda (Txt_Codigo_Moneda.text)
        
    Else
        MsgBox "Ingrese el código de la moneda", vbInformation
    End If
End Sub
Public Sub Consulta_Moneda(Cadena As String)
Dim Rs_Consulta_Cat_Monedas As rdoResultset   'Manejo de registro
    Txt_Descripcion_Moneda.Enabled = True
    Txt_Variacion.Enabled = True
    Txt_Decimal.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Monedas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Moneda"
    Mi_SQL = Mi_SQL & " WHERE (Moneda_ID ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Monedas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Monedas.EOF Then
        'Pone un encabezado en el grid
        Grid_Monedas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Decimales" & Chr(9) & "Variación" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Monedas
                Txt_Descripcion_Moneda.text = .rdoColumns("Descripcion")
                Txt_Decimal.text = .rdoColumns("Decimales")
                Txt_Variacion.text = (.rdoColumns("Variacion") * 100)
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Monedas.EOF
            Grid_Monedas.AddItem Rs_Consulta_Cat_Monedas.rdoColumns("Moneda_ID") & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Decimales") & Chr(9) & FormatNumber((Rs_Consulta_Cat_Monedas.rdoColumns("Variacion") * 100), 0) & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Estado")
            Grid_Monedas.FixedRows = 1
            Rs_Consulta_Cat_Monedas.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Monedas.FixedCols = 1
        Grid_Monedas.ColWidth(0) = 1000
        Grid_Monedas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Monedas.ColWidth(1) = 2800
        Grid_Monedas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(2) = 1000
        Grid_Monedas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(3) = 1000
        Grid_Monedas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(4) = 1000
        Grid_Monedas.ColAlignment(4) = flexAlignLeftCenter
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Monedas.Close
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Moneda <> "" Then
        Cambiar_Estado (Txt_Codigo_Moneda.text)
        
    Else
        MsgBox "Ingrese el código de la moneda", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Monedas As rdoResultset   'Manejo de registro
    
    Grid_Monedas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Moneda WHERE (Moneda_ID ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Monedas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Monedas.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Monedas.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Moneda SET Estado='INACTIVO' WHERE (Moneda_ID ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Moneda SET Estado='ACTIVO' WHERE (Moneda_ID ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Monedas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Monedas.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Moneda.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Moneda.Enabled = True
                Txt_Decimal.Enabled = True
                Txt_Variacion.Enabled = True
                    
            Else
                MsgBox "Ingrese código de la moneda", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion_Moneda.text) <> "" And Trim(Txt_Decimal.text) <> "" Then
                    Modifica_Moneda
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Moneda()
Dim Rs_Modificacion_Cat_Monedas As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Moneda"
    Mi_SQL = Mi_SQL & " WHERE Moneda_ID='" & Txt_Codigo_Moneda.text & "'"
    Set Rs_Modificacion_Cat_Monedas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Monedas.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Monedas
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Moneda.text)
                .rdoColumns("Decimales") = UCase(Txt_Decimal.text)
                .rdoColumns("Variacion") = Format((UCase(Txt_Variacion.text) / 100), "#0.000")
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Monedas.Close
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
                If Grid_Monedas.Rows <> 0 Then
                    Grid_Monedas.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Moneda.Enabled = True
                Txt_Decimal.Enabled = True
                Txt_Variacion.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Codigo_Postal", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Codigo_Moneda.text) <> "" And Trim(Txt_Descripcion_Moneda.text) <> "" And Trim(Txt_Decimal.text) <> "" Then
                Alta_Monedas
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Monedas()
Dim Rs_Alta_Cat_Monedas As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Moneda As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Moneda_ID FROM Cat_Moneda where Moneda_ID ='" & Trim(Txt_Codigo_Moneda.text) & "'"
    Set Rs_Consulta_Moneda = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Moneda.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Moneda.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Moneda.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Monedas = Conectar_Ayudante.Recordset_Agregar("Cat_Moneda")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Monedas
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Codigo_Postal", "Clave"), "00")
            .rdoColumns("Moneda_ID") = Txt_Codigo_Moneda.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Moneda.text)
            .rdoColumns("Decimales") = UCase(Txt_Decimal.text)
            .rdoColumns("Variacion") = Format((UCase(Txt_Variacion.text) / 100), "#0.000")
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Monedas.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Monedas.Rows = 0 Then
        Grid_Monedas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Decimales" & Chr(9) & "Variación" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    Grid_Monedas.AddItem UCase(Txt_Codigo_Moneda.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Moneda.text)) & Chr(9) & UCase(Txt_Decimal.text) & Chr(9) & UCase(Txt_Variacion.text) & Chr(9) & UCase("ACTIVO")
  
        Grid_Monedas.FixedCols = 1
        Grid_Monedas.ColWidth(0) = 1000
        Grid_Monedas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Monedas.ColWidth(1) = 2800
        Grid_Monedas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(2) = 1000
        Grid_Monedas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(3) = 1000
        Grid_Monedas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(4) = 1000
        Grid_Monedas.ColAlignment(4) = flexAlignLeftCenter
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
            Txt_Descripcion_Moneda.Enabled = False
            Txt_Variacion.Enabled = False
            Txt_Decimal.Enabled = False
        End If
        Grid_Monedas.Rows = 0
        Consulta
    End If
End Sub

Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Monedas As rdoResultset   'Manejo de registro
    
    Grid_Monedas.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Moneda"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Monedas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Monedas.EOF Then
        'Pone un encabezado en el grid
        Grid_Monedas.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Decimales" & Chr(9) & "Variación" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Monedas.EOF
            Grid_Monedas.AddItem Rs_Consulta_Cat_Monedas.rdoColumns("Moneda_ID") & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Decimales") & Chr(9) & FormatNumber((Rs_Consulta_Cat_Monedas.rdoColumns("Variacion") * 100), 0) & Chr(9) & Rs_Consulta_Cat_Monedas.rdoColumns("Estado")
            Grid_Monedas.FixedRows = 1
            Rs_Consulta_Cat_Monedas.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Monedas.FixedCols = 1
        Grid_Monedas.ColWidth(0) = 1000
        Grid_Monedas.ColAlignment(0) = flexAlignCenterCenter
        Grid_Monedas.ColWidth(1) = 2800
        Grid_Monedas.ColAlignment(1) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(2) = 1000
        Grid_Monedas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(3) = 1000
        Grid_Monedas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Monedas.ColWidth(4) = 1000
        Grid_Monedas.ColAlignment(4) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Monedas.Close
End Sub

Private Sub Grid_Monedas_Click()
Dim Rs_Consulta_Cat_Monedas As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Monedas.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Moneda.Enabled = True
        Txt_Decimal.Enabled = True
        Txt_Variacion.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Moneda"
        Mi_SQL = Mi_SQL & " WHERE Moneda_ID='" & Grid_Monedas.TextMatrix(Grid_Monedas.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Monedas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Monedas.EOF Then
            With Rs_Consulta_Cat_Monedas
                Txt_Codigo_Moneda.text = .rdoColumns("Moneda_ID")
                Txt_Descripcion_Moneda = .rdoColumns("Descripcion")
                Txt_Decimal = .rdoColumns("Decimales")
                Txt_Variacion = FormatNumber((.rdoColumns("Variacion") * 100), 0)
                
                
                              
            End With
        End If
        Rs_Consulta_Cat_Monedas.Close
    End If
End Sub

Private Sub Txt_Decimal_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Variacion_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Codigo_Moneda_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
