VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Tipo_Factor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Tipo de Factor"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1560
      Picture         =   "Cat_Tipo_Factor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "M"
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5880
      Picture         =   "Cat_Tipo_Factor.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3000
      Picture         =   "Cat_Tipo_Factor.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "B"
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4440
      Picture         =   "Cat_Tipo_Factor.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "C"
      Top             =   3600
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Cat_Tipo_Factor.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "A"
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   1350
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
      Top             =   1440
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid Grid_Factor 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2990
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
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.TextBox Txt_Descripcion_Factor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Txt_Clave_Factor 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Clave Factor"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Tipo de Factor"
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
      Width           =   6945
   End
End
Attribute VB_Name = "Frm_Cat_Tipo_Factor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub Consulta_Factor(Cadena As String)
Dim Rs_Consulta_Cat_Factor As rdoResultset   'Manejo de registro
        Grid_Factor.Rows = 0
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Factor.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipo_Factor"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Tipo_Factor ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Factor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Factor.EOF Then
        'Pone un encabezado en el grid
        Grid_Factor.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Factor
                    Txt_Descripcion_Factor.text = .rdoColumns("Descripcion")
              
        End With
        While Not Rs_Consulta_Cat_Factor.EOF
            Grid_Factor.AddItem Rs_Consulta_Cat_Factor.rdoColumns("Codigo_Tipo_Factor") & Chr(9) & Rs_Consulta_Cat_Factor.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Factor.rdoColumns("Estado")
            Grid_Factor.FixedRows = 1
            Rs_Consulta_Cat_Factor.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Factor.FixedCols = 1
        Grid_Factor.ColWidth(0) = 1300
        Grid_Factor.ColAlignment(0) = flexAlignCenterCenter
        Grid_Factor.ColWidth(1) = 4000
        Grid_Factor.ColAlignment(1) = flexAlignCenterLeft
        Grid_Factor.ColWidth(2) = 1200
        Grid_Factor.ColAlignment(2) = flexAlignCenterLeft
        
    Else
        MsgBox "El tipo de factor no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Factor.Close
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Factor As rdoResultset   'Manejo de registro
    
    Grid_Factor.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Tipo_Factor WHERE (Codigo_Tipo_Factor ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Factor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Factor.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Factor.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Tipo_Factor SET Estado='INACTIVO' WHERE (Codigo_Tipo_Factor='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Tipo_Factor SET Estado='ACTIVO' WHERE (Codigo_Tipo_Factor ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Factor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Factor.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub
Public Sub Modifica_Factor()
Dim Rs_Modificacion_Cat_Factor As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tipo_Factor"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Factor='" & Txt_Clave_Factor & "'"
    Set Rs_Modificacion_Cat_Factor = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Factor.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Factor
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Factor.text)
            .Update
        End With
        
    Else
        MsgBox "El tipo de factor no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Factor.Close
    MsgBox "El producto ha sido modificado", vbInformation
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
    If Txt_Clave_Factor.text <> "" Then
        Consulta_Factor (Txt_Clave_Factor.text)
        
    Else
        MsgBox "Ingrese el tipo de factor", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Clave_Factor <> "" Then
        Cambiar_Estado (Txt_Clave_Factor.text)
        
    Else
        MsgBox "Ingrese el tipo de factor", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Unidad.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Factor.Enabled = True
                                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Clave_Factor) <> "" Then
                    Modifica_Factor
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub

Private Sub Grid_Factor_Click()
Dim Rs_Consulta_Cat_Factor As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Factor.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Factor.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Tipo_Factor"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Factor='" & Grid_Factor.TextMatrix(Grid_Factor.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Factor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Factor.EOF Then
            With Rs_Consulta_Cat_Factor
                Txt_Clave_Factor = .rdoColumns("Codigo_Tipo_Factor")
                Txt_Descripcion_Factor = .rdoColumns("Descripcion")
            End With
        End If
        Rs_Consulta_Cat_Factor.Close
    End If
End Sub

Private Sub Txt_Clave_Factor_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Factor.Rows <> 0 Then
                    Grid_Factor.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Factor.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Clave_Factor.text) <> "" Then
                Alta_Factor
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
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
        
        Txt_Descripcion_Factor.Enabled = False
                
        If Grid_Factor.Rows <> 0 Then
            Grid_Factor.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
End Sub

Public Sub Alta_Factor()

Dim Rs_Alta_Cat_Factor As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Tipo_Factor FROM Cat_Tipo_Factor where Codigo_Tipo_Factor ='" & Trim(Txt_Clave_Factor.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Clave_Factor.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Factor = Conectar_Ayudante.Recordset_Agregar("Cat_Tipo_Factor")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Factor
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipo_Factor", "Clave"), "00")
            .rdoColumns("Codigo_Tipo_Factor") = Txt_Clave_Factor.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Factor.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Factor.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Factor.Rows = 0 Then
        Grid_Factor.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
     Grid_Factor.AddItem UCase(Txt_Clave_Factor.text) & Chr(9) & UCase(Txt_Descripcion_Factor.text) & Chr(9) & UCase("ACTIVO")
    
    
    Grid_Factor.FixedCols = 1
    Grid_Factor.ColWidth(0) = 1300
    Grid_Factor.ColAlignment(0) = flexAlignCenterCenter
    Grid_Factor.ColWidth(1) = 4000
    Grid_Factor.ColAlignment(1) = flexAlignCenterCenter
    Grid_Factor.ColWidth(2) = 1200
    Grid_Factor.ColAlignment(2) = flexAlignCenterCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Consulta()
Dim Rs_Consulta_Cat_Factor As rdoResultset   'Manejo de registro
    
    Grid_Factor.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipo_Factor"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Factor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Factor.EOF Then
        'Pone un encabezado en el grid
        Grid_Factor.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Factor.EOF
            Grid_Factor.AddItem Rs_Consulta_Cat_Factor.rdoColumns("Codigo_Tipo_Factor") & Chr(9) & Rs_Consulta_Cat_Factor.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Factor.rdoColumns("Estado")
            Grid_Factor.FixedRows = 1
            Rs_Consulta_Cat_Factor.MoveNext
        Wend
        'Tamaño de las columnas en el grid
     Grid_Factor.FixedCols = 1
    Grid_Factor.ColWidth(0) = 1300
    Grid_Factor.ColAlignment(0) = flexAlignCenterCenter
    Grid_Factor.ColWidth(1) = 4000
    Grid_Factor.ColAlignment(1) = flexAlignCenterLeft
    Grid_Factor.ColWidth(2) = 1200
    Grid_Factor.ColAlignment(2) = flexAlignCenterLeft
        
    End If
    Rs_Consulta_Cat_Factor.Close

End Sub
