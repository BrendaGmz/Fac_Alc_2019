VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_Cat_Impuestos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Impuestos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "A"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4440
      Picture         =   "Form2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "C"
      Top             =   6360
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3000
      Picture         =   "Form2.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "B"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5880
      Picture         =   "Form2.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1560
      Picture         =   "Form2.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "M"
      Top             =   6360
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
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid Grid_Impuestos 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   0
         Cols            =   7
         FixedRows       =   0
         BackColorBkg    =   16777215
      End
   End
   Begin VB.TextBox Txt_Descripcion_Impuesto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Txt_Clave_Impuesto 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1575
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
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7095
      Begin VB.TextBox Txt_Entidad 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Cmb_Tipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form2.frx":0552
         Left            =   5040
         List            =   "Form2.frx":055C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Traslado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form2.frx":0570
         Left            =   5040
         List            =   "Form2.frx":057A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Retencion 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form2.frx":0586
         Left            =   5040
         List            =   "Form2.frx":0590
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Entidad en la que aplica"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
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
         Index           =   4
         Left            =   4560
         TabIndex        =   10
         Top             =   1200
         Width           =   495
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
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Clave de Impuesto"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Retención"
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
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Traslado"
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
         Index           =   3
         Left            =   4200
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IMPUESTOS"
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
      TabIndex        =   14
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Frm_Cat_Impuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Consulta_Impuestos(Cadena As String)
Dim Rs_Consulta_Cat_Impuestos As rdoResultset   'Manejo de registro
    
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Impuestos.Rows = 0
    Txt_Descripcion_Impuesto.Enabled = True
    Txt_Entidad.Enabled = True
    Cmb_Retencion.Enabled = True
    Cmb_Traslado.Enabled = True
    Cmb_Tipo.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Impuestos"
    Mi_SQL = Mi_SQL & " WHERE (Impuesto_ID ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Impuestos.EOF Then
        'Pone un encabezado en el grid
        Grid_Impuestos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Retención" & Chr(9) & "Traslado" & Chr(9) & "Tipo" & Chr(9) & "Entidad" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Impuestos
                 Txt_Descripcion_Impuesto.text = .rdoColumns("Descripcion")
                For I = 0 To Cmb_Retencion.ListCount - 1
                    Cmb_Retencion.ListIndex = I
                    If Cmb_Retencion.text = .rdoColumns("Retencion") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Traslado.ListCount - 1
                    Cmb_Traslado.ListIndex = I
                    If Cmb_Traslado.text = .rdoColumns("Traslado") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Tipo.ListCount - 1
                    Cmb_Tipo.ListIndex = I
                    If Cmb_Tipo.text = .rdoColumns("Tipo") Then
                        Exit For
                        End If
                Next
                If Not IsNull(.rdoColumns("Entidad")) Then
                    Txt_Entidad.text = .rdoColumns("Entidad")
                Else
                    Txt_Entidad.text = ""
                End If
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Impuestos.EOF
            Grid_Impuestos.AddItem Rs_Consulta_Cat_Impuestos.rdoColumns("Impuesto_ID") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Retencion") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Traslado") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Tipo") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Entidad") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Estado")
            Grid_Impuestos.FixedRows = 1
            Rs_Consulta_Cat_Impuestos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Impuestos.FixedCols = 1
        Grid_Impuestos.ColWidth(0) = 1200
        Grid_Impuestos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Impuestos.ColWidth(1) = 2000
        Grid_Impuestos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(2) = 1000
        Grid_Impuestos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(3) = 1000
        Grid_Impuestos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(4) = 1000
        Grid_Impuestos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(5) = 1500
        Grid_Impuestos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(6) = 1000
        Grid_Impuestos.ColAlignment(6) = flexAlignLeftCenter
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Impuestos.Close
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Impuestos As rdoResultset   'Manejo de registro
    
    Grid_Impuestos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Impuestos WHERE (Impuesto_ID ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Impuestos.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Impuestos.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Impuestos SET Estado='INACTIVO' WHERE (Impuesto_ID ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Impuestos SET Estado='ACTIVO' WHERE (Impuesto_ID ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Impuestos.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub
Public Sub Modifica_Impuestos()
Dim Rs_Modificacion_Cat_Impuestos As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Impuestos"
    Mi_SQL = Mi_SQL & " WHERE Impuesto_ID='" & Txt_Clave_Impuesto.text & "'"
    Set Rs_Modificacion_Cat_Impuestos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Impuestos.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Impuestos
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Impuesto.text)
                .rdoColumns("Entidad") = UCase(Txt_Entidad.text)
                .rdoColumns("Retencion") = UCase(Cmb_Retencion.text)
                .rdoColumns("Traslado") = UCase(Cmb_Traslado.text)
                .rdoColumns("Tipo") = UCase(Cmb_Tipo.text)
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Impuestos.Close
    MsgBox "El impuesto ha sido modificado", vbInformation
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
    If Txt_Clave_Impuesto.text <> "" Then
        Consulta_Impuestos (Txt_Clave_Impuesto.text)
        
    Else
        MsgBox "Ingrese el código del impuesto", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Clave_Impuesto <> "" Then
        Cambiar_Estado (Txt_Clave_Impuesto.text)
        
    Else
        MsgBox "Ingrese el código del impuesto", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Clave_Impuesto.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Impuesto.Enabled = True
                Txt_Entidad.Enabled = True
                Cmb_Retencion.Enabled = True
                Cmb_Traslado.Enabled = True
                Cmb_Tipo.Enabled = True
                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion_Impuesto.text) <> "" And Cmb_Retencion.ListIndex > -1 And Cmb_Traslado.ListIndex > -1 And Cmb_Tipo.ListIndex > -1 Then
                    Modifica_Impuestos
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
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
        Txt_Descripcion_Impuesto.Enabled = False
        Txt_Entidad.Enabled = False
        Cmb_Retencion.Enabled = False
        Cmb_Traslado.Enabled = False
        Cmb_Tipo.Enabled = False
                
        If Grid_Impuestos.Rows <> 0 Then
            Grid_Impuestos.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Grid_Impuestos_Click()
Dim Rs_Consulta_Cat_Impuestos As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Impuestos.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Impuesto.Enabled = True
        Txt_Entidad.Enabled = True
        Cmb_Retencion.Enabled = True
        Cmb_Traslado.Enabled = True
        Cmb_Tipo.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Impuestos"
        Mi_SQL = Mi_SQL & " WHERE Impuesto_ID='" & Grid_Impuestos.TextMatrix(Grid_Impuestos.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Impuestos.EOF Then
            With Rs_Consulta_Cat_Impuestos
                Txt_Clave_Impuesto.text = .rdoColumns("Impuesto_ID")
                Txt_Descripcion_Impuesto.text = .rdoColumns("Descripcion")
                For I = 0 To Cmb_Retencion.ListCount - 1
                    Cmb_Retencion.ListIndex = I
                    If Cmb_Retencion.text = .rdoColumns("Retencion") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Traslado.ListCount - 1
                    Cmb_Traslado.ListIndex = I
                    If Cmb_Traslado.text = .rdoColumns("Traslado") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Tipo.ListCount - 1
                    Cmb_Tipo.ListIndex = I
                    If Cmb_Tipo.text = .rdoColumns("Tipo") Then
                        Exit For
                        End If
                Next
                If Not IsNull(.rdoColumns("Entidad")) Then
                    Txt_Entidad.text = .rdoColumns("Entidad")
                Else
                    Txt_Entidad.text = ""
                End If
                              
            End With
        End If
        Rs_Consulta_Cat_Impuestos.Close
    End If
End Sub

Private Sub Txt_Clave_Impuesto_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Btn_Nuevo_Click()
 If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Impuestos.Rows <> 0 Then
                    Grid_Impuestos.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Impuesto.Enabled = True
                Txt_Entidad.Enabled = True
                Cmb_Retencion.Enabled = True
                Cmb_Traslado.Enabled = True
                Cmb_Tipo.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Clave_Impuesto.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Impuestos", "Impuesto_ID"), "000")
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Servicios", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Btn_Eliminar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Impuesto.text) <> "" And Cmb_Retencion.ListIndex > -1 And Cmb_Traslado.ListIndex > -1 And Cmb_Tipo.ListIndex > -1 Then
                Alta_Impuestos
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub

Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
End Sub


Public Sub Alta_Impuestos()
Dim Rs_Alta_Cat_Impuesto As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Impuesto_ID FROM Cat_Impuestos where Impuesto_ID ='" & Trim(Txt_Clave_Impuesto.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Impuesto que quiere dar de alta ya existe", vbInformation
        Txt_Clave_Impuesto.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Impuesto = Conectar_Ayudante.Recordset_Agregar("Cat_Impuestos")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Impuesto
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Impuestos", "Clave"), "00")
            
            .rdoColumns("Impuesto_ID") = Txt_Clave_Impuesto.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Impuesto.text)
            .rdoColumns("Retencion") = UCase(Cmb_Retencion.text)
            .rdoColumns("Traslado") = UCase(Cmb_Traslado.text)
            .rdoColumns("Traslado") = UCase(Cmb_Traslado.text)
            .rdoColumns("Tipo") = UCase(Cmb_Tipo.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Impuesto.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Impuestos.Rows = 0 Then
        Grid_Impuestos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Retención" & Chr(9) & "Traslado" & Chr(9) & "Tipo" & Chr(9) & "Entidad" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    Grid_Impuestos.AddItem UCase(Txt_Clave_Impuesto.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Impuesto.text)) & Chr(9) & UCase(Cmb_Retencion.text) & Chr(9) & UCase(Cmb_Traslado.text) & Chr(9) & UCase(Cmb_Tipo.text) & Chr(9) & UCase(Txt_Entidad.text) & Chr(9) & UCase("ACTIVO")
        Grid_Impuestos.FixedCols = 1
        Grid_Impuestos.ColWidth(0) = 1200
        Grid_Impuestos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Impuestos.ColWidth(1) = 2000
        Grid_Impuestos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(2) = 1000
        Grid_Impuestos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(3) = 1000
        Grid_Impuestos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(4) = 1000
        Grid_Impuestos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(5) = 1500
        Grid_Impuestos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(6) = 1000
        Grid_Impuestos.ColAlignment(6) = flexAlignLeftCenter
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Impuestos As rdoResultset   'Manejo de registro
    
    Grid_Impuestos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Impuestos"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Impuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Impuestos.EOF Then
        'Pone un encabezado en el grid
        Grid_Impuestos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Retención" & Chr(9) & "Traslado" & Chr(9) & "Tipo" & Chr(9) & "Entidad" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Impuestos.EOF
            Grid_Impuestos.AddItem Rs_Consulta_Cat_Impuestos.rdoColumns("Impuesto_ID") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Retencion") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Traslado") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Tipo") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Entidad") & Chr(9) & Rs_Consulta_Cat_Impuestos.rdoColumns("Estado")
            Grid_Impuestos.FixedRows = 1
            Rs_Consulta_Cat_Impuestos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Impuestos.FixedCols = 1
        Grid_Impuestos.ColWidth(0) = 1200
        Grid_Impuestos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Impuestos.ColWidth(1) = 2000
        Grid_Impuestos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(2) = 1000
        Grid_Impuestos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(3) = 1000
        Grid_Impuestos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(4) = 1000
        Grid_Impuestos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(5) = 1500
        Grid_Impuestos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Impuestos.ColWidth(6) = 1000
        Grid_Impuestos.ColAlignment(6) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Impuestos.Close
End Sub
