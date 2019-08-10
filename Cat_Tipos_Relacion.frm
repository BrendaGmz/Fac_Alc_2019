VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Tipos_Relacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo Tipos de Relación"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   360
      Picture         =   "Cat_Tipos_Relacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "A"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5475
      Picture         =   "Cat_Tipos_Relacion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "C"
      Top             =   4440
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3765
      Picture         =   "Cat_Tipos_Relacion.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "B"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7185
      Picture         =   "Cat_Tipos_Relacion.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2070
      Picture         =   "Cat_Tipos_Relacion.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "M"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Vigencia 
      Caption         =   "Fechas de vigencia"
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
      Height          =   1575
      Left            =   5400
      TabIndex        =   8
      Top             =   600
      Width           =   3255
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90505217
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90505217
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio de Vigencia"
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
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.TextBox Txt_Clave_Tipo_Relacion 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Tipo_Relacion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1410
      Width           =   2775
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
      Top             =   2160
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid Grid_Relaciones 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
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
      Top             =   600
      Width           =   5055
      Begin VB.Label Txt_Clave_Tipo_Relaciones 
         Caption         =   "Clave Tipo de Relación"
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
         Width           =   2295
      End
      Begin VB.Label Txt_Descripcion_Tipo_Relaciones 
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
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "TIPO DE RELACIÓN"
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
      TabIndex        =   6
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "Frm_Cat_Tipos_Relacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
    If Txt_Clave_Tipo_Relacion.text <> "" Then
        Btn_Salir.Caption = "Regresar"
        Consulta_Relaciones (Txt_Clave_Tipo_Relacion.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Consulta_Relaciones(Cadena As String)
Dim Rs_Consulta_Cat_Relaciones As rdoResultset   'Manejo de registro
        Grid_Relaciones.Rows = 0
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Tipo_Relacion.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipos_Relacion"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Tipo_Relacion ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Relaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Relaciones.EOF Then
        'Pone un encabezado en el grid
        Grid_Relaciones.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Relaciones
                    
                    Txt_Descripcion_Tipo_Relacion.text = .rdoColumns("Descripcion")
                    Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                    If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
        
              
        End With
        While Not Rs_Consulta_Cat_Relaciones.EOF
        Grid_Relaciones.AddItem Rs_Consulta_Cat_Relaciones.rdoColumns("Codigo_Tipo_Relacion") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Estado")
            Grid_Relaciones.FixedRows = 1
            Rs_Consulta_Cat_Relaciones.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Relaciones.FixedCols = 1
        Grid_Relaciones.ColWidth(0) = 1300
        Grid_Relaciones.ColAlignment(0) = flexAlignCenterCenter
        Grid_Relaciones.ColWidth(1) = 3300
        Grid_Relaciones.ColAlignment(1) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(2) = 1250
        Grid_Relaciones.ColAlignment(2) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(3) = 1200
        Grid_Relaciones.ColAlignment(3) = flexAlignLeftCenter
        
        
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Relaciones.Close
End Sub


Private Sub Btn_Eliminar_Click()
If Txt_Clave_Tipo_Relacion <> "" Then
       Cambiar_Estado (Txt_Clave_Tipo_Relacion.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Relaciones As rdoResultset   'Manejo de registro
    
    Grid_Relaciones.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Tipos_Relacion WHERE (Codigo_Tipo_Relacion ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Relaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Relaciones.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Relaciones.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Tipos_Relacion SET Estado='INACTIVO' WHERE (Codigo_Tipo_Relacion ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Tipos_Relacion SET Estado='ACTIVO' WHERE (Codigo_Tipo_Relacion ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Relaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Relaciones.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Clave_Tipo_Relacion.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Tipo_Relacion.Enabled = True
                Fra_Vigencia.Enabled = True
                                    
            Else
                MsgBox "Ingrese código", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Clave_Tipo_Relacion.text) <> "" And Trim(Txt_Descripcion_Tipo_Relacion.text) <> "" Then
                    Modifica_Relaciones
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub

Public Sub Modifica_Relaciones()
Dim Rs_Modificacion_Cat_Relaciones As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tipos_Relacion"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Relacion='" & Txt_Clave_Tipo_Relacion.text & "'"
    Set Rs_Modificacion_Cat_Relaciones = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Relaciones.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Relaciones
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Tipo_Relacion.text)
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
                If Chc_Vigencia.Value = 1 Then
                    .rdoColumns("Fecha_Fin_Vigencia") = Null
                 Else
                .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
                  End If
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Relaciones.Close
    MsgBox "El producto ha sido modificado", vbInformation
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
                If Grid_Relaciones.Rows <> 0 Then
                    Grid_Relaciones.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Tipo_Relacion.Enabled = True
                Fra_Vigencia.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Clave_Tipo_Relacion.text) <> "" And Trim(Txt_Descripcion_Tipo_Relacion.text) <> "" Then
                Alta_Relaciones
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Relaciones()
Dim Rs_Alta_Cat_Relaciones As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Tipo_Relacion FROM Cat_Tipos_Relacion where Codigo_Tipo_Relacion ='" & Trim(Txt_Clave_Tipo_Relacion.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Clave_Tipo_Relacion.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Relaciones = Conectar_Ayudante.Recordset_Agregar("Cat_Tipos_Relacion")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Relaciones
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Relacion", "Clave"), "00")
            .rdoColumns("Codigo_Tipo_Relacion") = Txt_Clave_Tipo_Relacion.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Tipo_Relacion.text)
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Relaciones.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Relaciones.Rows = 0 Then
        Grid_Relaciones.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Relaciones.AddItem UCase(Txt_Clave_Tipo_Relacion.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Tipo_Relacion.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Relaciones.AddItem UCase(Txt_Clave_Tipo_Relacion.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Tipo_Relacion.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Relaciones.FixedCols = 1
        Grid_Relaciones.ColWidth(0) = 1300
        Grid_Relaciones.ColAlignment(0) = flexAlignCenterCenter
        Grid_Relaciones.ColWidth(1) = 3300
        Grid_Relaciones.ColAlignment(1) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(2) = 1250
        Grid_Relaciones.ColAlignment(2) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(3) = 1200
        Grid_Relaciones.ColAlignment(3) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
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
        Txt_Descripcion_Tipo_Relacion.Enabled = False
        Fra_Vigencia.Enabled = False
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        If Grid_Relaciones.Rows <> 0 Then
            Grid_Relaciones.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Chc_Vigencia_Click()
    If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia.Enabled = False
    Else
        Dtp_Fin_Vigencia.Enabled = True
    End If
End Sub

Private Sub Grid_Relaciones_Click()
Dim Rs_Consulta_Cat_Relaciones As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Relaciones.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Tipo_Relacion.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Tipos_Relacion"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Relacion='" & Grid_Relaciones.TextMatrix(Grid_Relaciones.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Relaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Relaciones.EOF Then
            With Rs_Consulta_Cat_Relaciones
                Txt_Clave_Tipo_Relacion.text = .rdoColumns("Codigo_Tipo_Relacion")
                Txt_Descripcion_Tipo_Relacion.text = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                
                
            End With
        End If
        Rs_Consulta_Cat_Relaciones.Close
    End If
End Sub

Private Sub Txt_Clave_Tipo_Relacion_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Relaciones As rdoResultset   'Manejo de registro
    
    Grid_Relaciones.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipos_Relacion"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Relaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Relaciones.EOF Then
        'Pone un encabezado en el grid
        Grid_Relaciones.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Relaciones.EOF
            Grid_Relaciones.AddItem Rs_Consulta_Cat_Relaciones.rdoColumns("Codigo_Tipo_Relacion") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Relaciones.rdoColumns("Estado")
            Grid_Relaciones.FixedRows = 1
            Rs_Consulta_Cat_Relaciones.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Relaciones.FixedCols = 1
        Grid_Relaciones.ColWidth(0) = 1300
        Grid_Relaciones.ColAlignment(0) = flexAlignCenterCenter
        Grid_Relaciones.ColWidth(1) = 3300
        Grid_Relaciones.ColAlignment(1) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(2) = 1250
        Grid_Relaciones.ColAlignment(2) = flexAlignLeftCenter
        Grid_Relaciones.ColWidth(3) = 1200
        Grid_Relaciones.ColAlignment(3) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Relaciones.Close
End Sub

