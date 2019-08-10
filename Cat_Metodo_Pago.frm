VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Cat_Metodo_Pago 
   Caption         =   "Catalogo de Métodos de Pago"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   240
      Picture         =   "Cat_Metodo_Pago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "A"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5355
      Picture         =   "Cat_Metodo_Pago.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "C"
      Top             =   4320
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3645
      Picture         =   "Cat_Metodo_Pago.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "B"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7065
      Picture         =   "Cat_Metodo_Pago.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1950
      Picture         =   "Cat_Metodo_Pago.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "M"
      Top             =   4320
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
      Left            =   4920
      TabIndex        =   7
      Top             =   480
      Width           =   3495
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   150274049
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   150274049
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   1935
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
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox Txt_Codigo_Metodo_Pago 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Metodo_Pago 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
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
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid Grid_Metodos 
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame Fra_Generales 
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4695
      Begin VB.Label Label2 
         Caption         =   "Código del método de pago"
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
         Width           =   2535
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
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MÉTODOS DE PAGO"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "Frm_Cat_Metodo_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Alta_Metodos()
Dim Rs_Alta_Cat_Metodos As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Metodo_ID FROM Cat_Metodo_Pago where Metodo_ID='" & Trim(Txt_Codigo_Metodo_Pago.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El mpetodo de pago que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Producto.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Metodos = Conectar_Ayudante.Recordset_Agregar("Cat_Metodo_Pago")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Metodos
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Unidades_Medida", "Clave"), "00")
            .rdoColumns("Metodo_ID") = Txt_Codigo_Metodo_Pago.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Metodo_Pago.text)
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
    Rs_Alta_Cat_Metodos.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Metodos.Rows = 0 Then
        Grid_Metodos.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Metodos.AddItem UCase(Txt_Codigo_Metodo_Pago.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Metodo_Pago.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Metodos.AddItem UCase(Txt_Codigo_Metodo_Pago.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Metodo_Pago.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Metodos.FixedCols = 1
        Grid_Metodos.ColWidth(0) = 1000
        Grid_Metodos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Metodos.ColWidth(1) = 2800
        Grid_Metodos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(2) = 1500
        Grid_Metodos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(3) = 1500
        Grid_Metodos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(4) = 1000
        Grid_Metodos.ColAlignment(4) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Consulta_Metodos(Cadena As String)
Dim Rs_Consulta_Cat_Metodos As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Metodos.Rows = 0
    Fra_Vigencia.Enabled = True
    Txt_Descripcion_Metodo_Pago.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Metodo_Pago"
    Mi_SQL = Mi_SQL & " WHERE (Metodo_ID ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Metodos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Metodos.EOF Then
        'Pone un encabezado en el grid
        Grid_Metodos.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Metodos
               Txt_Codigo_Metodo_Pago = .rdoColumns("Metodo_ID")
                Txt_Descripcion_Metodo_Pago = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    
                    End If
                
            End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Metodos.EOF
            Grid_Metodos.AddItem Rs_Consulta_Cat_Metodos.rdoColumns("Metodo_ID") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Estado")
            Grid_Metodos.FixedRows = 1
            Rs_Consulta_Cat_Metodos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Metodos.FixedCols = 1
        Grid_Metodos.ColWidth(0) = 1000
        Grid_Metodos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Metodos.ColWidth(1) = 2800
        Grid_Metodos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(2) = 1500
        Grid_Metodos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(3) = 1500
        Grid_Metodos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(4) = 1000
        Grid_Metodos.ColAlignment(4) = flexAlignLeftCenter
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Metodos.Close
End Sub
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Metodo_Pago.text <> "" Then
        Consulta_Metodos (Txt_Codigo_Metodo_Pago.text)
        
    Else
        MsgBox "Ingrese el código del método de pago", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Metodo_Pago <> "" Then
        Cambiar_Estado (Txt_Codigo_Metodo_Pago.text)
        
    Else
        MsgBox "Ingrese el código del método de pago", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Metodos As rdoResultset   'Manejo de registro
    
    Grid_Metodos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Metodo_Pago WHERE (Metodo_ID ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Metodos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Metodos.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Metodos.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Metodo_Pago SET Estado='INACTIVO' WHERE (Metodo_ID ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Metodo_Pago SET Estado='ACTIVO' WHERE (Metodo_ID ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Metodos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Metodos.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Metodo_Pago.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Fra_Vigencia.Enabled = True
                Txt_Descripcion_Metodo_Pago.Enabled = True
                                    
            Else
                MsgBox "Ingrese código del método de pago", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion_Metodo_Pago.text) <> "" And Trim(Txt_Codigo_Metodo_Pago.text) <> "" Then
                    Modifica_Metodos
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Metodos()
Dim Rs_Modificacion_Cat_Metodos As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Metodo_Pago"
    Mi_SQL = Mi_SQL & " WHERE Metodo_ID='" & Txt_Codigo_Metodo_Pago.text & "'"
    Set Rs_Modificacion_Cat_Metodos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Metodos.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Metodos
            .Edit
                
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Metodo_Pago.text)
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
    Rs_Modificacion_Cat_Metodos.Close
    MsgBox "El método de pago ha sido modificado", vbInformation
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
                If Grid_Metodos.Rows <> 0 Then
                    Grid_Metodos.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Fra_Vigencia.Enabled = True
                Txt_Descripcion_Metodo_Pago.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
                Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Metodo_Pago.text) <> "" And Trim(Txt_Codigo_Metodo_Pago.text) <> "" Then
                
                Alta_Metodos
                
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
        Fra_Vigencia.Enabled = False
        Txt_Descripcion_Metodo_Pago.Enabled = False
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        If Grid_Metodos.Rows <> 0 Then
            Grid_Metodos.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Chc_Vigencia_Click()
    If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia.Enabled = False
    Else
        Dtp_Fin_Vigencia.Enabled = True
        Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
        Dtp_Fin_Vigencia.Value = Dtp_Inicio_Vigencia.Value
    End If
End Sub



Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
End Sub

Public Sub Consulta()
Dim Rs_Consulta_Cat_Metodos As rdoResultset   'Manejo de registro
    
    Grid_Metodos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Metodo_Pago"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Metodos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Metodos.EOF Then
        'Pone un encabezado en el grid
        Grid_Metodos.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Metodos.EOF
            Grid_Metodos.AddItem Rs_Consulta_Cat_Metodos.rdoColumns("Metodo_ID") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Metodos.rdoColumns("Estado")
            Grid_Metodos.FixedRows = 1
            Rs_Consulta_Cat_Metodos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Metodos.FixedCols = 1
        Grid_Metodos.ColWidth(0) = 1000
        Grid_Metodos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Metodos.ColWidth(1) = 2800
        Grid_Metodos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(2) = 1500
        Grid_Metodos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(3) = 1500
        Grid_Metodos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Metodos.ColWidth(4) = 1000
        Grid_Metodos.ColAlignment(4) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Metodos.Close
End Sub

Private Sub Grid_Metodos_Click()
Dim Rs_Consulta_Cat_Metodos As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Metodos.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Fra_Vigencia.Enabled = True
        Txt_Descripcion_Metodo_Pago.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Metodo_Pago"
        Mi_SQL = Mi_SQL & " WHERE Metodo_ID='" & Grid_Metodos.TextMatrix(Grid_Metodos.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Metodos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Metodos.EOF Then
            With Rs_Consulta_Cat_Metodos
                Txt_Codigo_Metodo_Pago = .rdoColumns("Metodo_ID")
                Txt_Descripcion_Metodo_Pago = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    
                    End If
                
            End With
        End If
        Rs_Consulta_Cat_Metodos.Close
    End If
End Sub

Private Sub Txt_Codigo_Metodo_Pago_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
