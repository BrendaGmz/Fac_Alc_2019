VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Productos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Catalogo de Productos y  Servicios"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra_Vigencia 
      Caption         =   "Vigencia"
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
      Height          =   1455
      Left            =   4800
      TabIndex        =   20
      Top             =   600
      Width           =   3615
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia_Producto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   45744129
         CurrentDate     =   36892
         MinDate         =   -328716
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia_Producto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   45744129
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio de Vigencia"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1830
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "M"
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6945
      Picture         =   "Form1.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3525
      Picture         =   "Form1.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "B"
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5235
      Picture         =   "Form1.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "C"
      Top             =   5760
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Form1.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "A"
      Top             =   5760
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
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid Grid_ProdServ 
         Height          =   1695
         Left            =   120
         TabIndex        =   19
         Top             =   280
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame Fra_Comentarios 
      Caption         =   "Comentarios"
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
      Height          =   1335
      Left            =   2760
      TabIndex        =   10
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox Txt_Comentarios_Producto 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Fra_Impuestos 
      Caption         =   "Impuestos de Translado"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2535
      Begin VB.ComboBox Cmb_IEPS_Producto 
         Height          =   315
         ItemData        =   "Form1.frx":0552
         Left            =   600
         List            =   "Form1.frx":055F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox Cmb_IVA_Producto 
         Height          =   315
         ItemData        =   "Form1.frx":0575
         Left            =   600
         List            =   "Form1.frx":0582
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "IEPS"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "IVA"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.TextBox Txt_Descripcion_Producto 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Txt_Codigo_Producto 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1335
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
      TabIndex        =   2
      Top             =   600
      Width           =   4575
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código del Producto"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "PRODUCTOS Y SERVICIOS"
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
      TabIndex        =   13
      Top             =   120
      Width           =   8265
   End
End
Attribute VB_Name = "Frm_Cat_Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Clave As String

Private Sub Chc_Vigencia_Click()
    If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia_Producto.Enabled = False
    Else
        Dtp_Fin_Vigencia_Producto.Enabled = True
    End If
End Sub

Private Sub Txt_Codigo_Producto_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Productos As rdoResultset   'Manejo de registro
    
    Grid_ProdServ.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT Estado from Cat_Productos_Servicios WHERE (Clave_Producto_Servicio ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Productos.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Productos.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Productos_Servicios SET Estado='INACTIVO' WHERE (Clave_Producto_Servicio ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Productos_Servicios SET Estado='ACTIVO' WHERE (Clave_Producto_Servicio ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Productos.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Public Sub Consulta_Productos(Cadena As String)
Dim Rs_Consulta_Cat_Productos As rdoResultset   'Manejo de registro
    Txt_Descripcion_Producto.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_ProdServ.Rows = 0
    Fra_Impuestos.Enabled = True
    Fra_Comentarios.Enabled = True
    Fra_Vigencia.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Productos_Servicios"
    Mi_SQL = Mi_SQL & " WHERE (Clave_Producto_Servicio ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Productos_Servicios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Productos_Servicios.EOF Then
        'Pone un encabezado en el grid
        Grid_ProdServ.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Inicio Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Incluir IVA" & Chr(9) & "Incluir IEPS" & Chr(9) & "Complemento" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Productos_Servicios
                Txt_Codigo_Producto.text = .rdoColumns("Clave_Producto_Servicio")
                Txt_Descripcion_Producto.text = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia_Producto.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                Dtp_Fin_Vigencia_Producto.Value = .rdoColumns("Fecha_Fin_Vigencia")
                For I = 0 To Cmb_IVA_Producto.ListCount - 1
                Cmb_IVA_Producto.ListIndex = I
                    If Cmb_IVA_Producto.text = .rdoColumns("Incluir_IVA") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_IEPS_Producto.ListCount - 1
                Cmb_IEPS_Producto.ListIndex = I
                    If Cmb_IEPS_Producto.text = .rdoColumns("Incluir_IEPS") Then
                        Exit For
                        End If
                Next
                Txt_Comentarios_Producto.text = .rdoColumns("Complemento")
        End With
        'Llenado del grid
        While Not Rs_Consulta_Cat_Productos_Servicios.EOF
            Grid_ProdServ.AddItem Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Clave_Producto_Servicio") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Incluir_IVA") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Incluir_IEPS") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Complemento") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Estado")
            Grid_ProdServ.FixedRows = 1
            Rs_Consulta_Cat_Productos_Servicios.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_ProdServ.FixedCols = 1
        Grid_ProdServ.ColWidth(0) = 1200
        Grid_ProdServ.ColAlignment(0) = flexAlignCenterCenter
        Grid_ProdServ.ColWidth(1) = 3000
        Grid_ProdServ.ColAlignment(1) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(2) = 1200
        Grid_ProdServ.ColAlignment(2) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(3) = 1200
        Grid_ProdServ.ColAlignment(3) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(4) = 1000
        Grid_ProdServ.ColAlignment(4) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(5) = 1000
        Grid_ProdServ.ColAlignment(5) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(6) = 3000
        Grid_ProdServ.ColAlignment(6) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(7) = 1200
        Grid_ProdServ.ColAlignment(7) = flexAlignLeftCenter
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Productos_Servicios.Close
End Sub


Public Sub Modifica_Productos()
Dim Rs_Modificacion_Cat_Productos As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Productos_Servicios"
    Mi_SQL = Mi_SQL & " WHERE Clave_Producto_Servicio='" & Txt_Codigo_Producto.text & "'"
    Set Rs_Modificacion_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Productos.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Productos
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Producto.text)
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia_Producto.Value, "yyyy/MM/dd")
                
                If Chc_Vigencia.Value = 1 Then
                    .rdoColumns("Fecha_Fin_Vigencia") = Null
                 Else
                .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia_Producto.Value, "yyyy/MM/dd")
                  End If
                .rdoColumns("Incluir_IVA") = UCase(Cmb_IVA_Producto.text)
                .rdoColumns("Incluir_IEPS") = UCase(Cmb_IEPS_Producto.text)
                .rdoColumns("Complemento") = UCase(Txt_Comentarios_Producto.text)
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Productos.Close
    MsgBox "El producto ha sido modificado", vbInformation
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


Public Sub Alta_Productos()
Dim Rs_Alta_Cat_Producto As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Clave_Producto_Servicio FROM Cat_Productos_Servicios where Clave_Producto_Servicio ='" & Trim(Txt_Codigo_Producto.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Producto.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Producto = Conectar_Ayudante.Recordset_Agregar("Cat_Productos_Servicios")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Producto
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Servicios", "Clave"), "00")
            .rdoColumns("Clave_Producto_Servicio") = Txt_Codigo_Producto.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Producto.text)
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia_Producto.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia_Producto.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Incluir_IVA") = UCase(Cmb_IVA_Producto.text)
            .rdoColumns("Incluir_IEPS") = UCase(Cmb_IEPS_Producto.text)
            .rdoColumns("Complemento") = UCase(Txt_Comentarios_Producto.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Producto.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_ProdServ.Rows = 0 Then
        Grid_ProdServ.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Inicio Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Incluir IVA" & Chr(9) & "Incluir IEPS" & Chr(9) & "Complemento" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_ProdServ.AddItem UCase(Txt_Codigo_Producto.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Producto.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia_Producto.Value) & Chr(9) & UCase("") & Chr(9) & UCase(Cmb_IVA_Producto.text) & Chr(9) & UCase(Cmb_IEPS_Producto.text) & Chr(9) & UCase(Txt_Comentarios_Producto.text) & Chr(9) & UCase("ACTIVO")
            Else
            Grid_ProdServ.AddItem UCase(Txt_Codigo_Producto.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Producto.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia_Producto.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia_Producto.Value) & Chr(9) & UCase(Cmb_IVA_Producto.text) & Chr(9) & UCase(Cmb_IEPS_Producto.text) & Chr(9) & UCase(Txt_Comentarios_Producto.text) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_ProdServ.FixedCols = 1
        Grid_ProdServ.ColWidth(0) = 1200
        Grid_ProdServ.ColAlignment(0) = flexAlignCenterCenter
        Grid_ProdServ.ColWidth(1) = 3000
        Grid_ProdServ.ColAlignment(1) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(2) = 1200
        Grid_ProdServ.ColAlignment(2) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(3) = 1200
        Grid_ProdServ.ColAlignment(3) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(4) = 1000
        Grid_ProdServ.ColAlignment(4) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(5) = 1000
        Grid_ProdServ.ColAlignment(5) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(6) = 3000
        Grid_ProdServ.ColAlignment(6) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(7) = 1200
        Grid_ProdServ.ColAlignment(7) = flexAlignLeftCenter
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
    If Txt_Codigo_Producto.text <> "" Then
        Consulta_Productos (Txt_Codigo_Producto.text)
        
    Else
        MsgBox "Ingrese el código del producto", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Producto <> "" Then
        Cambiar_Estado (Txt_Codigo_Producto.text)
        
    Else
        MsgBox "Ingrese el código del producto", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
 If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Producto.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Producto.Enabled = True
                Fra_Vigencia.Enabled = True
                Fra_Impuestos.Enabled = True
                Fra_Comentarios.Enabled = True
                Txt_Codigo_Producto.Enabled = False
                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Codigo_Producto.text) <> "" And Cmb_IVA_Producto.ListIndex > -1 And Cmb_IEPS_Producto.ListIndex > -1 Then
                    Modifica_Productos
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_ProdServ.Rows <> 0 Then
                    Grid_ProdServ.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Producto.Enabled = True
                Fra_Vigencia.Enabled = True
                Fra_Impuestos.Enabled = True
                Fra_Comentarios.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Codigo_Producto.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Servicios", "Clave_Producto_Servicio"), "00000000")
                Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Productos_Servicios", "Clave"), "00")
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Producto.text) <> "" And Cmb_IVA_Producto.ListIndex > -1 And Cmb_IEPS_Producto.ListIndex > -1 Then
                Alta_Productos
                
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
        
        Txt_Descripcion_Producto.Enabled = False
        Fra_Vigencia.Enabled = False
        Fra_Impuestos.Enabled = False
        Fra_Comentarios.Enabled = False
        Dtp_Inicio_Vigencia_Producto.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia_Producto.Value = Format(Now(), "yyyy/MMM/dd")
                
        If Grid_ProdServ.Rows <> 0 Then
            Grid_ProdServ.Rows = 0
            End If
        Consulta
    End If

End Sub

Private Sub Form_Load()
    Set Conexion = New Conectar
    Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia_Producto.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia_Producto.Value = Format(Now, "yyyy/MMM/dd")

End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Productos_Servicios As rdoResultset   'Manejo de registro
    
    Grid_ProdServ.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Productos_Servicios"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Productos_Servicios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Productos_Servicios.EOF Then
        'Pone un encabezado en el grid
        Grid_ProdServ.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Inicio Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Incluir IVA" & Chr(9) & "Incluir IEPS" & Chr(9) & "Complemento" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Productos_Servicios.EOF
            Grid_ProdServ.AddItem Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Clave_Producto_Servicio") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Incluir_IVA") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Incluir_IEPS") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Complemento") & Chr(9) & Rs_Consulta_Cat_Productos_Servicios.rdoColumns("Estado")
            Grid_ProdServ.FixedRows = 1
            Rs_Consulta_Cat_Productos_Servicios.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_ProdServ.FixedCols = 1
        Grid_ProdServ.ColWidth(0) = 1200
        Grid_ProdServ.ColAlignment(0) = flexAlignCenterCenter
        Grid_ProdServ.ColWidth(1) = 3000
        Grid_ProdServ.ColAlignment(1) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(2) = 1200
        Grid_ProdServ.ColAlignment(2) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(3) = 1200
        Grid_ProdServ.ColAlignment(3) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(4) = 1000
        Grid_ProdServ.ColAlignment(4) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(5) = 1000
        Grid_ProdServ.ColAlignment(5) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(6) = 3000
        Grid_ProdServ.ColAlignment(6) = flexAlignLeftCenter
        Grid_ProdServ.ColWidth(7) = 1200
        Grid_ProdServ.ColAlignment(7) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Productos_Servicios.Close
End Sub

Private Sub Grid_ProdServ_Click()
Dim Rs_Consulta_Cat_Productos As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_ProdServ.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Producto.Enabled = True
        Fra_Impuestos.Enabled = True
        Fra_Comentarios.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Productos_Servicios"
        Mi_SQL = Mi_SQL & " WHERE Clave_Producto_Servicio='" & Grid_ProdServ.TextMatrix(Grid_ProdServ.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Productos.EOF Then
            With Rs_Consulta_Cat_Productos
                Txt_Codigo_Producto.text = .rdoColumns("Clave_Producto_Servicio")
                Txt_Descripcion_Producto.text = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia_Producto.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia_Producto.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                
                For I = 0 To Cmb_IVA_Producto.ListCount - 1
                Cmb_IVA_Producto.ListIndex = I
                    If Cmb_IVA_Producto.text = .rdoColumns("Incluir_IVA") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_IEPS_Producto.ListCount - 1
                Cmb_IEPS_Producto.ListIndex = I
                    If Cmb_IEPS_Producto.text = .rdoColumns("Incluir_IEPS") Then
                        Exit For
                        End If
                Next
                Txt_Comentarios_Producto.text = .rdoColumns("Complemento")
                
            End With
        End If
        Rs_Consulta_Cat_Productos.Close
    End If
End Sub
