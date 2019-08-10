VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Uso_Comprobantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Uso de Comprobantes"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1560
      Picture         =   "Cat_Uso_Comprobantes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "M"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5160
      Picture         =   "Cat_Uso_Comprobantes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2760
      Picture         =   "Cat_Uso_Comprobantes.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "B"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   3960
      Picture         =   "Cat_Uso_Comprobantes.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "C"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   360
      Picture         =   "Cat_Uso_Comprobantes.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "A"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
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
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   3495
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92995585
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92995585
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
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Fra_Impuestos 
      Caption         =   "Aplica para tipo de persona:"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
      Begin VB.ComboBox Cmb_Persona_Moral 
         Height          =   315
         ItemData        =   "Cat_Uso_Comprobantes.frx":0552
         Left            =   1080
         List            =   "Cat_Uso_Comprobantes.frx":055C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Persona_Fisica 
         Height          =   315
         ItemData        =   "Cat_Uso_Comprobantes.frx":0568
         Left            =   1080
         List            =   "Cat_Uso_Comprobantes.frx":0572
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Moral"
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
         Index           =   5
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Física"
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
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.TextBox Txt_Codigo_Comprobante 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txt_Descripcion_Comprobante 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2760
      MultiLine       =   -1  'True
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid Grid_Usos 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6135
      Begin VB.Label Label2 
         Caption         =   "Código uso de comprobante"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "USO DE COMPROBANTES"
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
      Width           =   6105
   End
End
Attribute VB_Name = "Frm_Cat_Uso_Comprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Comprobante.text <> "" Then
        Consulta_Usos (Txt_Codigo_Comprobante.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Consulta_Usos(Cadena As String)
Dim Rs_Consulta_Cat_Usos As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Usos.Rows = 0
    Txt_Descripcion_Comprobante.Enabled = True
    Fra_Impuestos.Enabled = True
    Fra_Vigencia.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Uso_Comprobantes"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Uso_Comprobante ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Usos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Usos.EOF Then
        'Pone un encabezado en el grid
        Grid_Usos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Usos.EOF
            Grid_Usos.AddItem Rs_Consulta_Cat_Usos.rdoColumns("Codigo_Uso_Comprobante") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Persona_Fisica") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Persona_Moral") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Estado")
            Grid_Usos.FixedRows = 1
            Rs_Consulta_Cat_Usos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Usos.FixedCols = 1
        Grid_Usos.ColWidth(0) = 1300
        Grid_Usos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Usos.ColWidth(1) = 3300
        Grid_Usos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Usos.ColWidth(2) = 1250
        Grid_Usos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Usos.ColWidth(3) = 1200
        Grid_Usos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Usos.ColWidth(4) = 1200
        Grid_Usos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Usos.ColWidth(5) = 1200
        Grid_Usos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Usos.ColWidth(6) = 1200
        Grid_Usos.ColAlignment(6) = flexAlignLeftCenter
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Usos.Close
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Comprobante <> "" Then
       Cambiar_Estado (Txt_Codigo_Comprobante.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Comprobantes As rdoResultset   'Manejo de registro
    
    Grid_Usos.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Uso_Comprobantes WHERE (Codigo_Uso_Comprobante ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Uso_Comprobantes SET Estado='INACTIVO' WHERE (Codigo_Uso_Comprobante ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Uso_Comprobantes SET Estado='ACTIVO' WHERE (Codigo_Uso_Comprobante ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Comprobantes.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Comprobante.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion_Comprobante.text) <> "" And Cmb_Persona_Fisica.ListIndex > -1 And Cmb_Persona_Moral.ListIndex > -1 Then
                    Modifica_Comprobantes
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Comprobantes()
Dim Rs_Modificacion_Cat_Comprobantes As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Uso_Comprobantes"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Uso_Comprobante='" & Txt_Codigo_Comprobante.text & "'"
    Set Rs_Modificacion_Cat_Comprobantes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Comprobantes.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Comprobantes
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Comprobante.text)
                .rdoColumns("Persona_Fisica") = UCase(Cmb_Persona_Fisica.text)
                .rdoColumns("Persona_Moral") = UCase(Cmb_Persona_Fisica.text)
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
    Rs_Modificacion_Cat_Comprobantes.Close
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
                If Grid_Usos.Rows <> 0 Then
                    Grid_Usos.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Fra_Impuestos.Enabled = True
                Fra_Vigencia.Enabled = True
                Txt_Descripcion_Comprobante.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion_Comprobante.text) <> "" And Trim(Txt_Codigo_Comprobante.text) <> "" And Cmb_Persona_Fisica.ListIndex > -1 And Cmb_Persona_Moral.ListIndex > -1 Then
                Alta_Usos
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Usos()
Dim Rs_Alta_Cat_Usos As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Uso_Comprobante FROM Cat_Uso_Comprobantes where Codigo_Uso_Comprobante ='" & Trim(Txt_Codigo_Comprobante.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Comprobante.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Usos = Conectar_Ayudante.Recordset_Agregar("Cat_Uso_Comprobantes")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Usos
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Uso_Comprobantes", "Clave"), "00")
            .rdoColumns("Codigo_Uso_Comprobante") = Txt_Codigo_Comprobante.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Comprobante.text)
            .rdoColumns("Persona_Fisica") = UCase(Cmb_Persona_Fisica.text)
            .rdoColumns("Persona_Moral") = UCase(Cmb_Persona_Moral.text)
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
    Rs_Alta_Cat_Usos.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Usos.Rows = 0 Then
        Grid_Usos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Usos.AddItem UCase(Txt_Codigo_Comprobante.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Comprobante.text)) & Chr(9) & UCase(Cmb_Persona_Fisica.text) & Chr(9) & UCase(Cmb_Persona_Moral.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Usos.AddItem UCase(Txt_Codigo_Comprobante.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Comprobante.text)) & Chr(9) & UCase(Cmb_Persona_Fisica.text) & Chr(9) & UCase(Cmb_Persona_Moral.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Usos.FixedCols = 1
        Grid_Usos.ColWidth(0) = 1300
        Grid_Usos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Usos.ColWidth(1) = 3300
        Grid_Usos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Usos.ColWidth(2) = 1250
        Grid_Usos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Usos.ColWidth(3) = 1200
        Grid_Usos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Usos.ColWidth(4) = 1200
        Grid_Usos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Usos.ColWidth(5) = 1200
        Grid_Usos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Usos.ColWidth(6) = 1200
        Grid_Usos.ColAlignment(6) = flexAlignLeftCenter
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
        Txt_Descripcion_Comprobante.Enabled = False
        Fra_Impuestos.Enabled = False
        Fra_Vigencia.Enabled = False
        Dtp_Fin_Vigencia.MinDate = Format(Now(), "yyyy/MMM/dd")
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Chc_Vigencia.Value = 1
        Chc_Vigencia_Click
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
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Usos As rdoResultset   'Manejo de registro
    
    Grid_Usos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Uso_Comprobantes ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Usos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Usos.EOF Then
        'Pone un encabezado en el grid
        Grid_Usos.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Usos.EOF
            Grid_Usos.AddItem Rs_Consulta_Cat_Usos.rdoColumns("Codigo_Uso_Comprobante") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Persona_Fisica") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Persona_Moral") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Usos.rdoColumns("Estado")
            Grid_Usos.FixedRows = 1
            Rs_Consulta_Cat_Usos.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Usos.FixedCols = 1
        Grid_Usos.ColWidth(0) = 1300
        Grid_Usos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Usos.ColWidth(1) = 3300
        Grid_Usos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Usos.ColWidth(2) = 1250
        Grid_Usos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Usos.ColWidth(3) = 1200
        Grid_Usos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Usos.ColWidth(4) = 1200
        Grid_Usos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Usos.ColWidth(5) = 1200
        Grid_Usos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Usos.ColWidth(6) = 1200
        Grid_Usos.ColAlignment(6) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Usos.Close
End Sub

Private Sub Grid_Usos_Click()
Dim Rs_Consulta_Cat_Usos As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Usos.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion_Comprobante.Enabled = True
        Fra_Impuestos.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Uso_Comprobantes"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Uso_Comprobante='" & Grid_Usos.TextMatrix(Grid_Usos.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Usos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Usos.EOF Then
            With Rs_Consulta_Cat_Usos
                Txt_Codigo_Comprobante.text = .rdoColumns("Codigo_Uso_Comprobante")
                Txt_Descripcion_Comprobante.text = .rdoColumns("Descripcion")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                
                For I = 0 To Cmb_Persona_Fisica.ListCount - 1
                Cmb_Persona_Fisica.ListIndex = I
                    If Cmb_Persona_Fisica.text = .rdoColumns("Persona_Fisica") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Persona_Moral.ListCount - 1
                Cmb_Persona_Moral.ListIndex = I
                    If Cmb_Persona_Moral.text = .rdoColumns("Persona_Moral") Then
                        Exit For
                        End If
                Next
                
            End With
        End If
        Rs_Consulta_Cat_Usos.Close
    End If
End Sub

Private Sub Txt_Codigo_Comprobante_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
