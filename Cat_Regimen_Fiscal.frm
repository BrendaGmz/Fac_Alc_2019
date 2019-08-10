VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Regimen_Fiscal 
   Caption         =   "Catalogo Regímenes Fiscales"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
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
      Left            =   360
      TabIndex        =   20
      Top             =   3720
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid Grid_Fiscales 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
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
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   600
      Picture         =   "Cat_Regimen_Fiscal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "A"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4200
      Picture         =   "Cat_Regimen_Fiscal.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "C"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3000
      Picture         =   "Cat_Regimen_Fiscal.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "B"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5400
      Picture         =   "Cat_Regimen_Fiscal.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1800
      Picture         =   "Cat_Regimen_Fiscal.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "M"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   975
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
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Width           =   6135
      Begin VB.TextBox Txt_Descripcion 
         Enabled         =   0   'False
         Height          =   525
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox Txt_Codigo_Regimen 
         Height          =   285
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   1335
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
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Código regimen fiscal"
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
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   2535
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
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
      Begin VB.ComboBox Cmb_Persona_Fisica 
         Height          =   315
         ItemData        =   "Cat_Regimen_Fiscal.frx":0552
         Left            =   1080
         List            =   "Cat_Regimen_Fiscal.frx":055C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Persona_Moral 
         Height          =   315
         ItemData        =   "Cat_Regimen_Fiscal.frx":0568
         Left            =   1080
         List            =   "Cat_Regimen_Fiscal.frx":0572
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   1215
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
         TabIndex        =   11
         Top             =   480
         Width           =   615
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92995585
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "REGIMEN FISCAL"
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
      TabIndex        =   0
      Top             =   0
      Width           =   6225
   End
End
Attribute VB_Name = "Frm_Cat_Regimen_Fiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Regimen <> "" Then
       Cambiar_Estado (Txt_Codigo_Regimen.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Fiscales As rdoResultset   'Manejo de registro
    
    Grid_Fiscales.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Regimen_Fiscal WHERE (Codigo_Regimen_Fiscal ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Fiscales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Fiscales.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Fiscales.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Regimen_Fiscal SET Estado='INACTIVO' WHERE (Codigo_Regimen_Fiscal ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Regimen_Fiscal SET Estado='ACTIVO' WHERE (Codigo_Regimen_Fiscal ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Fiscales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Fiscales.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Regimen.text <> "" Then
        Consulta_Fiscales (Txt_Codigo_Regimen.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Consulta_Fiscales(Cadena As String)
Dim Rs_Consulta_Cat_Fiscales As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Grid_Fiscales.Rows = 0
    Txt_Descripcion.Enabled = True
    Fra_Impuestos.Enabled = True
    Fra_Vigencia.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Regimen_Fiscal"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Regimen_Fiscal ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Fiscales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Fiscales.EOF Then
        'Pone un encabezado en el grid
        Grid_Fiscales.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Fiscales.EOF
            Grid_Fiscales.AddItem Rs_Consulta_Cat_Fiscales.rdoColumns("Codigo_Regimen_Fiscal") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Persona_Fisica") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Persona_Moral") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Estado")
            Grid_Fiscales.FixedRows = 1
            Rs_Consulta_Cat_Fiscales.MoveNext
        Wend
        'Tamaño de las columnas en el grid
         Grid_Fiscales.FixedCols = 1
        Grid_Fiscales.ColWidth(0) = 1300
        Grid_Fiscales.ColAlignment(0) = flexAlignCenterCenter
        Grid_Fiscales.ColWidth(1) = 3300
        Grid_Fiscales.ColAlignment(1) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(2) = 1250
        Grid_Fiscales.ColAlignment(2) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(3) = 1200
        Grid_Fiscales.ColAlignment(3) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(4) = 1200
        Grid_Fiscales.ColAlignment(4) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(5) = 1200
        Grid_Fiscales.ColAlignment(5) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(6) = 1200
        Grid_Fiscales.ColAlignment(6) = flexAlignLeftCenter
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Fiscales.Close
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Regimen.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Descripcion.text) <> "" And Cmb_Persona_Fisica.ListIndex > -1 And Cmb_Persona_Moral.ListIndex > -1 Then
                    Modifica_Fiscales
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Fiscales()
Dim Rs_Modificacion_Cat_Fiscales As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Regimen_Fiscal"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Regimen_Fiscal='" & Txt_Codigo_Regimen.text & "'"
    Set Rs_Modificacion_Cat_Fiscales = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Fiscales.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Fiscales
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion.text)
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
    Rs_Modificacion_Cat_Fiscales.Close
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
                If Grid_Fiscales.Rows <> 0 Then
                    Grid_Fiscales.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Fra_Impuestos.Enabled = True
                Fra_Vigencia.Enabled = True
                Txt_Descripcion.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Descripcion.text) <> "" And Trim(Txt_Codigo_Regimen.text) <> "" And Cmb_Persona_Fisica.ListIndex > -1 And Cmb_Persona_Moral.ListIndex > -1 Then
                Alta_Fiscales
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Fiscales()
Dim Rs_Alta_Cat_Fiscales As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Regimen_Fiscal FROM Cat_Regimen_Fiscal where Codigo_Regimen_Fiscal ='" & Trim(Txt_Codigo_Regimen.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Regimen.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Fiscales = Conectar_Ayudante.Recordset_Agregar("Cat_Regimen_Fiscal")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Fiscales
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Regimen_Fiscal", "Clave"), "00")
            .rdoColumns("Codigo_Regimen_Fiscal") = Txt_Codigo_Regimen.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion.text)
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
    Rs_Alta_Cat_Fiscales.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Fiscales.Rows = 0 Then
        Grid_Fiscales.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Fiscales.AddItem UCase(Txt_Codigo_Regimen.text) & Chr(9) & UCase(Trim(Txt_Descripcion.text)) & Chr(9) & UCase(Cmb_Persona_Fisica.text) & Chr(9) & UCase(Cmb_Persona_Moral.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Fiscales.AddItem UCase(Txt_Codigo_Regimen.text) & Chr(9) & UCase(Trim(Txt_Descripcion.text)) & Chr(9) & UCase(Cmb_Persona_Fisica.text) & Chr(9) & UCase(Cmb_Persona_Moral.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Fiscales.FixedCols = 1
        Grid_Fiscales.ColWidth(0) = 1300
        Grid_Fiscales.ColAlignment(0) = flexAlignCenterCenter
        Grid_Fiscales.ColWidth(1) = 3300
        Grid_Fiscales.ColAlignment(1) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(2) = 1250
        Grid_Fiscales.ColAlignment(2) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(3) = 1200
        Grid_Fiscales.ColAlignment(3) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(4) = 1200
        Grid_Fiscales.ColAlignment(4) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(5) = 1200
        Grid_Fiscales.ColAlignment(5) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(6) = 1200
        Grid_Fiscales.ColAlignment(6) = flexAlignLeftCenter
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
        Txt_Descripcion.Enabled = False
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
   ' Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub

Public Sub Consulta()
Dim Rs_Consulta_Cat_Fiscales As rdoResultset   'Manejo de registro
    
    Grid_Fiscales.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Regimen_Fiscal ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Fiscales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Fiscales.EOF Then
        'Pone un encabezado en el grid
        Grid_Fiscales.AddItem "Clave" & Chr(9) & "Descripcion" & Chr(9) & "Persona Física" & Chr(9) & "Persona Moral" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Fiscales.EOF
            Grid_Fiscales.AddItem Rs_Consulta_Cat_Fiscales.rdoColumns("Codigo_Regimen_Fiscal") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Persona_Fisica") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Persona_Moral") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Fiscales.rdoColumns("Estado")
            Grid_Fiscales.FixedRows = 1
            Rs_Consulta_Cat_Fiscales.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Fiscales.FixedCols = 1
        Grid_Fiscales.ColWidth(0) = 1300
        Grid_Fiscales.ColAlignment(0) = flexAlignCenterCenter
        Grid_Fiscales.ColWidth(1) = 3300
        Grid_Fiscales.ColAlignment(1) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(2) = 1250
        Grid_Fiscales.ColAlignment(2) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(3) = 1200
        Grid_Fiscales.ColAlignment(3) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(4) = 1200
        Grid_Fiscales.ColAlignment(4) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(5) = 1200
        Grid_Fiscales.ColAlignment(5) = flexAlignLeftCenter
        Grid_Fiscales.ColWidth(6) = 1200
        Grid_Fiscales.ColAlignment(6) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Fiscales.Close
End Sub

Private Sub Grid_Fiscales_Click()
Dim Rs_Consulta_Cat_Fiscales As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Fiscales.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Descripcion.Enabled = True
        Fra_Impuestos.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Regimen_Fiscal"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Regimen_Fiscal='" & Grid_Fiscales.TextMatrix(Grid_Fiscales.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Fiscales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Fiscales.EOF Then
            With Rs_Consulta_Cat_Fiscales
                Txt_Codigo_Regimen.text = .rdoColumns("Codigo_Regimen_Fiscal")
                Txt_Descripcion.text = .rdoColumns("Descripcion")
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
        Rs_Consulta_Cat_Fiscales.Close
    End If
End Sub

Private Sub Txt_Codigo_Regimen_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
