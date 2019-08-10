VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Cat_Unidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo Unidades de Medida"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   240
      Picture         =   "Facturasm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "A"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5355
      Picture         =   "Facturasm.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "C"
      Top             =   6000
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3645
      Picture         =   "Facturasm.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "B"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7065
      Picture         =   "Facturasm.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1950
      Picture         =   "Facturasm.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "M"
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Vigencia 
      Caption         =   "Fechas de Vigencia"
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
      Height          =   2295
      Left            =   5160
      TabIndex        =   12
      Top             =   600
      Width           =   3255
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia_Unidad 
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
         Left            =   1440
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   154861569
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia_Unidad 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   154861569
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio de Vigencia"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   2175
      End
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
      TabIndex        =   11
      Top             =   3840
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid Grid_Medida 
         Height          =   1695
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         ForeColorSel    =   -2147483638
         BackColorBkg    =   16777215
      End
   End
   Begin VB.TextBox Txt_Codigo_Unidad 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txt_Nombre_Unidad 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
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
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4935
      Begin VB.TextBox Txt_Simbolo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox Txt_Comentario_Unidad 
         Enabled         =   0   'False
         Height          =   855
         Left            =   1680
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   1
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Txt_Descripcion_Unidad 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   1000
         TabIndex        =   2
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Símbolo"
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   24
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nota"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Unidad"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
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
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "UNIDADES DE MEDIDA"
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
      TabIndex        =   10
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "Frm_Cat_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Consulta_Unidades(Cadena As String)
Dim Rs_Consulta_Cat_Medidas As rdoResultset   'Manejo de registro
        Grid_Medida.Rows = 0
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Unidad.Enabled = True
        Txt_Nombre_Unidad.Enabled = True
        Txt_Comentario_Unidad.Enabled = True
        Txt_Simbolo.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Unidades_Medida"
    Mi_SQL = Mi_SQL & " WHERE (Clave_Unidad ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Medidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Medidas.EOF Then
        'Pone un encabezado en el grid
        Grid_Medida.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Descripción" & Chr(9) & "Nota" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Símbolo" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Medidas
                    Txt_Nombre_Unidad.text = .rdoColumns("Nombre")
                    Txt_Descripcion_Unidad.text = .rdoColumns("Descripcion")
                    Txt_Comentario_Unidad.text = .rdoColumns("Nota")
                    Dtp_Inicio_Vigencia_Unidad.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                    If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia_Unidad.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                
                Txt_Simbolo.text = .rdoColumns("Simbolo")
        
              
        End With
        While Not Rs_Consulta_Cat_Medidas.EOF
            Grid_Medida.AddItem Rs_Consulta_Cat_Medidas.rdoColumns("Clave_Unidad") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Nota") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Simbolo") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Estado")
            Grid_Medida.FixedRows = 1
            Rs_Consulta_Cat_Medidas.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Medida.FixedCols = 1
        Grid_Medida.ColWidth(0) = 1200
        Grid_Medida.ColAlignment(0) = flexAlignCenterCenter
        Grid_Medida.ColWidth(1) = 2000
        Grid_Medida.ColAlignment(1) = flexAlignLeftCenter
        Grid_Medida.ColWidth(2) = 2500
        Grid_Medida.ColAlignment(2) = flexAlignLeftCenter
        Grid_Medida.ColWidth(3) = 3000
        Grid_Medida.ColAlignment(3) = flexAlignLeftCenter
        Grid_Medida.ColWidth(4) = 1500
        Grid_Medida.ColAlignment(4) = flexAlignLeftCenter
        Grid_Medida.ColWidth(5) = 1500
        Grid_Medida.ColAlignment(5) = flexAlignLeftCenter
        Grid_Medida.ColWidth(6) = 3000
        Grid_Medida.ColAlignment(6) = flexAlignLeftCenter
        Grid_Medida.ColWidth(7) = 1200
        Grid_Medida.ColAlignment(7) = flexAlignLeftCenter
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Medidas.Close
End Sub

Public Sub Consulta_Productos(Cadena As String)
Dim Rs_Consulta_Cat_Unidades As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Txt_Descripcion_Unidad.Enabled = True
    Txt_Nombre_Unidad.Enabled = True
    Txt_Comentario_Unidad.Enabled = True
    Txt_Simbolo.Enabled = True
    Fra_Vigencia.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Unidades_Medida"
    Mi_SQL = Mi_SQL & " WHERE (Txt_Codigo_Unidad ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Unidades.EOF Then
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
    Rs_Consulta_Cat_Medidas.Close
End Sub

Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Unidades As rdoResultset   'Manejo de registro
    
    Grid_Medida.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Unidades_Medida WHERE (Clave_Unidad ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Unidades.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Unidades.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Unidades_Medida SET Estado='INACTIVO' WHERE (Clave_Unidad ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Unidades_Medida SET Estado='ACTIVO' WHERE (Clave_Unidad ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Unidades.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Public Sub Modifica_Unidades()
Dim Rs_Modificacion_Cat_Unidades As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Unidades_Medida"
    Mi_SQL = Mi_SQL & " WHERE Clave_Unidad='" & Txt_Codigo_Unidad.text & "'"
    Set Rs_Modificacion_Cat_Unidades = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Unidades.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Unidades
            .Edit
                .rdoColumns("Nombre") = UCase(Txt_Nombre_Unidad.text)
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Unidad.text)
                .rdoColumns("Nota") = UCase(Txt_Comentario_Unidad.text)
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia_Unidad.Value, "yyyy/MM/dd")
                If Chc_Vigencia.Value = 1 Then
                    .rdoColumns("Fecha_Fin_Vigencia") = Null
                 Else
                .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia_Unidad.Value, "yyyy/MM/dd")
                  End If
                .rdoColumns("Simbolo") = UCase(Txt_Simbolo.text)
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Unidades.Close
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
    If Txt_Codigo_Unidad.text <> "" Then
        Consulta_Unidades (Txt_Codigo_Unidad.text)
        
    Else
        MsgBox "Ingrese el código de la unidad", vbInformation
    End If
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Unidad <> "" Then
        Cambiar_Estado (Txt_Codigo_Unidad.text)
        
    Else
        MsgBox "Ingrese el código de la unidad", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Unidad.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Txt_Descripcion_Unidad.Enabled = True
                Txt_Nombre_Unidad.Enabled = True
                Txt_Comentario_Unidad.Enabled = True
                Txt_Simbolo.Enabled = True
                Fra_Vigencia.Enabled = True
                                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Codigo_Unidad.text) <> "" And Trim(Txt_Nombre_Unidad.text) <> "" Then
                    Modifica_Unidades
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
        Txt_Nombre_Unidad.Enabled = False
        Txt_Descripcion_Unidad.Enabled = False
        Txt_Comentario_Unidad.Enabled = False
        Txt_Simbolo.Enabled = False
        Fra_Vigencia.Enabled = False
        Dtp_Inicio_Vigencia_Unidad.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia_Unidad.Value = Format(Now(), "yyyy/MMM/dd")
        If Grid_Medida.Rows <> 0 Then
            Grid_Medida.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Chc_Vigencia_Click()
    If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia_Unidad.Enabled = False
    Else
        Dtp_Fin_Vigencia_Unidad.Enabled = True
    End If
End Sub
Private Sub Txt_Codigo_Unidad_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Public Sub Alta_Medidas()
Dim Rs_Alta_Cat_Medida As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Clave_Unidad FROM Cat_Unidades_Medida where Clave_Unidad ='" & Trim(Txt_Codigo_Unidad.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Producto.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Medida = Conectar_Ayudante.Recordset_Agregar("Cat_Unidades_Medida")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Medida
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Unidades_Medida", "Clave"), "00")
            .rdoColumns("Clave_Unidad") = Txt_Codigo_Unidad.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Nombre") = UCase(Txt_Nombre_Unidad.text)
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Unidad.text)
            .rdoColumns("Nota") = UCase(Txt_Comentario_Unidad.text)
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia_Unidad.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia_Unidad.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Simbolo") = UCase(Txt_Simbolo.text)
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Medida.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Medida.Rows = 0 Then
        Grid_Medida.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Descripción" & Chr(9) & "Nota" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Símbolo" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Medida.AddItem UCase(Txt_Codigo_Unidad.text) & Chr(9) & UCase(Trim(Txt_Nombre_Unidad.text)) & Chr(9) & UCase(Trim(Txt_Descripcion_Unidad.text)) & Chr(9) & UCase(Trim(Txt_Comentario_Unidad.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia_Unidad.Value) & Chr(9) & UCase("") & Chr(9) & UCase(Txt_Simbolo.text) & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Medida.AddItem UCase(Txt_Codigo_Unidad.text) & Chr(9) & UCase(Trim(Txt_Nombre_Unidad.text)) & Chr(9) & UCase(Trim(Txt_Descripcion_Unidad.text)) & Chr(9) & UCase(Trim(Txt_Comentario_Unidad.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia_Unidad.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia_Unidad.Value) & Chr(9) & UCase(Txt_Simbolo.text) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Medida.FixedCols = 1
        Grid_Medida.ColWidth(0) = 1200
        Grid_Medida.ColAlignment(0) = flexAlignCenterCenter
        Grid_Medida.ColWidth(1) = 2000
        Grid_Medida.ColAlignment(1) = flexAlignLeftCenter
        Grid_Medida.ColWidth(2) = 2500
        Grid_Medida.ColAlignment(2) = flexAlignLeftCenter
        Grid_Medida.ColWidth(3) = 3000
        Grid_Medida.ColAlignment(3) = flexAlignLeftCenter
        Grid_Medida.ColWidth(4) = 1500
        Grid_Medida.ColAlignment(4) = flexAlignLeftCenter
        Grid_Medida.ColWidth(5) = 1500
        Grid_Medida.ColAlignment(5) = flexAlignLeftCenter
        Grid_Medida.ColWidth(6) = 3000
        Grid_Medida.ColAlignment(6) = flexAlignLeftCenter
        Grid_Medida.ColWidth(7) = 1200
        Grid_Medida.ColAlignment(7) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Private Sub Btn_Nuevo_Click()
  If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Medida.Rows <> 0 Then
                    Grid_Medida.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Descripcion_Unidad.Enabled = True
                Txt_Nombre_Unidad.Enabled = True
                Txt_Comentario_Unidad.Enabled = True
                Txt_Simbolo.Enabled = True
                Fra_Vigencia.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Nombre_Unidad.text) <> "" And Trim(Txt_Codigo_Unidad.text) <> "" Then
                Alta_Medidas
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub

Private Sub Form_Load()
'Set Conexion = New Conectar
   ' Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia_Unidad.Value = Format(Now(), "yyyy/MMM/dd")
    Dtp_Fin_Vigencia_Unidad.Value = Format(Now(), "yyyy/MMM/dd")
End Sub

Private Sub Grid_Medida_Click()
Dim Rs_Consulta_Cat_Unidades As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Medida.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Unidad.Enabled = True
        Txt_Nombre_Unidad.Enabled = True
        Txt_Comentario_Unidad.Enabled = True
        Txt_Simbolo.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Unidades_Medida"
        Mi_SQL = Mi_SQL & " WHERE Clave_Unidad='" & Grid_Medida.TextMatrix(Grid_Medida.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Unidades.EOF Then
            With Rs_Consulta_Cat_Unidades
                Txt_Codigo_Unidad.text = .rdoColumns("Clave_Unidad")
                Txt_Nombre_Unidad.text = .rdoColumns("Nombre")
                'If Not IsNull(.rdoColumns("Descripcion")) Then
                '    Txt_Descripcion_Unidad.text = .rdoColumns("Descripcion")
                '    End If
                Txt_Descripcion_Unidad.text = IIf(IsNull(.rdoColumns(3)), "", .rdoColumns(3))
                If Not IsNull(.rdoColumns(4)) Then
                     Txt_Comentario_Unidad.text = .rdoColumns(4)
                     End If
                Dtp_Inicio_Vigencia_Unidad.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia_Unidad.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                If Not IsNull(.rdoColumns("Simbolo")) Then
                    Txt_Simbolo.text = .rdoColumns("Simbolo")
                    End If
                
            End With
        End If
        Rs_Consulta_Cat_Medidas.Close
    End If
End Sub

Public Sub Consulta()
Dim Rs_Consulta_Cat_Medidas As rdoResultset   'Manejo de registro
    
    Grid_Medida.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Unidades_Medida"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Medidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Medidas.EOF Then
        'Pone un encabezado en el grid
        Grid_Medida.AddItem "Clave" & Chr(9) & "Nombre" & Chr(9) & "Descripción" & Chr(9) & "Nota" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Símbolo" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Medidas.EOF
            Grid_Medida.AddItem Rs_Consulta_Cat_Medidas.rdoColumns("Clave_Unidad") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Nota") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Simbolo") & Chr(9) & Rs_Consulta_Cat_Medidas.rdoColumns("Estado")
            Grid_Medida.FixedRows = 1
            Rs_Consulta_Cat_Medidas.MoveNext
        Wend
        'Tamaño de las columnas en el grid
     Grid_Medida.FixedCols = 1
        Grid_Medida.ColWidth(0) = 1200
        Grid_Medida.ColAlignment(0) = flexAlignCenterCenter
        Grid_Medida.ColWidth(1) = 2000
        Grid_Medida.ColAlignment(1) = flexAlignLeftCenter
        Grid_Medida.ColWidth(2) = 2500
        Grid_Medida.ColAlignment(2) = flexAlignLeftCenter
        Grid_Medida.ColWidth(3) = 3000
        Grid_Medida.ColAlignment(3) = flexAlignLeftCenter
        Grid_Medida.ColWidth(4) = 1500
        Grid_Medida.ColAlignment(4) = flexAlignLeftCenter
        Grid_Medida.ColWidth(5) = 1500
        Grid_Medida.ColAlignment(5) = flexAlignLeftCenter
        Grid_Medida.ColWidth(6) = 3000
        Grid_Medida.ColAlignment(6) = flexAlignLeftCenter
        Grid_Medida.ColWidth(7) = 1200
        Grid_Medida.ColAlignment(7) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Medidas.Close
End Sub
