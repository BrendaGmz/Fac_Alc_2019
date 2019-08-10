VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Tipos_Comprobantes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2070
      Picture         =   "Cat_Tipos_Comprobantes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "M"
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7185
      Picture         =   "Cat_Tipos_Comprobantes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3765
      Picture         =   "Cat_Tipos_Comprobantes.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "B"
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5475
      Picture         =   "Cat_Tipos_Comprobantes.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "C"
      Top             =   4800
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   360
      Picture         =   "Cat_Tipos_Comprobantes.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "A"
      Top             =   4800
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
      Height          =   1935
      Left            =   5160
      TabIndex        =   7
      Top             =   600
      Width           =   3375
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
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
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90505217
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
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
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.TextBox Txt_Clave_Tipo_Comprobante 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Txt_Descripcion_Tipo_Comprobante 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   1410
      Width           =   1935
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
      Top             =   2640
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid Grid_Comprobantes 
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   240
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4815
      Begin VB.CheckBox Chc_Maximo 
         Caption         =   "Dos valores máximos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Txt_NS 
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         Text            =   "NS"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Txt_NdS 
         Height          =   285
         Left            =   3600
         TabIndex        =   21
         Text            =   "NdS"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Txt_Valor_Maximo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Máximo"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Clave tipo de comprobante"
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
         Width           =   2415
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
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "TIPOS DE COMPROBANTES"
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
      Width           =   8265
   End
End
Attribute VB_Name = "Frm_Cat_Tipos_Comprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Clave_Tipo_Comprobante.text <> "" Then
        Consulta_Comprobantes (Txt_Clave_Tipo_Comprobante.text)
        
    Else
        MsgBox "Ingrese el código del comprobante", vbInformation
    End If
End Sub
Public Sub Consulta_Comprobantes(Cadena As String)
Dim Rs_Consulta_Cat_Comprobantes As rdoResultset   'Manejo de registro
        Grid_Comprobantes.Rows = 0
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Tipo_Comprobante.Enabled = True
        Txt_Valor_Maximo.Enabled = True
        Txt_NS.Enabled = True
        Txt_NdS.Enabled = True
        Chc_Maximo.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipos_Comprobantes"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Tipo_Comprobante ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Comprobantes.EOF Then
        'Pone un encabezado en el grid
        Grid_Comprobantes.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Valor Máximo" & Chr(9) & "Valor Máximo-NS" & Chr(9) & "Valor Máximo-NdS" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Comprobantes
                    Txt_Descripcion_Tipo_Comprobante.text = .rdoColumns("Descripcion")
                If Not IsNull(.rdoColumns("NdS")) And Not IsNull(.rdoColumns("NS")) And .rdoColumns("NdS") <> "" And .rdoColumns("NS") <> "" Then
                    Chc_Maximo.Value = 1
                    Chc_Maximo_Click
                    Txt_NS.text = .rdoColumns("NS")
                    Txt_NdS.text = .rdoColumns("NdS")
                    Else
                    Chc_Maximo.Value = 0
                    Chc_Maximo_Click
                    Txt_Valor_Maximo.text = .rdoColumns("Valor_Maximo")
                    End If
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
         While Not Rs_Consulta_Cat_Comprobantes.EOF
                If Not IsNull(Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS")) And Not IsNull(Rs_Consulta_Cat_Comprobantes.rdoColumns("NS")) And Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS") <> "" And Rs_Consulta_Cat_Comprobantes.rdoColumns("NS") <> "" Then
                    Grid_Comprobantes.AddItem Rs_Consulta_Cat_Comprobantes.rdoColumns("Codigo_Tipo_Comprobante") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Descripcion") & Chr(9) & "" & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("NS") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado")
                    Grid_Comprobantes.FixedRows = 1
                    Rs_Consulta_Cat_Comprobantes.MoveNext
                Else
                    Grid_Comprobantes.AddItem Rs_Consulta_Cat_Comprobantes.rdoColumns("Codigo_Tipo_Comprobante") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Valor_Maximo") & Chr(9) & "" & Chr(9) & "" & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado")
                    Grid_Comprobantes.FixedRows = 1
                    Rs_Consulta_Cat_Comprobantes.MoveNext
                    End If
                Wend
            'Tamaño de las columnas en el grid
            Grid_Comprobantes.FixedCols = 1
            Grid_Comprobantes.ColWidth(0) = 1000
            Grid_Comprobantes.ColAlignment(0) = flexAlignCenterCenter
            Grid_Comprobantes.ColWidth(1) = 2000
            Grid_Comprobantes.ColAlignment(1) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(2) = 1200
            Grid_Comprobantes.ColAlignment(2) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(3) = 1500
            Grid_Comprobantes.ColAlignment(3) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(4) = 1500
            Grid_Comprobantes.ColAlignment(4) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(5) = 1500
            Grid_Comprobantes.ColAlignment(5) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(6) = 1500
            Grid_Comprobantes.ColAlignment(6) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(7) = 1500
            Grid_Comprobantes.ColAlignment(7) = flexAlignLeftCenter
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Comprobantes.Close
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Clave_Tipo_Comprobante <> "" Then
        Cambiar_Estado (Txt_Clave_Tipo_Comprobante.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Comprobantes As rdoResultset   'Manejo de registro
    
    Grid_Comprobantes.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Tipos_Comprobantes WHERE (Codigo_Tipo_Comprobante ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Tipos_Comprobantes SET Estado='INACTIVO' WHERE (Codigo_Tipo_Comprobante ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Tipos_Comprobantes SET Estado='ACTIVO' WHERE (Codigo_Tipo_Comprobante ='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Comprobantes.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub


Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Clave_Tipo_Comprobante.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                                    
            Else
                MsgBox "Ingrese el tipo de comprobante", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" And Trim(Txt_Descripcion_Tipo_Comprobante.text) <> "" Then
                If Trim(Txt_Valor_Maximo.text) <> "" Or (Txt_NS.text <> "" And Txt_NdS.text <> "") Then
                    Modifica_Comprobantes
                 Else
                    MsgBox "Faltan datos para actualizar", vbInformation
                End If
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        
End Sub
Public Sub Modifica_Comprobantes()
Dim Rs_Modificacion_Cat_Comprobantes As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tipos_Comprobantes"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Comprobante='" & Txt_Clave_Tipo_Comprobante.text & "'"
    Set Rs_Modificacion_Cat_Comprobantes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Comprobantes.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Comprobantes
            .Edit
                .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Tipo_Comprobante.text)
                If Chc_Maximo.Value = 1 Then
                .rdoColumns("Valor_Maximo") = UCase("")
                 .rdoColumns("NS") = UCase(Txt_NS.text)
                 .rdoColumns("NdS") = UCase(Txt_NdS.text)
                Else
                 .rdoColumns("Valor_Maximo") = UCase(Txt_Valor_Maximo.text)
                 .rdoColumns("NS") = UCase("")
                 .rdoColumns("NdS") = UCase("")
                End If
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
                If Grid_Comprobantes.Rows <> 0 Then
                    Grid_Comprobantes.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Clave_Tipo_Comprobante.Enabled = True
                Txt_Descripcion_Tipo_Comprobante.Enabled = True
                Txt_Valor_Maximo.Enabled = True
                Fra_Vigencia.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
                Chc_Maximo.Enabled = True
                Chc_Maximo.Value = 0
                Chc_Maximo_Click
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Clave_Tipo_Comprobante.text) <> "" And Trim(Txt_Descripcion_Tipo_Comprobante.text) <> "" Then
            If Trim(Txt_Valor_Maximo.text) <> "" Or (Txt_NS.text <> "" And Txt_NdS.text <> "") Then
                Alta_Comprobantes
                Else
                MsgBox "Faltan datos para dar de alta", vbInformation
                End If
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Comprobantes()
Dim Rs_Alta_Cat_Comprobantes As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Tipo_Comprobante FROM Cat_Tipos_Comprobantes where Codigo_Tipo_Comprobante ='" & Trim(Txt_Clave_Tipo_Comprobante.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Clave_Tipo_Comprobante.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Agregar("Cat_Tipos_Comprobantes")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Comprobantes
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Comprobantes", "Clave"), "00")
            .rdoColumns("Codigo_Tipo_Comprobante") = Txt_Clave_Tipo_Comprobante.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion_Tipo_Comprobante.text)
            If Chc_Maximo.Value = 1 Then
                .rdoColumns("NS") = UCase(Txt_NS.text)
                .rdoColumns("NdS") = UCase(Txt_NdS.text)
                Else
                .rdoColumns("Valor_Maximo") = UCase(Txt_Valor_Maximo.text)
                End If
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
    Rs_Alta_Cat_Comprobantes.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Comprobantes.Rows = 0 Then
        Grid_Comprobantes.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Valor Máximo" & Chr(9) & "Valor Máximo-NS" & Chr(9) & "Valor Máximo-NdS" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Comprobantes.AddItem UCase(Txt_Clave_Tipo_Comprobante.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Tipo_Comprobante.text)) & Chr(9) & UCase(Trim(Txt_Valor_Maximo.text)) & Chr(9) & UCase(Trim(Txt_NS.text)) & Chr(9) & UCase(Trim(Txt_NdS.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Comprobantes.AddItem UCase(Txt_Clave_Tipo_Comprobante.text) & Chr(9) & UCase(Trim(Txt_Descripcion_Tipo_Comprobante.text)) & Chr(9) & UCase(Trim(Txt_Valor_Maximo.text)) & Chr(9) & UCase(Trim(Txt_NS.text)) & Chr(9) & UCase(Trim(Txt_NdS.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Comprobantes.FixedCols = 1
            Grid_Comprobantes.ColWidth(0) = 1000
            Grid_Comprobantes.ColAlignment(0) = flexAlignCenterCenter
            Grid_Comprobantes.ColWidth(1) = 2000
            Grid_Comprobantes.ColAlignment(1) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(2) = 1200
            Grid_Comprobantes.ColAlignment(2) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(3) = 1500
            Grid_Comprobantes.ColAlignment(3) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(4) = 1500
            Grid_Comprobantes.ColAlignment(4) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(5) = 1500
            Grid_Comprobantes.ColAlignment(5) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(6) = 1500
            Grid_Comprobantes.ColAlignment(6) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(7) = 1500
            Grid_Comprobantes.ColAlignment(7) = flexAlignLeftCenter
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
        Txt_Descripcion_Tipo_Comprobante.Enabled = False
        Txt_Valor_Maximo.Enabled = False
        Chc_Maximo.Value = 0
        Chc_Maximo.Enabled = False
        Fra_Vigencia.Enabled = False
        Dtp_Fin_Vigencia.MinDate = Format(Now(), "yyyy/MMM/dd")
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Consulta
    End If
End Sub

Private Sub Chc_Maximo_Click()
    If Chc_Maximo.Value = 1 Then
        Txt_NS.Visible = True
        Txt_NdS.Visible = True
        Txt_NS.text = ""
        Txt_NdS.text = ""
        Txt_Valor_Maximo.Visible = False
    Else
        Txt_NS.Visible = False
        Txt_NdS.Visible = False
        Txt_Valor_Maximo.Visible = True
        Txt_Valor_Maximo.text = ""
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
Private Sub Txt_Clave_Tipo_Comprobante_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Valor_Maximo_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_NS_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_NdS_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Comprobantes As rdoResultset   'Manejo de registro
    
    Grid_Comprobantes.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipos_Comprobantes ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Comprobantes.EOF Then
        'Pone un encabezado en el grid
        Grid_Comprobantes.AddItem "Clave" & Chr(9) & "Descripción" & Chr(9) & "Valor Máximo" & Chr(9) & "Valor Máximo-NS" & Chr(9) & "Valor Máximo-NdS" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        
            'Llenado del grid
            While Not Rs_Consulta_Cat_Comprobantes.EOF
                If Not IsNull(Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS")) And Not IsNull(Rs_Consulta_Cat_Comprobantes.rdoColumns("NS")) And Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS") <> "" And Rs_Consulta_Cat_Comprobantes.rdoColumns("NS") <> "" Then
                    Grid_Comprobantes.AddItem Rs_Consulta_Cat_Comprobantes.rdoColumns("Codigo_Tipo_Comprobante") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Descripcion") & Chr(9) & "" & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("NS") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("NdS") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado")
                    Grid_Comprobantes.FixedRows = 1
                    Rs_Consulta_Cat_Comprobantes.MoveNext
                Else
                    Grid_Comprobantes.AddItem Rs_Consulta_Cat_Comprobantes.rdoColumns("Codigo_Tipo_Comprobante") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Descripcion") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Valor_Maximo") & Chr(9) & "" & Chr(9) & "" & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Comprobantes.rdoColumns("Estado")
                    Grid_Comprobantes.FixedRows = 1
                    Rs_Consulta_Cat_Comprobantes.MoveNext
                    End If
                Wend
            'Tamaño de las columnas en el grid
            Grid_Comprobantes.FixedCols = 1
            Grid_Comprobantes.ColWidth(0) = 1000
            Grid_Comprobantes.ColAlignment(0) = flexAlignCenterCenter
            Grid_Comprobantes.ColWidth(1) = 2000
            Grid_Comprobantes.ColAlignment(1) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(2) = 1200
            Grid_Comprobantes.ColAlignment(2) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(3) = 1500
            Grid_Comprobantes.ColAlignment(3) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(4) = 1500
            Grid_Comprobantes.ColAlignment(4) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(5) = 1500
            Grid_Comprobantes.ColAlignment(5) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(6) = 1500
            Grid_Comprobantes.ColAlignment(6) = flexAlignLeftCenter
            Grid_Comprobantes.ColWidth(7) = 1500
            Grid_Comprobantes.ColAlignment(7) = flexAlignLeftCenter
    End If
    Rs_Consulta_Cat_Comprobantes.Close
End Sub


Private Sub Grid_Comprobantes_Click()
Dim Rs_Consulta_Cat_Comprobantes As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Comprobantes.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Txt_Descripcion_Tipo_Comprobante.Enabled = True
        Txt_Valor_Maximo.Enabled = True
        Txt_NS.Enabled = True
        Txt_NdS.Enabled = True
        Chc_Maximo.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Tipos_Comprobantes"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Tipo_Comprobante='" & Grid_Comprobantes.TextMatrix(Grid_Comprobantes.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Comprobantes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Comprobantes.EOF Then
            With Rs_Consulta_Cat_Comprobantes
                Txt_Clave_Tipo_Comprobante.text = .rdoColumns("Codigo_Tipo_Comprobante")
                Txt_Descripcion_Tipo_Comprobante.text = .rdoColumns("Descripcion")
                If Not IsNull(.rdoColumns("NdS")) And Not IsNull(.rdoColumns("NS")) And .rdoColumns("NdS") <> "" And .rdoColumns("NS") <> "" Then
                    Chc_Maximo.Value = 1
                    Chc_Maximo_Click
                    Txt_NS.text = .rdoColumns("NS")
                    Txt_NdS.text = .rdoColumns("NdS")
                    Else
                    Chc_Maximo.Value = 0
                    Chc_Maximo_Click
                    Txt_Valor_Maximo.text = .rdoColumns("Valor_Maximo")
                    End If
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
        Rs_Consulta_Cat_Comprobantes.Close
    End If
End Sub

