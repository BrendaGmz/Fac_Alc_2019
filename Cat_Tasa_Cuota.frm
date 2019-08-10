VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Cat_Tasa_Cuota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo Tasas o Cuotas de Impuestos"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   480
      Picture         =   "Cat_Tasa_Cuota.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "A"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4080
      Picture         =   "Cat_Tasa_Cuota.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "C"
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2880
      Picture         =   "Cat_Tasa_Cuota.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "B"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5280
      Picture         =   "Cat_Tasa_Cuota.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1680
      Picture         =   "Cat_Tasa_Cuota.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "M"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   975
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
      TabIndex        =   24
      Top             =   4560
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid Grid_Tasa_Cuota 
         Height          =   1695
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   11
         FixedRows       =   0
         ForeColorSel    =   -2147483638
         BackColorBkg    =   16777215
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
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   3495
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   166199297
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   166199297
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   23
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
         Index           =   10
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox Txt_Minimo 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
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
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6375
      Begin VB.TextBox Txt_Clave_Tasa 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Factor 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Tasa_Cuota.frx":0552
         Left            =   4080
         List            =   "Cat_Tasa_Cuota.frx":0589
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Retencion 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Tasa_Cuota.frx":05C9
         Left            =   4080
         List            =   "Cat_Tasa_Cuota.frx":05D0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Traslado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Tasa_Cuota.frx":05E3
         Left            =   4080
         List            =   "Cat_Tasa_Cuota.frx":05EA
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Impuesto 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Tasa_Cuota.frx":05FC
         Left            =   1320
         List            =   "Cat_Tasa_Cuota.frx":0603
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Txt_Maximo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Tipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Cat_Tasa_Cuota.frx":0615
         Left            =   4080
         List            =   "Cat_Tasa_Cuota.frx":061C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Clave"
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
         Index           =   9
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Factor"
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
         Index           =   8
         Left            =   3480
         TabIndex        =   14
         Top             =   1680
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
         Index           =   7
         Left            =   3240
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
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
         Index           =   5
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Impuesto"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
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
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Mínimo"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   975
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
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "TASAS O CUOTAS"
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
      TabIndex        =   17
      Top             =   0
      Width           =   6585
   End
End
Attribute VB_Name = "Frm_Cat_Tasa_Cuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Llena_Combos()
    Call Conectar_Ayudante.Llena_Combo_Item("Impuesto_ID,Descripcion", "Cat_Impuestos ORDER BY Clave ASC", Cmb_Impuesto, 0, "")
    Tipo = Conectar_Ayudante.Llena_Combo("Codigo_Tipo_Factor,Clave", "Cat_Tipo_Factor", Cmb_Factor, 0, "Clave ASC")
End Sub



Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Clave_Tasa.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                                    
            Else
                MsgBox "Ingrese código del producto", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" And Trim(Txt_Clave_Tasa.text) <> "" And Trim(Txt_Maximo.text) <> "" And Cmb_Impuesto.ListIndex > -1 And Cmb_Tipo.ListIndex > -1 And Cmb_Retencion.ListIndex > -1 And Cmb_Traslado.ListIndex > -1 And Cmb_FactorListIndex > -1 Then
         If Cmb_Tipo.text = "RANGO" And Txt_Minimo.text = "" Then
                MsgBox "Ingrese el valor mínimo"
                Else
                  Modifica_Tasas
            End If
            'If Trim(Txt_Descripcion_Comprobante.text) <> "" And Cmb_Persona_Fisica.ListIndex > -1 And Cmb_Persona_Moral.ListIndex > -1 Then
                   
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
        End If
End Sub
Public Sub Modifica_Tasas()
Dim Rs_Modificacion_Cat_Tasas As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tasas_Cuotas"
    Mi_SQL = Mi_SQL & " WHERE Clave=" & Txt_Clave_Tasa.text & ""
    Set Rs_Modificacion_Cat_Tasas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Tasas.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Tasas
            .Edit
                .rdoColumns("Tipo") = Cmb_Tipo.text
                .rdoColumns("Valor_Minimo") = Txt_Minimo.text
                .rdoColumns("Valor_Maximo") = Txt_Maximo.text
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
                .rdoColumns("Impuesto") = Cmb_Impuesto.text
                .rdoColumns("Factor") = Cmb_Factor.text
                .rdoColumns("Traslado") = Cmb_Traslado.text
                .rdoColumns("Retencion") = Cmb_Retencion.text
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
    Rs_Modificacion_Cat_Tasas.Close
    MsgBox "Modificación exitosa", vbInformation
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
                If Grid_Tasa_Cuota.Rows <> 0 Then
                    Grid_Tasa_Cuota.Rows = 0
                End If
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Clave_Tasa.Locked = True
                Txt_Clave_Tasa = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tasas_Cuotas", "Clave"), "00")
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Txt_Minimo.Enabled = True
                Txt_Maximo.Enabled = True
                Cmb_Impuesto.Enabled = True
                Cmb_Tipo.Enabled = True
                Cmb_Retencion.Enabled = True
                Cmb_Traslado.Enabled = True
                Cmb_Factor.Enabled = True
                Fra_Vigencia.Enabled = True
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Clave_Tasa.text) <> "" And Trim(Txt_Maximo.text) <> "" And Cmb_Impuesto.ListIndex > -1 And Cmb_Tipo.ListIndex > -1 And Cmb_Retencion.ListIndex > -1 And Cmb_Traslado.ListIndex > -1 And Cmb_FactorListIndex > -1 Then
        If Cmb_Tipo.text = "RANGO" And Txt_Minimo.text = "" Then
            MsgBox "Ingrese el valor mínimo"
            Else
                Alta_Tasas
            End If
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Tasas()
Dim Rs_Alta_Cat_Tasas As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Clave FROM Cat_Tasas_Cuotas where Clave =" & Trim(Txt_Clave_Tasa.text) & ""
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Clave_Tasa.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Tasas = Conectar_Ayudante.Recordset_Agregar("Cat_Tasas_Cuotas")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Tasas
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tasas_Cuotas", "Clave"), "00")
            '.rdoColumns("Codigo_Tasa_Cuota") = Txt_Clave_Tasa.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Valor_Minimo") = UCase(Txt_Minimo.text)
            .rdoColumns("Valor_Maximo") = UCase(Txt_Maximo.text)
            .rdoColumns("Tipo") = UCase(Cmb_Tipo.text)
            .rdoColumns("Impuesto") = UCase(Cmb_Impuesto.text)
            .rdoColumns("Factor") = UCase(Cmb_Factor.text)
            .rdoColumns("Traslado") = UCase(Cmb_Traslado.text)
            .rdoColumns("Retencion") = UCase(Cmb_Retencion.text)
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
    Rs_Alta_Cat_Tasas.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Tasa_Cuota.Rows = 0 Then
        Grid_Tasa_Cuota.AddItem "Clave" & Chr(9) & "Tipo" & Chr(9) & "Mínimo" & Chr(9) & "Máximo" & Chr(9) & "Impuesto" & Chr(9) & "Factor" & Chr(9) & "Traslado" & Chr(9) & "Retención" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Tasa_Cuota.AddItem UCase(Txt_Clave_Tasa.text) & Chr(9) & UCase(Trim(Cmb_Tipo.text)) & Chr(9) & UCase(Txt_Minimo.text) & Chr(9) & UCase(Txt_Maximo.text) & Chr(9) & UCase(Cmb_Impuesto.text) & Chr(9) & UCase(Cmb_Factor.text) & Chr(9) & UCase(Cmb_Traslado.text) & Chr(9) & UCase(Cmb_Retencion.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Tasa_Cuota.AddItem UCase(Txt_Clave_Tasa.text) & Chr(9) & UCase(Trim(Cmb_Tipo.text)) & Chr(9) & UCase(Txt_Minimo.text) & Chr(9) & UCase(Txt_Maximo.text) & Chr(9) & UCase(Cmb_Impuesto.text) & Chr(9) & UCase(Cmb_Factor.text) & Chr(9) & UCase(Cmb_Traslado.text) & Chr(9) & UCase(Cmb_Retencion.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
      Grid_Tasa_Cuota.FixedCols = 1
        Grid_Tasa_Cuota.ColWidth(0) = 800
        Grid_Tasa_Cuota.ColAlignment(0) = flexAlignCenterCenter
        Grid_Tasa_Cuota.ColWidth(1) = 800
        Grid_Tasa_Cuota.ColAlignment(1) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(2) = 1000
        Grid_Tasa_Cuota.ColAlignment(2) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(3) = 1000
        Grid_Tasa_Cuota.ColAlignment(3) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(4) = 800
        Grid_Tasa_Cuota.ColAlignment(4) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(5) = 800
        Grid_Tasa_Cuota.ColAlignment(5) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(6) = 800
        Grid_Tasa_Cuota.ColAlignment(6) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(7) = 1200
        Grid_Tasa_Cuota.ColAlignment(7) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(8) = 1200
        Grid_Tasa_Cuota.ColAlignment(8) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColAlignment(6) = flexAlignLeftCenter
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
        Txt_Minimo.Enabled = False
        Txt_Maximo.Enabled = False
        Cmb_Impuesto.Enabled = False
        Cmb_Tipo.Enabled = False
        Cmb_Retencion.Enabled = False
        Cmb_Traslado.Enabled = False
        Cmb_Factor.Enabled = False
        Fra_Vigencia.Enabled = False
        Dtp_Fin_Vigencia.MinDate = Format(Now(), "yyyy/MMM/dd")
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Chc_Vigencia.Value = 1
        Chc_Vigencia_Click
        Txt_Clave_Tasa.Locked = False
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

Private Sub Cmb_Tipo_Click()
    If Cmb_Tipo = "FIJO" Then
        Txt_Minimo.Enabled = False
        Txt_Minimo.text = ""
    Else
        Txt_Minimo.Enabled = True
        End If
End Sub

Private Sub Form_Load()
Set Conectar_Ayudante = New Ayudante
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Llena_Combos
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub

Private Sub Grid_Tasa_Cuota_Click()
Dim Rs_Consulta_Cat_Tasa As rdoResultset
'Si el grid tiene filas, entonces hace la consulta
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Tasa_Cuota.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Txt_Minimo.Enabled = True
        Txt_Maximo.Enabled = True
        Cmb_Impuesto.Enabled = True
        Cmb_Tipo.Enabled = True
        Cmb_Retencion.Enabled = True
        Cmb_Traslado.Enabled = True
        Cmb_Factor.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        Txt_Minimo.Enabled = True
        Txt_Minimo.Locked = False
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Tasas_Cuotas"
        Mi_SQL = Mi_SQL & " WHERE Clave=" & Grid_Tasa_Cuota.TextMatrix(Grid_Tasa_Cuota.RowSel, 0) & ""
        Set Rs_Consulta_Cat_Tasa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Tasa.EOF Then
            With Rs_Consulta_Cat_Tasa
                If Not IsNull(.rdoColumns("Clave")) Then Txt_Clave_Tasa.text = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Valor_Minimo")) Then Txt_Minimo.text = CStr(.rdoColumns("Valor_Minimo"))
                If Not IsNull(.rdoColumns("Valor_Maximo")) Then Txt_Maximo.text = CStr(.rdoColumns("Valor_Maximo"))
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
                
                For I = 0 To Cmb_Impuesto.ListCount - 1
                Cmb_Impuesto.ListIndex = I
                    If Cmb_Impuesto.text = .rdoColumns("Impuesto") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Tipo.ListCount - 1
                Cmb_Tipo.ListIndex = I
                    If Cmb_Tipo.text = .rdoColumns("Tipo") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Traslado.ListCount - 1
                Cmb_Traslado.ListIndex = I
                    If Cmb_Traslado.text = .rdoColumns("Traslado") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Retencion.ListCount - 1
                Cmb_Retencion.ListIndex = I
                    If Cmb_Retencion.text = .rdoColumns("Retencion") Then
                        Exit For
                        End If
                Next
                For I = 0 To Cmb_Factor.ListCount - 1
                Cmb_Factor.ListIndex = I
                    If Cmb_Factor.text = .rdoColumns("Factor") Then
                        Exit For
                        End If
                Next
            End With
        End If
        Rs_Consulta_Cat_Tasa.Close
    End If
End Sub

Private Sub Txt_Minimo_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Maximo_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Clave_Tasa_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Tasas As rdoResultset   'Manejo de registro
    
    Grid_Tasa_Cuota.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tasas_Cuotas ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Tasas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Tasas.EOF Then
        'Pone un encabezado en el grid
        Grid_Tasa_Cuota.AddItem "Clave" & Chr(9) & "Tipo" & Chr(9) & "Mínimo" & Chr(9) & "Máximo" & Chr(9) & "Impuesto" & Chr(9) & "Factor" & Chr(9) & "Traslado" & Chr(9) & "Retención" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Tasas.EOF
            With Rs_Consulta_Cat_Tasas
                Grid_Tasa_Cuota.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Tipo") & Chr(9) & .rdoColumns("Valor_Minimo") & Chr(9) & .rdoColumns("Valor_Maximo") & Chr(9) & .rdoColumns("Impuesto") & Chr(9) & .rdoColumns("Factor") & Chr(9) & .rdoColumns("Traslado") & Chr(9) & .rdoColumns("Retencion") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
                Grid_Tasa_Cuota.FixedRows = 1
                Rs_Consulta_Cat_Tasas.MoveNext
            End With
        Wend
        'Tamaño de las columnas en el grid
        Grid_Tasa_Cuota.FixedCols = 1
        Grid_Tasa_Cuota.ColWidth(0) = 800
        Grid_Tasa_Cuota.ColAlignment(0) = flexAlignCenterCenter
        Grid_Tasa_Cuota.ColWidth(1) = 800
        Grid_Tasa_Cuota.ColAlignment(1) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(2) = 1000
        Grid_Tasa_Cuota.ColAlignment(2) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(3) = 1000
        Grid_Tasa_Cuota.ColAlignment(3) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(4) = 800
        Grid_Tasa_Cuota.ColAlignment(4) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(5) = 800
        Grid_Tasa_Cuota.ColAlignment(5) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(6) = 800
        Grid_Tasa_Cuota.ColAlignment(6) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(7) = 1200
        Grid_Tasa_Cuota.ColAlignment(7) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(8) = 1200
        Grid_Tasa_Cuota.ColAlignment(8) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Tasas.Close
End Sub
Private Sub Btn_Eliminar_Click()
If Txt_Clave_Tasa.text <> "" Then
       Cambiar_Estado (Txt_Clave_Tasa.text)
        
    Else
        MsgBox "Ingrese el código", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Tasas As rdoResultset   'Manejo de registro
    
    Grid_Tasa_Cuota.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Tasas_Cuotas WHERE (Clave =" & Cadena & ")"
    Set Rs_Consulta_Cat_Tasas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Tasas.rdoColumns("Estado") = "ACTIVO" Then 'Or Rs_Consulta_Cat_Tasas.rdoColumns("Estado") <> "" Then
        Mi_SQL = "UPDATE Cat_Tasas_Cuotas SET Estado='INACTIVO' WHERE (Clave =" & Cadena & ")"
    Else
        Mi_SQL = "UPDATE Cat_Tasas_Cuotas SET Estado='ACTIVO' WHERE (Clave =" & Cadena & ")"
        End If
    Set Rs_Consulta_Cat_Tasas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Tasas.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Clave_Tasa.text <> "" Then
        Consulta_Tasa (Txt_Clave_Tasa.text)
        
    Else
        MsgBox "Ingrese la clave", vbInformation
    End If
End Sub
Public Sub Consulta_Tasa(Cadena As String)
Dim Rs_Consulta_Cat_Tasas As rdoResultset   'Manejo de registro
    Grid_Tasa_Cuota.Rows = 0
    Btn_Salir.Caption = "Regresar"
    Txt_Minimo.Enabled = True
    Txt_Maximo.Enabled = True
    Cmb_Impuesto.Enabled = True
    Cmb_Tipo.Enabled = True
    Cmb_Retencion.Enabled = True
    Cmb_Traslado.Enabled = True
    Cmb_Factor.Enabled = True
    Fra_Vigencia.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Txt_Minimo.Enabled = True
    Txt_Minimo.Locked = False
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tasas_Cuotas"
    Mi_SQL = Mi_SQL & " WHERE (Clave =" & Cadena & ")"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Tasas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Tasas.EOF Then
        'Pone un encabezado en el grid
        Grid_Tasa_Cuota.AddItem "Clave" & Chr(9) & "Tipo" & Chr(9) & "Mínimo" & Chr(9) & "Máximo" & Chr(9) & "Impuesto" & Chr(9) & "Factor" & Chr(9) & "Traslado" & Chr(9) & "Retención" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Tasas.EOF
            With Rs_Consulta_Cat_Tasas
                Grid_Tasa_Cuota.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Tipo") & Chr(9) & .rdoColumns("Valor_Minimo") & Chr(9) & .rdoColumns("Valor_Maximo") & Chr(9) & .rdoColumns("Impuesto") & Chr(9) & .rdoColumns("Factor") & Chr(9) & .rdoColumns("Traslado") & Chr(9) & .rdoColumns("Retencion") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
                Grid_Tasa_Cuota.FixedRows = 1
                Rs_Consulta_Cat_Tasas.MoveNext
            End With
        Wend
        'Tamaño de las columnas en el grid
        Grid_Tasa_Cuota.FixedCols = 1
        Grid_Tasa_Cuota.ColWidth(0) = 800
        Grid_Tasa_Cuota.ColAlignment(0) = flexAlignCenterCenter
        Grid_Tasa_Cuota.ColWidth(1) = 800
        Grid_Tasa_Cuota.ColAlignment(1) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(2) = 1000
        Grid_Tasa_Cuota.ColAlignment(2) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(3) = 1000
        Grid_Tasa_Cuota.ColAlignment(3) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(4) = 800
        Grid_Tasa_Cuota.ColAlignment(4) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(5) = 800
        Grid_Tasa_Cuota.ColAlignment(5) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(6) = 800
        Grid_Tasa_Cuota.ColAlignment(6) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(7) = 1200
        Grid_Tasa_Cuota.ColAlignment(7) = flexAlignLeftCenter
        Grid_Tasa_Cuota.ColWidth(8) = 1200
        Grid_Tasa_Cuota.ColAlignment(8) = flexAlignLeftCenter
        
        
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Tasas.Close
End Sub
