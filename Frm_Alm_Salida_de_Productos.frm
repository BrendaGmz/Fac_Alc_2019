VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Alm_Salidas_de_Producto 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Salidas de Almacen"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   9105
   Begin VB.PictureBox Pic_Salidas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8715
      Left            =   0
      ScaleHeight     =   8715
      ScaleWidth      =   9240
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      Begin VB.Frame Fra_Productos_Orden_Compra 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Productos de la orden de Compra"
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
         Height          =   2580
         Left            =   90
         TabIndex        =   21
         Top             =   2160
         Width           =   8880
         Begin VB.TextBox Txt_Cantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   2175
            Width           =   825
         End
         Begin VB.ComboBox Cmb_Descripcion 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   2175
            Width           =   6435
         End
         Begin VB.CommandButton Btn_Agregar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7425
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Alm_Salida_de_Productos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2100
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Entradas_Productos 
            Height          =   1740
            Left            =   105
            TabIndex        =   7
            Top             =   225
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3069
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Lbl_Descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Descripción"
            Height          =   195
            Left            =   960
            TabIndex        =   30
            Top             =   1995
            Width           =   840
         End
         Begin VB.Label Lbl_Cantidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1995
            Width           =   630
         End
      End
      Begin VB.Frame Fra_Datos_Salida 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos de la Salida"
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
         Height          =   1770
         Left            =   90
         TabIndex        =   18
         Top             =   375
         Width           =   8880
         Begin VB.TextBox Txt_Cliente_ID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   607
            Width           =   3030
         End
         Begin VB.ComboBox Cmb_Tipo_Entrada 
            Height          =   315
            ItemData        =   "Frm_Alm_Salida_de_Productos.frx":32B6
            Left            =   5760
            List            =   "Frm_Alm_Salida_de_Productos.frx":32C3
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   255
            Width           =   3030
         End
         Begin VB.TextBox Txt_No_Salida 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1020
            TabIndex        =   1
            Top             =   255
            Width           =   3990
         End
         Begin VB.ComboBox Cmb_Orden_Compra 
            Height          =   315
            Left            =   1020
            TabIndex        =   4
            Top             =   990
            Width           =   4035
         End
         Begin VB.ComboBox Cmb_Cliente_Salida_Almacen 
            Height          =   315
            Left            =   1020
            TabIndex        =   3
            Top             =   615
            Width           =   4035
         End
         Begin VB.TextBox Txt_Observaciones 
            Height          =   315
            Left            =   1020
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1320
            Width           =   7770
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Salida 
            Height          =   315
            Left            =   5760
            TabIndex        =   5
            Top             =   990
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd '  de  ' MMMM '  de  ' yyyy"
            Format          =   58851328
            CurrentDate     =   39706
            MaxDate         =   73415
            MinDate         =   32874
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente ID"
            Height          =   195
            Left            =   5040
            TabIndex        =   33
            Top             =   630
            Width           =   690
         End
         Begin VB.Label Lbl_Tipo_Entrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo "
            Height          =   195
            Left            =   5040
            TabIndex        =   31
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Lbl_No_Salida 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Salida"
            Height          =   195
            Left            =   105
            TabIndex        =   27
            Top             =   315
            Width           =   690
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pedido"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   26
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   25
            Top             =   645
            Width           =   810
         End
         Begin VB.Label Lbl_Observaciones 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   105
            TabIndex        =   20
            Top             =   1425
            Width           =   870
         End
         Begin VB.Label Lbl_Fecha_Salida 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   5040
            TabIndex        =   19
            Top             =   1050
            Width           =   450
         End
      End
      Begin VB.Frame Fra_Detalles_Salida 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles de la Salida"
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
         Height          =   2520
         Left            =   90
         TabIndex        =   22
         Top             =   4755
         Width           =   8880
         Begin VB.TextBox Txt_Modifica_Cantidad 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   135
            TabIndex        =   35
            Top             =   270
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton Btn_Elimina_detalle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   105
            Picture         =   "Frm_Alm_Salida_de_Productos.frx":32E4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2025
            Width           =   1260
         End
         Begin VB.TextBox Txt_Total_Salida 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2130
            Width           =   1515
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Detalle_Salidas 
            Height          =   1740
            Left            =   105
            TabIndex        =   11
            Top             =   240
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   3069
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Lbl_Total_Salida 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Salida"
            Height          =   195
            Left            =   6300
            TabIndex        =   23
            Top             =   2190
            Width           =   840
         End
      End
      Begin VB.CommandButton Btn_Salir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7605
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":6596
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "A"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Regresar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regresar"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7605
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":9C95
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "A"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7605
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":D1F5
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "A"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3750
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":1088C
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "C"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Imprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3750
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":13E18
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "A"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   90
         Picture         =   "Frm_Alm_Salida_de_Productos.frx":172DE
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "A"
         Top             =   7305
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Label Lbl_Salidas_Almacen 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SALIDAS ALMACEN"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   24
         Top             =   -45
         Width           =   3570
      End
   End
End
Attribute VB_Name = "Frm_Alm_Salidas_de_Producto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Agregar_Click
'DESCRIPCION            : Agrega los productos a los detalles de la salida
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Agregar_Click()
Dim Cont_Filas As Integer
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
Dim Año As Integer
Dim Mes As Integer
Dim Dia As Integer
Dim Fecha_Valida As Date
Dim Agregar_Partida As Boolean

On Error GoTo Handler
    
    
    If Cmb_Descripcion.ListIndex > -1 And Val(Txt_Cantidad.Text) > 0 Then
         Grid_Detalle_Salidas.Cols = 4
         If Grid_Detalle_Salidas.Rows = 0 Then
             'Pone el encabezado en las columnas
             Grid_Detalle_Salidas.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID"
         End If
         'SE ASIGNAN LOS DETALLES AL GRID
         Grid_Detalle_Salidas.AddItem Txt_Cantidad.Text & Chr(9) & _
         Cmb_Descripcion.Text & Chr(9) & _
         Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & _
         "MANUAL"
         Grid_Detalle_Salidas.FixedRows = 1
         'SE ELIMINA DEL GRIN DE PRODUCTOS DE ORDEN DE COMPRA
         If Grid_Entradas_Productos.Rows > 2 Then
             Grid_Entradas_Productos.RemoveItem Grid_Entradas_Productos.RowSel
         Else
             Grid_Entradas_Productos.Rows = 0
         End If
         'Grid Partidas
        If Grid_Detalle_Salidas.Rows > 1 Then
             'Configura el grid
             Grid_Detalle_Salidas.ColWidth(0) = 800              'Ancho de las columnas
             Grid_Detalle_Salidas.ColWidth(1) = 7390
             Grid_Detalle_Salidas.ColAlignment(1) = 1            'Alínea la columna a la derecha
             Grid_Detalle_Salidas.ColWidth(2) = 0
             Grid_Detalle_Salidas.ColWidth(3) = 0
             'Pone el setfocus en la primera fila del Grid
             With Grid_Detalle_Salidas
                 .Col = 1
                 .Row = 1
                 .ColSel = .Cols - 1
                 .RowSel = 1
                 .TopRow = .Row
                 .SetFocus
             End With
         End If
         Call Calcula_Total_salida
         Txt_Cantidad.Text = ""
         Cmb_Descripcion.ListIndex = -0
    Else
      If Cmb_Orden_Compra.ListIndex > -1 Then
        Grid_Detalle_Salidas.Cols = 4
        If Grid_Detalle_Salidas.Rows = 0 Then
            'Pone el encabezado en las columnas
            Grid_Detalle_Salidas.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID"
        End If
        'SE ASIGNAN LOS DETALLES AL GRID
        Grid_Detalle_Salidas.AddItem Grid_Entradas_Productos.TextMatrix(Grid_Entradas_Productos.RowSel, 0) & Chr(9) & _
        Grid_Entradas_Productos.TextMatrix(Grid_Entradas_Productos.RowSel, 1) & Chr(9) & _
        Grid_Entradas_Productos.TextMatrix(Grid_Entradas_Productos.RowSel, 2) & Chr(9) & _
        "PEDIDO"
        Grid_Detalle_Salidas.FixedRows = 1
        'SE ELIMINA DEL GRIN DE PRODUCTOS DE ORDEN DE COMPRA
        If Grid_Entradas_Productos.Rows > 2 Then
            Grid_Entradas_Productos.RemoveItem Grid_Entradas_Productos.RowSel
        Else
            Grid_Entradas_Productos.Rows = 0
        End If
        'Grid Partidas
           If Grid_Detalle_Salidas.Rows > 1 Then
                'Configura el grid
                Grid_Detalle_Salidas.ColWidth(0) = 800              'Ancho de las columnas
                Grid_Detalle_Salidas.ColWidth(1) = 7390
                Grid_Detalle_Salidas.ColAlignment(1) = 1            'Alínea la columna a la derecha
                Grid_Detalle_Salidas.ColWidth(2) = 0
                Grid_Detalle_Salidas.ColWidth(3) = 0
                'Pone el setfocus en la primera fila del Grid
                With Grid_Detalle_Salidas
                    .Col = 1
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
            Call Calcula_Total_salida
        Else
           MsgBox "No hay datos que agregar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    End If
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub Btn_Buscar_Click()
Dim No_Salida As String
    Dtp_Fecha_Salida.Value = Now
    Grid_Entradas_Productos.Rows = 0
    Grid_Detalle_Salidas.Rows = 0
    Cmb_Cliente_Salida_Almacen.Text = ""
    Cmb_Orden_Compra.Text = ""
     Btn_Nuevo.Caption = "Nuevo"
     Btn_Salir.Visible = True
     Btn_Regresar.Visible = False
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    No_Salida = Format(InputBox("Numero de Salida a consultar"), "0000000000")
    No_Salida = Conectar_Ayudante.Quitar_Caracter(No_Salida, "'")
    If Trim(No_Salida) <> "" Then
        Call Consulta_Salida(No_Salida)
    Else
        Btn_Nuevo.Visible = True
    End If
End Sub

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Cancelar_Click
'DESCRIPCION            : Agrega los productos a los detalles de la salida
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 24-Febrero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Cancelar_Click()
Dim Mi_SQL As String
Dim Rs_Ope_Entrada_Detalles As rdoResultset
Dim Rs_Ope_Entrada As rdoResultset
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Consulta_Producto As rdoResultset
Dim Entrada_ID As String
Dim No_Control As String
Dim Rs_Alta_Tmp_Facturas_Proveedores As rdoResultset
Dim Comentarios_Cancelacion As String


    If MsgBox("¿Está seguro de cancelar la salida?", vbQuestion + vbYesNo) = vbYes Then
        Comentarios_Cancelacion = InputBox("Teclee los comentarios de la cancelación")
        Conexion_Base.BeginTrans
        Cantidad_Total = 0
        'SE OBTIENE LA CANTIDAD TOTAL DE LA SALIDA
        For Cont_Fila = 1 To Grid_Detalle_Salidas.Rows - 1 Step 1
            Cantidad_Total = Cantidad_Total + Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
        Next Cont_Fila
        'SE EDITA EL ESTATUS
        Mi_SQL = " SELECT * FROM  Alm_Salidas_Almacen WHERE No_Salida= '" & Format(Txt_No_Salida.Text, "0000000000") & "'"
        Mi_SQL = Mi_SQL & " AND Estatus='RECEPCION'"
        Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Ope_Entrada.EOF Then
            With Rs_Ope_Entrada
                .Edit
                    .rdoColumns("Estatus") = "CANCELADA"
                    .rdoColumns("Comentarios") = Comentarios_Cancelacion
                    .rdoColumns("Usuario_Modifico") = Trim(Nombre_Usuario)
                    .rdoColumns("Fecha_Modifico") = Now
                .Update
            End With
            Rs_Ope_Entrada.Close
            
            'SE REGISTRA LA FACTURA EN EL SISTEMA
            Set Rs_Alta_Tmp_Facturas_Proveedores = Conectar_Ayudante.Recordset_Agregar("Tmp_Proveedores_Facturas")
            With Rs_Alta_Tmp_Facturas_Proveedores
                .AddNew
                     No_Control = Format(Conectar_Ayudante.Maximo_Catalogo("Tmp_Proveedores_Facturas", "No_Control"), "0000000000")
                    .rdoColumns("No_Control") = No_Control
                    .rdoColumns("Fecha_Recepcion") = Format(Now, "MM/dd/yyyy")
                    .rdoColumns("Subtotal") = 0
                    .rdoColumns("IVA") = 0
                    .rdoColumns("Total") = 0
                    .rdoColumns("Flete") = 0
                    .rdoColumns("Total_Factura") = 0
                    .rdoColumns("Cancelada") = "NO"
                    .rdoColumns("Facturar") = "NO"
                    .rdoColumns("Aplicada") = "NO"
                    .rdoColumns("Comentarios") = Trim(UCase("ENTRADA HECHA POR CANCELACION DE SALIDA No.")) & Format(Txt_No_Salida.Text, "0000000000")
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
            Rs_Alta_Tmp_Facturas_Proveedores.Close
            
            'SE DAN DE ALTA LOS DATOS GENERALES DE LA ENTRADA
            Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas")
            With Rs_Ope_Entrada
                .AddNew
                    .rdoColumns("Proveedor_ID") = "00046"
                    .rdoColumns("No_Control") = No_Control
                    .rdoColumns("Fecha_Factura") = Format(Now, "MM/dd/yyyy")
                    .rdoColumns("Fecha_Recepcion_Factura") = Format(Now, "MM/dd/yyyy")
                    Entrada_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Alm_Entradas", "Entrada_ID"), "0000000000")
                    .rdoColumns("Entrada_ID") = Entrada_ID
                    .rdoColumns("Tipo_Entrada") = "AJUSTE"
                    .rdoColumns("Estatus") = "AJUSTE X CANCELACION"
                    .rdoColumns("Observaciones") = UCase("ENTRADA HECHA POR CANCELACION DE SALIDA No.") & Format(Txt_No_Salida.Text, "0000000000")
                    .rdoColumns("Usuario_Creo") = Trim(Nombre_Usuario)
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
            Rs_Ope_Entrada.Close
                   
            For Cont_Fila = 1 To Grid_Detalle_Salidas.Rows - 1 Step 1
                'SE HACE LA ENTRADA DE PRODUCTO AL ALMACEN
                Set Rs_Ope_Entrada_Detalles = Conectar_Ayudante.Recordset_Agregar("Alm_Entradas_Detalles")
                With Rs_Ope_Entrada_Detalles
                    .AddNew
                        Rs_Ope_Entrada_Detalles.rdoColumns("Entrada_ID") = Entrada_ID
                        Rs_Ope_Entrada_Detalles.rdoColumns("Producto_ID") = Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 4)
                        Mi_SQL = "SELECT * FROM Cat_Productos WHERE Producto_ID='" & Format(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 4), "00000") & "'"
                        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            Rs_Ope_Entrada_Detalles.rdoColumns("Descripcion") = Rs_Consulta_Producto!Nombre
                            Rs_Ope_Entrada_Detalles.rdoColumns("Costo") = Val(Rs_Consulta_Producto!Costo)
                            Rs_Ope_Entrada_Detalles.rdoColumns("Importe") = (Val(Rs_Consulta_Producto!Costo) + (Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0)) * (Val(Rs_Consulta_Producto!Costo) * Val(PG_Retencion_IVA)))) * Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Cantidad") = Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Faltante") = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0), ","))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Impuesto") = Val(Rs_Consulta_Producto!Impuesto)
                            Rs_Ope_Entrada_Detalles.rdoColumns("IVA") = Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0)) * (Val(Rs_Consulta_Producto!Costo) * Val(PG_Retencion_IVA))
                            Rs_Ope_Entrada_Detalles.rdoColumns("Estatus") = "AJUSTE X CANCELACION"
                        Rs_Consulta_Producto.Close
                    .Update
                End With
                Rs_Ope_Entrada_Detalles.Close
                'ACTUALIZA EXISTENCIA EN EL CATALOGO DE PRODUCTOS
                If Cmb_Tipo_Entrada.Text = "VENTA" Then
                    Mi_SQL = " SELECT * FROM Cat_Productos "
                    Mi_SQL = Mi_SQL & " WHERE Producto_ID ='" & Format(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 4), "00000") & "' "
                    Set Rs_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Cat_Productos
                        .Edit
                            If Not IsNull(.rdoColumns("Existencia")) Then
                                .rdoColumns("Existencia") = Val(.rdoColumns("Existencia")) + Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            Else
                                .rdoColumns("Existencia") = 0 + Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            End If
                            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                            .rdoColumns("Fecha_Modifico") = Now
                        .Update
                    End With
                    Rs_Cat_Productos.Close
                End If
            Next Cont_Fila
        Else
            MsgBox "La salida no se puede cancelar", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            Btn_Nuevo.Caption = "Nuevo"
            Btn_Salir.Visible = True
            Btn_Regresar.Visible = False
            Btn_Cancelar.Visible = False
            Btn_Salir.Visible = True
            Btn_Nuevo.Visible = True
            Exit Sub
        End If
        Conexion_Base.CommitTrans
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Visible = True
        Btn_Regresar.Visible = False
        Btn_Cancelar.Visible = False
        Btn_Salir.Visible = True
        Btn_Nuevo.Visible = True
        MsgBox "Salida de almacen cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Visible = True
        Btn_Regresar.Visible = False
        Btn_Cancelar.Visible = False
        Btn_Salir.Visible = True
        Btn_Nuevo.Visible = True
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    If Err.Number = 482 Then
        MsgBox "Impresion Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        MsgBox Err.Description
    End If
End Sub

''*******************************************************************************
'NOMBRE DE LA FUNCION   : Btn_Elimina_detalle_Click
'DESCRIPCION            : Elimina los productos a los detalles de la salida
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Btn_Elimina_detalle_Click()
Dim Cont_Filas As Integer
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
Dim Año As Integer
Dim Mes As Integer
Dim Dia As Integer
Dim Fecha_Valida As Date
Dim Agregar_Partida As Boolean

On Error GoTo Handler
    
    If Trim(Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 3)) = "PEDIDO" Then
         Grid_Entradas_Productos.Cols = 3
         If Grid_Entradas_Productos.Rows = 0 Then
             'SE AGREGA ENCABEZADO
             Grid_Entradas_Productos.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID"
         End If
         'SE ASIGNAN LOS DETALLES AL GRID
         Grid_Entradas_Productos.AddItem Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 0) & Chr(9) & _
         Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 1) & Chr(9) & _
         Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 2)
         Grid_Entradas_Productos.FixedRows = 1
         'SE ELIMINA DEL GRIN DE PRODUCTOS DE ORDEN DE COMPRA
         If Grid_Detalle_Salidas.Rows > 2 Then
             Grid_Detalle_Salidas.RemoveItem Grid_Detalle_Salidas.RowSel
         Else
             Grid_Detalle_Salidas.Rows = 0
         End If
         'Grid Partidas
        If Grid_Entradas_Productos.Rows > 1 Then
             'Configura el grid
             Grid_Entradas_Productos.ColWidth(0) = 800              'Ancho de las columnas
             Grid_Entradas_Productos.ColWidth(1) = 7390
             Grid_Entradas_Productos.ColAlignment(1) = 1            'Alínea la columna a la derecha
             Grid_Entradas_Productos.ColWidth(2) = 0
             'Pone el setfocus en la primera fila del Grid
             With Grid_Entradas_Productos
                 .Col = 1
                 .Row = 1
                 .ColSel = .Cols - 1
                 .RowSel = 1
                 .TopRow = .Row
                 .SetFocus
             End With
         End If
    Else
        If Grid_Detalle_Salidas.Rows = 2 Then
            Grid_Detalle_Salidas.FixedRows = 0
            'Quita el item del grid
            Grid_Detalle_Salidas.RemoveItem (Grid_Detalle_Salidas.RowSel + 1)
        Else
            If Grid_Detalle_Salidas.Rows > 2 Then
                Grid_Detalle_Salidas.RemoveItem (Grid_Detalle_Salidas.RowSel)
            End If
        End If
    End If
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub Btn_Imprimir_Click()
    If Txt_No_Salida.Text <> "" Then
        Call Imprimir_Salida
        ''Call Imprimir_Remision
    Else
        MsgBox "No hay datos que Imprimir", vbInformation
    End If
End Sub




Private Sub Btn_Nuevo_Click()
    'Prepara la ventana para realizar una nueva salida
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Salir.Visible = False
        Btn_Regresar.Visible = True
        Fra_Datos_Salida.Enabled = True
        Fra_Productos_Orden_Compra.Enabled = True
        Fra_Detalles_Salida.Enabled = True
        Dtp_Fecha_Salida.Value = Now
        Grid_Entradas_Productos.Rows = 0
        Grid_Detalle_Salidas.Rows = 0
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        ''Txt_No_Salida.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Alm_Salidas_Almacen", "No_Salida"), "0000000000")
        Txt_No_Salida.Text = Conectar_Ayudante.Maximo_Catalogo("Alm_Salidas_Almacen", "No_Salida")
        Cmb_Cliente_Salida_Almacen.Text = ""
        Cmb_Orden_Compra.Text = ""
        Cmb_Tipo_Entrada.ListIndex = 0
    Else
        If Cmb_Cliente_Salida_Almacen.ListIndex > -1 Then
            If Cmb_Tipo_Entrada.Text <> "" Then
                Call Captura_Salida_Almacen
            Else
                MsgBox "Seleccione el tipo de salida", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
                Cmb_Tipo_Entrada.SetFocus
            End If
        Else
            MsgBox "Seleccione un cliente", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            Cmb_Cliente_Salida_Almacen.SetFocus
        End If
    End If
End Sub
Private Sub Btn_Regresar_Click()
    If Btn_Regresar.Visible = True Then
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Visible = True
        Btn_Regresar.Visible = False
        Fra_Datos_Salida.Enabled = False
        Fra_Productos_Orden_Compra.Enabled = False
        Fra_Detalles_Salida.Enabled = False
        Dtp_Fecha_Salida.Value = Now
        Grid_Entradas_Productos.Rows = 0
        Grid_Detalle_Salidas.Rows = 0
        Cmb_Cliente_Salida_Almacen.ListIndex = -1
        Cmb_Orden_Compra.ListIndex = -1
        Cmb_Descripcion.ListIndex = -1
        Call Conectar_Ayudante.Limpiar_Textos(Me)
    End If
End Sub
Private Sub Btn_Salir_Click()
    If Btn_Salir.Value = True Then
        Unload Me
    End If
End Sub
''*******************************************************************************
'NOMBRE DE LA FUNCION   : Cmb_Cliente_Salida_Almacen_Click
'DESCRIPCION            : Consulta las ordenes de compra relacionados con el cliente seleccionado
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Cmb_Cliente_Salida_Almacen_Click()
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
Dim Rs_Consulta_ID_Cliente As rdoResultset

    If Cmb_Cliente_Salida_Almacen.ListIndex > -1 Then
        Cmb_Orden_Compra.Clear
        'SE CONSULTAN LAS ORDENES RELACIONADAS  CON EL CLIENTE
        Mi_SQL = " SELECT Distinct Ope_Pedidos.Pedido_ID "
        Mi_SQL = Mi_SQL & " From   Ope_Pedidos,Ope_Pedidos_Detalles "
        Mi_SQL = Mi_SQL & " WHERE  Ope_Pedidos.Cliente_ID ='" & Format(Cmb_Cliente_Salida_Almacen.ItemData(Cmb_Cliente_Salida_Almacen.ListIndex), "00000") & "' "
        Mi_SQL = Mi_SQL & " AND    Ope_Pedidos.Pedido_ID = Ope_Pedidos_Detalles.Pedido_ID "
        Mi_SQL = Mi_SQL & " AND    Ope_Pedidos.Estatus='PENDIENTE' "
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta.EOF Then
            'SE ASIGNAN LAS ORDENES AL COMBO DE ORDEN DE COMPRA
            While Not Rs_Consulta.EOF
                Cmb_Orden_Compra.AddItem Rs_Consulta!Pedido_ID
                Cmb_Orden_Compra.ItemData(Cmb_Orden_Compra.NewIndex) = Rs_Consulta!Pedido_ID
                Rs_Consulta.MoveNext
            Wend
        End If
    End If
    If Cmb_Cliente_Salida_Almacen.ListIndex > -1 Then
        Txt_Cliente_ID.Text = Format(Cmb_Cliente_Salida_Almacen.ItemData(Cmb_Cliente_Salida_Almacen.ListIndex), "00000")
    End If
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Cliente_Salida_Almacen_KeyPress
'DESCRIPCIÓN          : Llena el combo con los clientes
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 11-Enero-2011
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Cliente_Salida_Almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Salida_Almacen, 1, "Nombre")
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Cliente_Salida_Almacen.hwnd, &H14F, True, 0)
    End If
End Sub


''*******************************************************************************
'NOMBRE DE LA FUNCION   : Cmb_Orden_Compra_Click
'DESCRIPCION            : Consulta los detalles de la requisición
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Private Sub Cmb_Orden_Compra_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Detalles_Pedido As rdoResultset

    If Cmb_Orden_Compra.ListIndex > -1 Then
            'Prepara el recordset para consultar los detalles del pedido
            Mi_SQL = "SELECT * FROM Ope_Pedidos_Detalles WHERE  Pedido_ID='" & Format(Cmb_Orden_Compra.Text, "0000000000") & "'"
            Set Rs_Consulta_Detalles_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Grid_Entradas_Productos.Rows = 0
            Grid_Entradas_Productos.Cols = 3
            'Pone el encabezado en las columnas
            Grid_Entradas_Productos.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID"
            'Llenado del grid
            While Not Rs_Consulta_Detalles_Pedido.EOF
                Grid_Entradas_Productos.AddItem Rs_Consulta_Detalles_Pedido!Cantidad & Chr(9) & Rs_Consulta_Detalles_Pedido!Descripcion & Chr(9) & Rs_Consulta_Detalles_Pedido!Producto_ID
                Grid_Entradas_Productos.FixedRows = 1
                Rs_Consulta_Detalles_Pedido.MoveNext
            Wend
            Rs_Consulta_Detalles_Pedido.Close
            If Grid_Entradas_Productos.Rows > 1 Then
                'Configura el grid
                Grid_Entradas_Productos.ColWidth(0) = 800              'Ancho de las columnas
                Grid_Entradas_Productos.ColWidth(1) = 7390
                Grid_Entradas_Productos.ColAlignment(1) = 1            'Alínea la columna a la derecha
                Grid_Entradas_Productos.ColWidth(2) = 0
                'Pone el setfocus en la primera fila del Grid
                With Grid_Entradas_Productos
                    .Col = 1
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                    .SetFocus
                End With
            End If
        End If
End Sub
Private Sub Form_Load()
    Me.Height = 8520
    Me.Width = 9225
    Me.Top = 100
    Me.Left = 100
    Me.Left = (Screen.Width - Me.Width) / 2
    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Salida_Almacen, 1, " Status='A' AND Nombre")
    Call Cmb_Descripcion_KeyPress(13)
End Sub
''*******************************************************************************
'NOMBRE DE LA FUNCION   : Captura_Salida_Almacen
'DESCRIPCION            : Captura la salida del almacen
'PARAMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 11-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACION     :
'*******************************************************************************
Public Sub Captura_Salida_Almacen()
Dim Rs_Ope_Entrada As rdoResultset
Dim Rs_Ope_Entrada_Detalles As rdoResultset
Dim Rs_Alm_Entradas_Almacen As rdoResultset
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Pedido As rdoResultset
Dim Rs_Alta_Tmp_Facturas_Proveedores As rdoResultset
Dim Mi_SQL As String
Dim Cont_Fila As Integer
Dim No_Salida As String
Dim Cantidad_Total As Double
Dim Rs_Ope_Pedido_detalles As rdoResultset

On Error GoTo Handler
    Conexion_Base.BeginTrans
    
    Cantidad_Total = 0
    No_Salida = Format(Txt_No_Salida.Text, "0000000000")
    
    'SE OBTIENE LA CANTIDAD TOTAL DE LA SALIDA
    For Cont_Fila = 1 To Grid_Detalle_Salidas.Rows - 1 Step 1
        Cantidad_Total = Cantidad_Total + Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
    Next Cont_Fila
    
    'SE DAN DE ALTA LOS DATOS GENERALES
    Set Rs_Ope_Entrada = Conectar_Ayudante.Recordset_Agregar("Alm_Salidas_Almacen")
    With Rs_Ope_Entrada
        .AddNew
            .rdoColumns("No_Salida") = Format(Txt_No_Salida.Text, "0000000000")
            .rdoColumns("Cliente_ID") = Format(Cmb_Cliente_Salida_Almacen.ItemData(Cmb_Cliente_Salida_Almacen.ListIndex), "00000")
            If Cmb_Orden_Compra.ListIndex > -1 Then
                .rdoColumns("Pedido_ID") = Format(Cmb_Orden_Compra.ItemData(Cmb_Orden_Compra.ListIndex), "0000000000")
            End If
            .rdoColumns("Estatus") = "RECEPCION"
            .rdoColumns("Tipo_Salida") = Trim(Cmb_Tipo_Entrada.Text)
            .rdoColumns("Comentarios") = UCase(Txt_Observaciones.Text)
            .rdoColumns("Cantidad_Total") = Cantidad_Total
            .rdoColumns("Fecha_Salida") = Format(Dtp_Fecha_Salida.Value, "MM/dd/yyyy")
            .rdoColumns("Usuario_Creo") = Trim(Nombre_Usuario)
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Ope_Entrada.Close
    
    If Cmb_Orden_Compra.ListIndex > -1 Then
        'ACTUALIZA EL PEDIDO
        Mi_SQL = " SELECT * FROM Ope_Pedidos "
        Mi_SQL = Mi_SQL & " WHERE Pedido_ID ='" & Format(Cmb_Orden_Compra.ItemData(Cmb_Orden_Compra.ListIndex), "0000000000") & "' "
        Mi_SQL = Mi_SQL & " AND Estatus ='PENDIENTE' "
        Set Rs_Ope_Pedido_detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Ope_Pedido_detalles.EOF Then
            With Rs_Ope_Pedido_detalles
                .Edit
                    .rdoColumns("Estatus") = "SURTIDO"
                .Update
            End With
        End If
        Rs_Ope_Pedido_detalles.Close
    End If
    
    'SE DAN DE ALTA LOS DETALLES
    For Cont_Fila = 1 To Grid_Detalle_Salidas.Rows - 1 Step 1
        Set Rs_Ope_Entrada_Detalles = Conectar_Ayudante.Recordset_Agregar("Alm_Salidas_Almacen_Detalles")
        With Rs_Ope_Entrada_Detalles
            .AddNew
                Rs_Ope_Entrada_Detalles.rdoColumns("No_Salida") = Format(Txt_No_Salida.Text, "0000000000")
                Rs_Ope_Entrada_Detalles.rdoColumns("Descripcion") = Trim(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 1))
                Rs_Ope_Entrada_Detalles.rdoColumns("Cantidad") = Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0)
                Rs_Ope_Entrada_Detalles.rdoColumns("Producto_ID") = Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 2)
                Rs_Ope_Entrada_Detalles.rdoColumns("Facturado") = "NO"
                'SI ES VENTA ACTUALIZA LA EXISTENCIA EN ALMACEN
                If Cmb_Tipo_Entrada.Text = "VENTA" Then
                    'ACTUALIZA EXISTENCIA EN EL CATALOGO DE PRODUCTOS
                    Mi_SQL = " SELECT * FROM Cat_Productos "
                    Mi_SQL = Mi_SQL & " WHERE Producto_ID ='" & Format(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 2), "00000") & "' "
                    Set Rs_Cat_Productos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Cat_Productos
                        .Edit
                            If Not IsNull(.rdoColumns("Existencia")) Then
                                .rdoColumns("Existencia") = Val(.rdoColumns("Existencia")) - Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            Else
                                .rdoColumns("Existencia") = 0 - Val(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 0))
                            End If
                            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                            .rdoColumns("Fecha_Modifico") = Now
                        .Update
                    End With
                    Rs_Cat_Productos.Close
                End If
                If Cmb_Orden_Compra.ListIndex > -1 Then
                    'ACTUALIZA EL PEDIDO
                    Mi_SQL = " SELECT * FROM Ope_Pedidos_Detalles "
                    Mi_SQL = Mi_SQL & " WHERE Pedido_ID ='" & Format(Cmb_Orden_Compra.ItemData(Cmb_Orden_Compra.ListIndex), "0000000000") & "' "
                    Mi_SQL = Mi_SQL & " AND Producto_ID ='" & Trim(Grid_Detalle_Salidas.TextMatrix(Cont_Fila, 2)) & "' "
                    Mi_SQL = Mi_SQL & " AND Estatus ='PENDIENTE' "
                    Set Rs_Ope_Pedido_detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Ope_Pedido_detalles.EOF Then
                        With Rs_Ope_Pedido_detalles
                            .Edit
                                .rdoColumns("Estatus") = "SURTIDO"
                            .Update
                        End With
                    End If
                    Rs_Ope_Pedido_detalles.Close
                End If
                
            .Update
        End With
        Rs_Ope_Entrada_Detalles.Close
    Next Cont_Fila
        
    Conexion_Base.CommitTrans
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Salir.Visible = True
    Btn_Regresar.Visible = False
    MsgBox "Salida de almacen dada de alta dada de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
''    If MsgBox("Salida de almacen dada de alta dada de alta" & Chr(13) & "¿Desea enviarla a impresión?", vbQuestion + vbYesNo) = vbYes Then
''        Call Imprimir_Salida
''        ''Call Imprimir_Remision
''    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    If Err.Number = 482 Then
        MsgBox "Impresion Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        MsgBox Err.Description
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprimir_Salida
'DESCRIPCIÓN            : Envia a impresion la salida generada o consultada
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 12-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Imprimir_Salida()
Dim Rs_Consulta_Cat_Materias_Primas As rdoResultset 'Consulta el costo promedio del producto}
Dim Fila As Integer                           'Contador de filas en el grid
Dim Costo As Double                  'Indica el costo promedio del producto
Dim Importe_Costo_Promedio As Double          'Indica el importe de costo promedio del producto
Dim Total_Importe_Costo_Promedio As Double    'Indica el total de la salida
Dim Partidas As Integer                       'Indica la cantidad de partidas que tiene la salida
Dim Total_Productos As Double                 'Indica el total de productos a los cuales se les da salida
Dim Impresora As String            'Tomna el nombre la impresora
Dim Mi_Impresora As Printer        'Toma el nombre de la impresora
Dim Ubicacion_Impresora As String  'Toma el valor de la ubicacion dela impresora
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset

On Error GoTo Handler
    Printer.Font = "COURIER NEW"
    Printer.FontSize = 7
    Printer.FontSize = 9
    Printer.Print
    Printer.Print
    Printer.Print "*********************************************************************************************"
    Printer.Print Conectar_Ayudante.Centrar_Texto("ALCOHOLERA DEL CENTRO, S.A. de C.V.", Len("*********************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("Av. Juan Jose Torres landa No. 636 Col. Independencia C.P.36559 Irapuato, Gto.", Len("*********************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("Tel, Y Fax 01(462)633-20-32 633-20-33 y 633-20-34", Len("*********************************************************************************************"))
    Printer.Print Conectar_Ayudante.Centrar_Texto("alcesa@prodigy.net.mx     www.alcesa.com.mx", Len("*********************************************************************************************"))
    Printer.Print
    Printer.Print Spc(2); Format(Now, "HH:mm:ss"); Conectar_Ayudante.Centrar_Texto("SALIDA DE ALMACEN", Len("*********************************************************************************************"))
    Printer.Print "_____________________________________________________________________________________________"
    Printer.Print
    Printer.Print " No. Salida   :  "; Txt_No_Salida.Text
    Printer.Print " Fecha Salida :  "; Format(Dtp_Fecha_Salida.Value, "dd/MMM/yyyy")
    Mi_SQL = " SELECT * FROM Cat_clientes WHERE Cliente_ID='" & Format(Cmb_Cliente_Salida_Almacen.ItemData(Cmb_Cliente_Salida_Almacen.ListIndex), "00000") & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta.EOF Then
        Printer.Print " Cliente        : "; Rs_Consulta!Nombre
        Printer.Print " RFC            : "; Rs_Consulta!RFC
        Printer.Print " Dirección      : "; Rs_Consulta!Direccion
        Printer.Print " Colonia        : "; Rs_Consulta!Colonia
        Printer.Print " CP             : "; Rs_Consulta!CP
        Printer.Print " Ciudad         : "; Rs_Consulta!Ciudad
    End If
    Rs_Consulta.Close
    Printer.Print
    Printer.Print "*********************************************************************************************"
    Printer.Print "  Cantidad           Descripcion                                                             "
    Printer.Print "_____________________________________________________________________________________________"
    Printer.Print
    Importe_Costo_Promedio = 0
    Total_Importe_Costo_Promedio = 0
    Total_Productos = 0
    Partidas = Grid_Detalle_Salidas.Rows - 1
    For Fila = 1 To Grid_Detalle_Salidas.Rows - 1
        Costo = 0
        Importe_Costo_Promedio = 0
        Total_Productos = Total_Productos + Val(Grid_Detalle_Salidas.TextMatrix(Fila, 0))
        ''Costo = Val(Grid_Detalle_Salidas.TextMatrix(Fila, 5))
        Total_Importe_Costo_Promedio = Total_Importe_Costo_Promedio + Importe_Costo_Promedio
        Printer.Print Spc(3); Mid(Grid_Detalle_Salidas.TextMatrix(Fila, 0), 1, 15); _
                      Spc(18 - Len(Mid(Grid_Detalle_Salidas.TextMatrix(Fila, 0), 1, 15))); Mid(Grid_Detalle_Salidas.TextMatrix(Fila, 1), 1, 55)
                      ''Spc(39 - Len(Mid(Grid_Detalle_Salidas.TextMatrix(Fila, 1), 1, 34))); Conectar_Ayudante.Alinea_Derecha(Format(Grid_Detalle_Salidas.TextMatrix(Fila, 2), "#0.00"), 9); _
                      ''Spc(9); Conectar_Ayudante.Alinea_Derecha(Format(Grid_Detalle_Salidas.TextMatrix(Fila, 3), "#0.00"), 9)
    Next Fila
    Printer.Print " "
    Printer.Print "                                                           -------------------------"
    ''Printer.Print "                                                 Costo Total     :  $     "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Importe_Costo_Promedio, "#,##0.00"), 8)
    Printer.Print "                                                 Total Productos :    "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Productos, "#,##0"), 12)
    Printer.Print "                                                 Partidas        :        "; Conectar_Ayudante.Alinea_Derecha(Format(Grid_Detalle_Salidas.Rows - 1, "#,##0"), 8)
    Printer.Print "---------------------------------------------------------------------------------------------"
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "       _______________________________                 ______________________________"
    Printer.Print "          NOMBRE Y FIRMA (ENTREGO)                         NOMBRE Y FIRMA (RECIBIO)"
    Printer.Print
    Printer.Print
    Printer.Print "                                  _______________________________"
    Printer.Print "                                      NOMBRE Y FIRMA (AUTORIZO)"
    Printer.EndDoc
    MsgBox "Salida enviada a impresion", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
Handler:
    For Each Er In rdoErrors
        If Mid(Er, 1, 5) = "01S02" Then MsgBox "No se encontro la impresora", vbCritical
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Consulta_Salida
'DESCRIPCIÓN            : Consulta la salida
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Consulta_Salida(No_Salida As String)
Dim Rs_Consultar_Salida As rdoResultset
Dim Rs_Consultar_Salida_Detalles As rdoResultset
Dim Mi_SQL As String

On Error GoTo Handler

    Grid_Detalle_Salidas.Rows = 0
    Grid_Detalle_Salidas.Cols = 6
    'SE AGREGA ENCABEZADO
    Grid_Detalle_Salidas.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Precio" & Chr(9) & "Importe" & Chr(9) & "Producto ID" & Chr(9) & "IVA"
    'SE CONSULTAN LOS DATOS GENERALES
    Mi_SQL = " SELECT * FROM Alm_Salidas_Almacen "
    Mi_SQL = Mi_SQL & " WHERE No_Salida='" & No_Salida & "'"
    Set Rs_Consultar_Salida = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consultar_Salida.EOF Then
        Txt_No_Salida.Text = Rs_Consultar_Salida!No_Salida
        Call Conectar_Ayudante.Asigna_Item_Combo(Rs_Consultar_Salida!Cliente_ID, Cmb_Cliente_Salida_Almacen)
        If Not IsNull(Rs_Consultar_Salida!Pedido_ID) Then Cmb_Orden_Compra.Text = Rs_Consultar_Salida!Pedido_ID
        Dtp_Fecha_Salida.Value = Rs_Consultar_Salida!Fecha_Salida
        Txt_Observaciones.Text = Rs_Consultar_Salida!Comentarios
        Cmb_Tipo_Entrada.Text = Rs_Consultar_Salida!Tipo_Salida
        'SE CONSULTAN LOS DETALLES
        Mi_SQL = " SELECT * FROM Alm_Salidas_Almacen_Detalles "
        Mi_SQL = Mi_SQL & " WHERE No_Salida='" & No_Salida & "'"
        Set Rs_Consultar_Salida_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consultar_Salida_Detalles.EOF Then
            'SE ASIGNAN LOS DETALLES AL GRID
            While Not Rs_Consultar_Salida_Detalles.EOF
                Grid_Detalle_Salidas.AddItem Rs_Consultar_Salida_Detalles!Cantidad & Chr(9) & _
                Rs_Consultar_Salida_Detalles!Descripcion & Chr(9) & _
                Rs_Consultar_Salida_Detalles!Precio_Venta & Chr(9) & _
                Rs_Consultar_Salida_Detalles!Importe & Chr(9) & _
                Rs_Consultar_Salida_Detalles!Producto_ID & Chr(9) & _
                Rs_Consultar_Salida_Detalles!Impuesto
                Grid_Detalle_Salidas.FixedRows = 1
                Rs_Consultar_Salida_Detalles.MoveNext
              Wend
            'Grid Partidas
            If Grid_Detalle_Salidas.Rows > 1 Then
                'Configura el grid
                Grid_Detalle_Salidas.ColWidth(0) = 800              'Ancho de las columnas
                Grid_Detalle_Salidas.ColWidth(1) = 7390
                Grid_Detalle_Salidas.ColAlignment(1) = 1            'Alínea la columna a la derecha
                Grid_Detalle_Salidas.ColWidth(2) = 0
                Grid_Detalle_Salidas.ColWidth(3) = 0
                Grid_Detalle_Salidas.ColWidth(4) = 0
                Grid_Detalle_Salidas.ColWidth(5) = 0
                'Pone el setfocus en la primera fila del Grid
                With Grid_Detalle_Salidas
                    .Col = 1
                    .Row = 1
                    .ColSel = .Cols - 1
                    .RowSel = 1
                    .TopRow = .Row
                End With
             End If
        End If
        Rs_Consultar_Salida_Detalles.Close
        Btn_Salir.Visible = False
        Btn_Cancelar.Visible = True
        Fra_Detalles_Salida.Enabled = False
    Else
        MsgBox "Salida no encontrada"
    End If
    Rs_Consultar_Salida.Close
    Call Calcula_Total_salida
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Calcula_Total_salida
'DESCRIPCIÓN            : calcula el total de producto que va a salir
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 13-Nov-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Calcula_Total_salida()
Dim Fila As Integer
    
    Txt_Total_Salida.Text = ""
    For Fila = 1 To Grid_Detalle_Salidas.Rows - 1 Step 1
        Txt_Total_Salida.Text = Val(Txt_Total_Salida.Text) + Val(Grid_Detalle_Salidas.TextMatrix(Fila, 0))
    Next
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Descripcion_KeyPress
'DESCRIPCIÓN          : Llena el combo con los productos
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Descripcion, 1, "Nombre")
    End If
End Sub



'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Txt_Precio_Change
'DESCRIPCIÓN          : Modifica, el importe
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 4-Enero-2011
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************'
Private Sub Txt_Precio_Change()
    Txt_Importe.Text = ""
    Txt_Importe.Text = Val(Txt_Precio.Text) * Val(Txt_Cantidad.Text)
End Sub


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprimir_Remision
'DESCRIPCIÓN            : Imprime la Remision o vale de salida
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 06-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Imprimir_Remision()
Dim Mi_SQL As String                            'Cadena para general las consultas
Dim Rs_Formato  As rdoResultset                 'Manejo de registro para la tabla Cfg_Formatos
Dim Rs_Formato_Generales  As rdoResultset       'Manejo de registro para la tabla Cfg_Formatos_Generales
Dim Rs_Formato_Detalles As rdoResultset         'Manejo de registro para la tabla Cfg_Formatos_Detalles
Dim Rs_Generales_Salida As rdoResultset
Dim Rs_Detalles_Salida As rdoResultset
Dim Clave As String
Dim Cantidad As Integer
Dim ARTICULO As String
Dim Nombre As String
Dim DOMICILIO As String
Dim Ciudad As String
Dim Fecha As String

        
        
    On Error GoTo Handler
    
    'Consulta para la configuración de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Set Rs_Formato = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta para la configuración general de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'General'"
    Set Rs_Formato_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta para la configuración a detalle de facturas
    Mi_SQL = "SELECT * FROM Cfg_Formatos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Nombre = 'REMISION'"
    Mi_SQL = Mi_SQL & " AND Tipo = 'Detalle' ORDER BY Campo"
    Set Rs_Formato_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Impresión de la factura
    If Not Rs_Formato.EOF Then
    
        'Configura la fuente de la factura para generales
        With Rs_Formato
            Printer.ScaleMode = vbCentimeters
            Printer.FontSize = .rdoColumns("Tamaño_Generales")
            Printer.Font = .rdoColumns("Letra_Generales")
            If .rdoColumns("Estilo_Generales") = "Negrita" Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
        End With
        
        'SE CONSULTAN LOS DATOS GENERALES DE LA SALIDA
        Mi_SQL = " SELECT Alm_Salidas_Almacen.*,Cat_Clientes.* FROM Alm_Salidas_Almacen,Cat_Clientes WHERE   Alm_Salidas_Almacen.No_Salida = '" & Format(Txt_No_Salida.Text, "0000000000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Almacen.Cliente_ID= Cat_Clientes.Cliente_ID"
        Set Rs_Generales_Salida = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Imprime los datos del cliente
        With Rs_Formato_Generales
            While Not .EOF
                    Printer.CurrentX = .rdoColumns("X")
                    Printer.CurrentY = .rdoColumns("Y")
                    Longitud = .rdoColumns("Longitud")
                    If .rdoColumns("Campo") = "NOMBRE" Then
                       Printer.Print Mid(Rs_Generales_Salida!Nombre, 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "DOMICILIO" Then
                       Printer.Print Mid(Rs_Generales_Salida.rdoColumns("Direccion") & " " & Rs_Generales_Salida.rdoColumns("Colonia"), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "CIUDAD" Then
                       Printer.Print Mid(Rs_Generales_Salida.rdoColumns("Ciudad") & "," & Rs_Generales_Salida.rdoColumns("Estado"), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "FECHA" Then
                       Printer.Print Format(Rs_Generales_Salida.rdoColumns("Fecha_Salida"), "dd/MM/yyyy")
                    End If
                    If .rdoColumns("Campo") = "NOMBRE_ALMACEN" Then
                       Printer.Print Mid(UCase(Txt_Nombre_Almacenista.Text), 1, Longitud)
                    End If
                    If .rdoColumns("Campo") = "NOMBRE_RECIBE" Then
                       Printer.Print Mid(UCase(Txt_Recibe.Text), 1, Longitud)
                    End If
                .MoveNext
            Wend
        End With
        Rs_Generales_Salida.Close
        
        'Configura la fuente de la factura para detalles
        With Rs_Formato
            Printer.FontSize = .rdoColumns("Tamaño_Detalles")
            Printer.Font = .rdoColumns("Letra_Detalles")
            If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
        End With
        
        'Consulta de la tabla Adm_Descripcion_Facturas con el número de facturas
        Mi_SQL = "SELECT Alm_Salidas_Almacen_Detalles.*,Cat_Productos.Clave FROM Cat_Productos, Alm_Salidas_Almacen_Detalles"
        Mi_SQL = Mi_SQL & " WHERE No_Salida = '" & Txt_No_Salida.Text & "'"
        Mi_SQL = Mi_SQL & " AND  Cat_Productos.Producto_ID = Alm_Salidas_Almacen_Detalles.Producto_ID "
        Set Rs_Detalles_Salida = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Configura la fuente para empresión de los detalles
        If Not Rs_Formato.EOF Then
            With Rs_Formato
                Printer.ScaleMode = vbCentimeters
                Printer.FontSize = .rdoColumns("Tamaño_Detalles")
                Printer.Font = .rdoColumns("Letra_Detalles")
                If .rdoColumns("Estilo_Detalles") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
                Salto = .rdoColumns("Separacion_Detalles")
            End With
            Cont_Renglon = 0
            'Imprime la factura
            While Not Rs_Detalles_Salida.EOF
                Cont_Renglon = Cont_Renglon + Salto
                While Not Rs_Formato_Detalles.EOF
                    Printer.CurrentX = Rs_Formato_Detalles.rdoColumns("X")
                    Printer.CurrentY = Rs_Formato_Detalles.rdoColumns("Y") + Cont_Renglon
                    Longitud = Rs_Formato_Detalles.rdoColumns("Longitud")
                    If Rs_Formato_Detalles.rdoColumns("Campo") = "CLAVE" Then
                        Printer.Print Trim(Rs_Detalles_Salida.rdoColumns("Clave"))
                    End If
                    If Trim(Rs_Formato_Detalles.rdoColumns("Campo")) = "CANTIDAD" Then
                        Printer.Print Conectar_Ayudante.Alinea_Derecha(Trim(Rs_Detalles_Salida.rdoColumns("Cantidad")), 7)
                    End If
                    If Rs_Formato_Detalles.rdoColumns("Campo") = "ARTICULO" Then
                        Printer.Print Trim(Rs_Detalles_Salida.rdoColumns("Descripcion"))
                    End If
                    Rs_Formato_Detalles.MoveNext
                Wend
                Rs_Formato_Detalles.MoveFirst
                Rs_Detalles_Salida.MoveNext
            Wend
            Rs_Detalles_Salida.Close
        End If
        Printer.EndDoc
    End If
    Rs_Formato.Close
    Rs_Formato_Generales.Close
    Rs_Formato_Detalles.Close
    MsgBox "Remisión enviada a Impresión", vbInformation
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub Grid_Detalle_Salidas_Click()
    If Grid_Detalle_Salidas.Rows > 0 Then
        Grid_Detalle_Salidas.SelectionMode = flexSelectionFree
        Grid_Detalle_Salidas.Refresh
        'En notas de cargo no utiliza el campo de modificacion de cantidad
        If Val(Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 0)) > 0 Then
            If Grid_Detalle_Salidas.Rows > 1 Then
                Txt_Modifica_Cantidad.Text = Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 0)
            End If
            If Grid_Detalle_Salidas.Rows = 1 Then
                Txt_Modifica_Cantidad.Text = Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 0)
            End If
            If Grid_Detalle_Salidas.Rows <= 1 Then Exit Sub
            If ((Grid_Detalle_Salidas.Col = 0)) Then
                Call Mover_Control_Grid_TextBox(Grid_Detalle_Salidas, Txt_Modifica_Cantidad)
            Else
                Txt_Modifica_Cantidad.Visible = False
            End If
        Else
            If Grid_Detalle_Salidas.ColSel = 2 Then
                MsgBox "La partida no tiene cantidad"
                Txt_Modifica_Cantidad.Visible = False
            Else
               If Grid_Detalle_Salidas.ColSel = 1 Then
                    Txt_Modifica_Cantidad.Text = Grid_Detalle_Salidas.TextMatrix(Grid_Detalle_Salidas.RowSel, 1)
                    If Grid_Detalle_Salidas.Rows <= 1 Then Exit Sub
                    Call Mover_Control_Grid_TextBox(Grid_Detalle_Salidas, Txt_Modifica_Cantidad)
                    Txt_Modifica_Cantidad.Visible = True
               End If
            End If
        End If
    End If
End Sub

Private Sub Grid_Detalle_Salidas_EnterCell()
    If (Grid_Detalle_Salidas.Col = 0) And Grid_Detalle_Salidas.Rows > 1 Then
        Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Detalle_Salidas, Txt_Modifica_Cantidad)
    End If
End Sub


Private Sub Grid_Detalle_Salidas_LeaveCell()
    Grid_Detalle_Salidas.CellBackColor = vbWhite
End Sub

Private Sub Txt_Modifica_Cantidad_Change()
With Grid_Detalle_Salidas
        If .RowSel > 0 Then
            If .ColSel = 0 Then
                If Val(Txt_Modifica_Cantidad.Text) >= 0 Then
                    .TextMatrix(.RowSel, 0) = Txt_Modifica_Cantidad.Text
                End If
            Else
                Txt_Modifica_Cantidad.Visible = False
            End If
        End If
    End With
End Sub


Private Sub Txt_Modifica_Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
        If KeyCode > 37 Then Grid_Detalle_Salidas.SetFocus
        If KeyCode = 37 Then
            If Txt_Modifica_Cantidad.SelStart = 0 Then
                Grid_Detalle_Salidas.SetFocus
            End If
        End If
        If Grid_Detalle_Salidas.Row > 1 Then
            If KeyCode = 38 Then Grid_Detalle_Salidas.Row = Grid_Detalle_Salidas.RowSel - 1
            If KeyCode = 40 Then
                If Grid_Detalle_Salidas.Row < Grid_Detalle_Salidas.Rows - 1 Or Grid_Detalle_Salidas.Row = 1 Then
                     Grid_Detalle_Salidas.Row = Grid_Detalle_Salidas.RowSel + 1
                Else
                    Txt_Modifica_Cantidad.Visible = False
                    Exit Sub
                End If
            End If
            If Grid_Detalle_Salidas.Col = 9 Then
                If KeyCode = 39 Then Grid_Detalle_Salidas.Col = Grid_Detalle_Salidas.ColSel + 1
            End If
        Else
            If KeyCode = 40 Then
                If Grid_Detalle_Salidas.Row < Grid_Detalle_Salidas.Rows - 1 Or Grid_Detalle_Salidas.Row <> 1 Then
                    Grid_Detalle_Salidas.Row = Grid_Detalle_Salidas.RowSel + 1
                Else
                    Txt_Modifica_Cantidad.Visible = False
                    Exit Sub
                End If
            End If
            If Grid_Detalle_Salidas.Col = 2 Then
                If KeyCode = 39 Then Grid_Detalle_Salidas.Col = Grid_Detalle_Salidas.ColSel + 1
            End If
        End If
        If Txt_Modifica_Cantidad.Visible = True Then
            Txt_Modifica_Cantidad.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End If
    If KeyCode = 13 Then
        Txt_Modifica_Cantidad.Visible = False
    End If
End Sub

