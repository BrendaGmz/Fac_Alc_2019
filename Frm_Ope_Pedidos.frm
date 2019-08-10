VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Ope_Pedidos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pedidos Clientes"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10725
   Begin VB.PictureBox Pic_Pedidos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Frame Fra_Detalle_Pedido 
         BackColor       =   &H8000000E&
         Caption         =   "Detalle Pedido"
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
         Height          =   3900
         Left            =   120
         TabIndex        =   42
         Top             =   2415
         Width           =   10575
         Begin VB.CommandButton Btn_Eliminar 
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
            Height          =   390
            Left            =   9210
            Picture         =   "Frm_Ope_Pedidos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3450
            Width           =   1230
         End
         Begin VB.Frame Fra_Comentarios 
            BackColor       =   &H00FFFFFF&
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
            Height          =   570
            Left            =   120
            TabIndex        =   45
            Top             =   3270
            Width           =   9090
            Begin VB.TextBox Txt_Comentarios 
               Height          =   330
               Left            =   75
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   195
               Width           =   8955
            End
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Detalle_Pedido 
            Height          =   2460
            Left            =   105
            TabIndex        =   16
            Top             =   810
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4339
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
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
            Left            =   9210
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Ope_Pedidos.frx":32B2
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   390
            Width           =   1230
         End
         Begin VB.TextBox Txt_Cantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   825
         End
         Begin VB.ComboBox Cmb_Descripcion 
            Height          =   315
            Left            =   945
            TabIndex        =   14
            Top             =   480
            Width           =   8265
         End
         Begin VB.Label Lbl_Descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Descripción"
            Height          =   195
            Left            =   1020
            TabIndex        =   44
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Lbl_Cantidad 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   165
            TabIndex        =   43
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.CommandButton Btn_Modificar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modificar"
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
         Left            =   2400
         Picture         =   "Frm_Ope_Pedidos.frx":6568
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "M"
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
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
         Left            =   9345
         Picture         =   "Frm_Ope_Pedidos.frx":9C99
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "A"
         Top             =   6360
         UseMaskColor    =   -1  'True
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
         Left            =   120
         Picture         =   "Frm_Ope_Pedidos.frx":D398
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "A"
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Imprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir"
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
         Left            =   4680
         Picture         =   "Frm_Ope_Pedidos.frx":108CF
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "A"
         Top             =   6345
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buscar"
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
         Left            =   7080
         Picture         =   "Frm_Ope_Pedidos.frx":13D95
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "C"
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Datos_Cliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del Cliente"
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
         Height          =   2385
         Left            =   120
         TabIndex        =   29
         Top             =   15
         Width           =   7005
         Begin VB.TextBox Txt_Cliente_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   825
            MaxLength       =   15
            TabIndex        =   49
            Top             =   225
            Width           =   2370
         End
         Begin VB.TextBox Txt_RFC 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   855
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1275
            Width           =   2370
         End
         Begin VB.TextBox Txt_Direccion 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   855
            TabIndex        =   2
            Top             =   915
            Width           =   6075
         End
         Begin VB.TextBox Txt_Colonia 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   855
            TabIndex        =   4
            Top             =   1635
            Width           =   2370
         End
         Begin VB.TextBox Txt_CP 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   855
            TabIndex        =   5
            Top             =   1995
            Width           =   2370
         End
         Begin VB.TextBox Txt_Telefono 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   4380
            TabIndex        =   6
            Top             =   1275
            Width           =   2550
         End
         Begin VB.TextBox Txt_Limite_Credito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   300
            Left            =   4380
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1995
            Width           =   2550
         End
         Begin VB.TextBox Txt_Descuento_Cliente 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6210
            TabIndex        =   30
            Top             =   2580
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox Txt_Ciudad 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   4380
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1635
            Width           =   2550
         End
         Begin VB.ComboBox Cmb_Cliente 
            Height          =   315
            Left            =   855
            TabIndex        =   1
            Top             =   600
            Width           =   6075
         End
         Begin VB.TextBox Txt_Cliente 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   31
            Top             =   607
            Visible         =   0   'False
            Width           =   6075
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente ID"
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre"
            Height          =   195
            Left            =   90
            TabIndex        =   41
            Top             =   660
            Width           =   555
         End
         Begin VB.Label Lbl_RFC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "RFC"
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   1328
            Width           =   315
         End
         Begin VB.Label Lbl_Direccion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Direccion"
            Height          =   195
            Left            =   90
            TabIndex        =   39
            Top             =   968
            Width           =   675
         End
         Begin VB.Label Lbl_Colonia 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Colonia"
            Height          =   195
            Left            =   90
            TabIndex        =   38
            Top             =   1688
            Width           =   525
         End
         Begin VB.Label Lbl_CP 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "C.P."
            Height          =   195
            Left            =   90
            TabIndex        =   37
            Top             =   2048
            Width           =   300
         End
         Begin VB.Label Lbl_Telefono 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Telefono"
            Height          =   195
            Left            =   3330
            TabIndex        =   36
            Top             =   1328
            Width           =   630
         End
         Begin VB.Label Lbl_Limite_Credito 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Credito"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3330
            TabIndex        =   35
            Top             =   2048
            Width           =   495
         End
         Begin VB.Label Lbl_Descuento_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Desc."
            Height          =   195
            Left            =   5700
            TabIndex        =   34
            Top             =   2640
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Lbl_Porcentaje_Descuento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "%"
            Height          =   195
            Left            =   7050
            TabIndex        =   33
            Top             =   2640
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Lbl_Ciudad_estado 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   3330
            TabIndex        =   32
            Top             =   1688
            Width           =   495
         End
      End
      Begin VB.Frame Fra_Datos_Pedido 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del Pedido"
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
         Height          =   2385
         Left            =   7140
         TabIndex        =   24
         Top             =   15
         Width           =   3540
         Begin VB.TextBox Txt_Vendedor 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1110
            TabIndex        =   12
            Top             =   1635
            Width           =   2355
         End
         Begin VB.TextBox Txt_No_Pedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   615
            Width           =   2325
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Pedido 
            Height          =   315
            Left            =   1110
            TabIndex        =   10
            Top             =   960
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   59113475
            CurrentDate     =   39451
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Entrega 
            Height          =   315
            Left            =   1110
            TabIndex        =   11
            Top             =   1305
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   59113475
            CurrentDate     =   39451
         End
         Begin VB.ComboBox Cmb_Vendedor 
            Height          =   315
            Left            =   1110
            TabIndex        =   46
            Top             =   1253
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.Label Lbl_Pedidos 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Pedidos"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   45
            TabIndex        =   48
            Top             =   1995
            Width           =   3435
         End
         Begin VB.Label Lbl_Vendedor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   90
            TabIndex        =   47
            Top             =   1695
            Width           =   690
         End
         Begin VB.Label Lbl_No_Pedido 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Pedido"
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   675
            Width           =   795
         End
         Begin VB.Label Lbl_Fecha 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Lbl_Fecha_Entrega 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Entrega"
            Height          =   195
            Left            =   75
            TabIndex        =   26
            Top             =   1365
            Width           =   555
         End
         Begin VB.Label Lbl_Ventas 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2745
            TabIndex        =   25
            Top             =   2070
            Visible         =   0   'False
            Width           =   45
         End
      End
   End
End
Attribute VB_Name = "Frm_Ope_Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Btn_Agregar_Click
'DESCRIPCIÓN          : Agrega los productos al grid
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Btn_Agregar_Click()
Dim Cont_Detalles As Integer        'Usada para contar detalle del grid
Dim Suma As Double                  'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double              'Suma I.V.A.
Dim Mi_SQL As String
Dim Rs_Consultar_Producto As rdoResultset
Dim Cantidad_Cajas As Double
Dim Aplica_Caja As String
    
    'Valida que los campos tengan valores
    If Val(Txt_Cantidad.Text) > 0 And (Cmb_Descripcion.Text <> "") Then
        If Grid_Detalle_Pedido.Rows = 0 Then
            'Coloca el número de columnas
            Grid_Detalle_Pedido.Cols = 4
            'Pone el encabezado en las columnas
            Grid_Detalle_Pedido.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID" & Chr(9) & "Cantidad_Envases"
        End If
        Mi_SQL = "SELECT * FROM Cat_Productos WHERE Producto_ID='" & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & "'"
        Set Rs_Consultar_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consultar_Producto.EOF Then
            If Trim(Rs_Consultar_Producto!Aplica_Caja) = "SI" Then
                Cantidad_Cajas = Val(Rs_Consultar_Producto!Cantidad_Cajas)
                Aplica_Caja = "SI"
            Else
                Cantidad_Cajas = 0
                Aplica_Caja = "NO"
            End If
        End If
        Rs_Consultar_Producto.Close
        'Agrega el dato en el grid
        If Cmb_Descripcion.ListIndex > -1 Then
            If Aplica_Caja = "SI" Then
                Grid_Detalle_Pedido.AddItem Txt_Cantidad.Text & Chr(9) & UCase(Trim(Cmb_Descripcion.Text)) & Chr(9) & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000") & Chr(9) & Cantidad_Cajas * Val(Txt_Cantidad.Text)
            Else
                Grid_Detalle_Pedido.AddItem Txt_Cantidad.Text & Chr(9) & UCase(Trim(Cmb_Descripcion.Text)) & Chr(9) & Format(Cmb_Descripcion.ItemData(Cmb_Descripcion.ListIndex), "00000")
            End If
        End If
        Grid_Detalle_Pedido.FixedRows = 1
        Grid_Detalle_Pedido.ColWidth(0) = 800              'Ancho de las columnas
        Grid_Detalle_Pedido.ColWidth(1) = 9150
        Grid_Detalle_Pedido.ColAlignment(1) = 1            'Alínea la columna a la derecha
        Grid_Detalle_Pedido.ColWidth(2) = 0
        Grid_Detalle_Pedido.ColWidth(3) = 0
        Btn_Agregar.Default = False
        'Llamada para limpiar los campos de productos
        Cmb_Descripcion.Text = ""
        Txt_Cantidad.Text = ""
        Txt_Cantidad.SetFocus
        Btn_Buscar.Enabled = False
    Else
        MsgBox "Faltan datos para agregar", vbExclamation
    End If
End Sub

Private Sub Btn_Buscar_Click()
    Call Busca_Pedido
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Btn_Eliminar_Click
'DESCRIPCIÓN          : Elimina las partidas del pedido
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Btn_Eliminar_Click()
Dim Cont_Detalles As Integer            'Usada para contar detalles del grid
Dim Suma As Double                      'Usada para sumar el importe y manejo del I.V.A.
Dim Suma_IVA As Double                  'Suma I.V.A.
Dim Resp As Integer

    If Grid_Detalle_Pedido.Rows > 1 Then
  
        Resp = MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbExclamation)
        If Resp = 6 Then
            'Si la respuesta es afirmativa elimina el registro seleccionado
            If Grid_Detalle_Pedido.Rows = 2 Then
                Grid_Detalle_Pedido.FixedRows = 0
                'Quita el item del grid
                Grid_Detalle_Pedido.RemoveItem (Grid_Detalle_Pedido.RowSel + 1)
                Btn_Buscar.Enabled = True
            Else
                If Grid_Detalle_Pedido.Rows > 2 Then
                    Grid_Detalle_Pedido.RemoveItem (Grid_Detalle_Pedido.RowSel)
                End If
            End If
            Suma = 0
            Productos = 0
            'Llamada para limpiar los campos de productos
            Cmb_Descripcion.Text = ""
            Txt_Cantidad.Text = ""
            Txt_Cantidad.SetFocus
            Btn_Buscar.Enabled = False
        End If
    End If
End Sub

Private Sub Btn_Imprimir_Click()
    If Grid_Detalle_Pedido.Rows > 0 Then
        Call Imprime_Pedido
    Else
        MsgBox "Faltan datos para hacer la impresión", vbInformation
    End If
End Sub

Private Sub Btn_Modificar_Click()
    If Grid_Detalle_Pedido.Rows > 0 Then
        If Btn_Modificar.Caption = "Modificar" Then
            Btn_Salir.Caption = "Cancelar"
            Btn_Modificar.Caption = "Actualizar"
            Btn_Imprimir.Enabled = False
            Btn_Nuevo.Enabled = False
            Fra_Datos_Cliente.Enabled = True
            Fra_Datos_Pedido.Enabled = True
            Fra_Detalle_Pedido.Enabled = True
            Fra_Comentarios.Enabled = True
            Btn_Buscar.Enabled = False
            Btn_Nuevo.Visible = True
            Btn_Buscar.Enabled = False
        Else
            Call Actualizar_Pedido
        End If
    Else
        MsgBox "Faltan datos para poder hacer la modificación", vbInformation
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Btn_Nuevo_Click
'DESCRIPCIÓN          : Prepara los controles de pantalla para un nuevo PEDIDO
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Btn_Nuevo_Click()
Set Conectar_Ayudante = New Ayudante  'Manejador del ayudante
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        'Habilita los controles para poder capturar la información y deshabilita otros
        Fra_Datos_Cliente.Enabled = True
        Fra_Datos_Pedido.Enabled = True
        Fra_Detalle_Pedido.Enabled = True
        Fra_Comentarios.Enabled = True
        Cmb_Cliente.Text = ""
        Call Cmb_Cliente_KeyPress(13)
        Cmb_Cliente.SetFocus
        Call Cmb_Descripcion_KeyPress(13)
        Cmb_Descripcion.Text = ""
        Btn_Imprimir.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Nuevo.Visible = True
        Btn_Buscar.Enabled = False
        Btn_Modificar.Enabled = False
        Btn_Salir.Caption = "Cancelar"
        Btn_Nuevo.Caption = "Dar de Alta"
        Dtp_Fecha_Pedido.Value = Now              'Asigna la fecha al día actual
        Dtp_Fecha_Entrega.Value = Now
        Grid_Detalle_Pedido.Rows = 0
        Txt_Cantidad = "1"
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Txt_No_Pedido.Text = Conectar_Ayudante.Maximo_Catalogo("Ope_Pedidos", "Pedido_ID")
    Else
        'Validacion de que esten todos los datos requeridos para dar de alta el Pedido
        If Cmb_Cliente.ListIndex > -1 And Grid_Detalle_Pedido.Rows > 1 Then
            Alta_Pedido 'Da de alta el pedido
            Grid_Detalle_Pedido.Rows = 0
            Cmb_Cliente.Text = ""
            Call Conectar_Ayudante.Limpiar_Textos(Me)
        Else
            MsgBox "Faltan datos para dar de alta el Pedido", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Btn_Salir_Click
'DESCRIPCIÓN          : Calcela para no dar de alta el pedido
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then 'Cierra la pantalla
        Unload Me
    Else 'Cancela el alta
        Fra_Datos_Cliente.Enabled = False
        Fra_Datos_Pedido.Enabled = False
        Fra_Detalle_Pedido.Enabled = False
        Fra_Comentarios.Enabled = False
        Cmb_Cliente.Text = ""
        Cmb_Descripcion.Text = ""
        Btn_Imprimir.Enabled = False
        Btn_Buscar.Enabled = True
        Btn_Nuevo.Visible = True
        Btn_Modificar.Enabled = False
        Btn_Salir.Caption = "Salir"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Nuevo.Enabled = True
        Grid_Detalle_Pedido.Rows = 0
        Call Conectar_Ayudante.Limpiar_Textos(Me)
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Cliente_Click
'DESCRIPCIÓN          : Asigna los datos del cleinte a las cajas de texto
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Cliente_Click()
Dim Mi_SQL As String
Dim Rs_Consulta As rdoResultset
    
On Error GoTo Handler
        Mi_SQL = "SELECT * FROM Cat_Clientes WHERE Cliente_ID='" & Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000") & "' "
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta.EOF Then
            If Not IsNull(Rs_Consulta!Direccion) Then Txt_Direccion.Text = Rs_Consulta!Direccion
            If Not IsNull(Rs_Consulta!RFC) Then Txt_RFC.Text = Rs_Consulta!RFC
            If Not IsNull(Rs_Consulta!Telefono) Then Txt_Telefono.Text = Rs_Consulta!Telefono
            If Not IsNull(Rs_Consulta!Colonia) Then Txt_Colonia.Text = Rs_Consulta!Colonia
            If Not IsNull(Rs_Consulta!Ciudad) Then Txt_Ciudad.Text = Rs_Consulta!Ciudad & ", " & Rs_Consulta!Estado
            If Not IsNull(Rs_Consulta!CP) Then Txt_CP.Text = Rs_Consulta!CP
            If Not IsNull(Rs_Consulta!Dias_Credito) Then Txt_Limite_Credito.Text = Rs_Consulta!Dias_Credito
            If Not IsNull(Rs_Consulta!Cliente_ID) Then Txt_Cliente_ID.Text = Rs_Consulta!Cliente_ID
            Txt_Vendedor = Nombre_Usuario
        End If
        Rs_Consulta.Close
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Cmb_Cliente_KeyPress
'DESCRIPCIÓN          : Llena el combo con los registros del cliente
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Private Sub Cmb_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente, 1, "Nombre")
    Else
        'SE DEPLEGA LA LISTA DEL COMBO
        Despliega_Lista = SendMessageLong(Cmb_Cliente.hwnd, &H14F, True, 0)
    End If
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
Private Sub Form_Load()
    Me.Height = 7605
    Me.Width = 10935
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) \ 2
    Me.Top = (MDIFrm_Apl_Principal.Height - Me.Height) \ 4
    Call Cmb_Cliente_KeyPress(13)
    Call Cmb_Descripcion_KeyPress(13)
    Me.Btn_Imprimir.Enabled = False
    Me.Btn_Modificar.Enabled = False
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN : Alta_Pedido
'DESCRIPCIÓN          : Hace el alta de la factura o una remisión en la base de datos
'PARÁMETROS           :
'CREO                 : Julio Cruz
'FECHA_CREO           : 24-Dic-2010
'MODIFICO             :
'FECHA_MODIFICO       :
'CAUSA_MODIFICACIÓN   :
'*******************************************************************************
Public Sub Alta_Pedido()
Dim Rs_Alta_Pedido As rdoResultset
Dim Rs_Alta_Pedidos_Detalles As rdoResultset
Dim Cont_Detalles_Pedidos As Integer

On Error GoTo Handler
    Conexion_Base.BeginTrans
    
    'Alta del pedido
    Set Rs_Alta_Pedido = Conectar_Ayudante.Recordset_Agregar("Ope_Pedidos")
    'Llena la tabla de Pedidos con los datos contenidos en las cajas de textos
    With Rs_Alta_Pedido
        .AddNew
            .rdoColumns("Pedido_ID") = Format(Txt_No_Pedido.Text, "0000000000")
            .rdoColumns("Cliente_ID") = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000")
            .rdoColumns("Fecha_Pedido") = Format(Dtp_Fecha_Pedido.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Entrega") = Format(Dtp_Fecha_Entrega.Value, "MM/dd/yyyy")
            .rdoColumns("Comentarios") = UCase(Txt_Comentarios.Text)
            .rdoColumns("Vendedor") = UCase(Txt_Vendedor.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now()
            .rdoColumns("ESTATUS") = "PENDIENTE"
        .Update
    End With
    Rs_Alta_Pedido.Close
    
    Set Rs_Alta_Pedidos_Detalles = Conectar_Ayudante.Recordset_Agregar("Ope_Pedidos_Detalles")
    For Cont_Detalles_Factura = 1 To Grid_Detalle_Pedido.Rows - 1
        'Llena la tabla de Ope_Pedidos_Detalles con los datos contenidos en el grid
        With Rs_Alta_Pedidos_Detalles
            .AddNew
              .rdoColumns("Pedido_ID") = Format(Txt_No_Pedido.Text, "0000000000")
              .rdoColumns("Producto_ID") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 2)
              .rdoColumns("Cantidad") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 0)
              .rdoColumns("Descripcion") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 1)
              .rdoColumns("ESTATUS") = "PENDIENTE"
              .rdoColumns("Cantidad_Envases") = Val(Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 3))
            .Update
        End With
    Next Cont_Detalles_Factura
    Rs_Alta_Pedidos_Detalles.Close
  
    Conexion_Base.CommitTrans
    If MsgBox("El pedido ha sido dado de alta" & Chr(13) & " ¿Desea enviarlo a Imprimir?", vbYesNo + vbInformation) = vbYes Then
        Imprime_Pedido 'Se imprime el pedido
    End If
    
    'Deshabilita controles y habilita los necesarios
    Call Btn_Salir_Click
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Busca_Pedido
'DESCRIPCIÓN            : Busca el pedido
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 26-Dic-2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Public Sub Busca_Pedido()
Dim Mi_SQL As String
Dim Rs_Consulta_Pedido As rdoResultset
Dim Rs_Consulta_Detalles_Pedido As rdoResultset
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim No_Pedido As String
Dim Consulta_Estatus_Detalles As rdoResultset
    
    No_Pedido = InputBox("Teclee el número de pedido a consultar", "Consulta de Pedidos")
    Mi_SQL = "SELECT * FROM Ope_Pedidos WHERE Pedido_ID ='" & Format(No_Pedido, "0000000000") & "'"
    Set Rs_Consulta_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena los controles con los datos de la consulta
    If Not Rs_Consulta_Pedido.EOF Then
        Txt_No_Pedido.Text = Rs_Consulta_Pedido!Pedido_ID
        Dtp_Fecha_Pedido.Value = Rs_Consulta_Pedido!Fecha_Pedido
        Dtp_Fecha_Entrega.Value = Rs_Consulta_Pedido!Fecha_Entrega
        Txt_Vendedor = Rs_Consulta_Pedido!Vendedor
        'Consulta del cliente del pedido
        Mi_SQL = "SELECT Cliente_ID, Cat_Clientes.Nombre as Nombre_C, Dias_Credito, RFC, Direccion, Colonia, Ciudad, Telefono, Estado, CP"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID ='" & Format(Rs_Consulta_Pedido!Cliente_ID, "00000") & "'"
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena los controles con los datos de la búsqueda
        If Not Rs_Consulta_Cat_Clientes.EOF Then
            Cmb_Cliente.Clear
            Call Cmb_Cliente_KeyPress(13)
            Call Conectar_Ayudante.Asigna_Item_Combo(Rs_Consulta_Cat_Clientes.rdoColumns(0), Cmb_Cliente)
            If Not IsNull(Rs_Consulta_Cat_Clientes!RFC) Then Txt_RFC.Text = Rs_Consulta_Cat_Clientes!RFC
            If Not IsNull(Rs_Consulta_Cat_Clientes!Direccion) Then Txt_Direccion.Text = Rs_Consulta_Cat_Clientes!Direccion
            If Not IsNull(Rs_Consulta_Cat_Clientes!Colonia) Then Txt_Colonia.Text = Rs_Consulta_Cat_Clientes!Colonia
            If Not IsNull(Rs_Consulta_Cat_Clientes!Ciudad) Then Txt_Ciudad.Text = Rs_Consulta_Cat_Clientes!Ciudad & ", " & Rs_Consulta_Cat_Clientes!Estado
            If Not IsNull(Rs_Consulta_Cat_Clientes!Telefono) Then Txt_Telefono.Text = Rs_Consulta_Cat_Clientes!Telefono
            If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("CP")) Then Txt_CP.Text = Rs_Consulta_Cat_Clientes.rdoColumns("CP")
        End If
        If Not IsNull(Rs_Consulta_Pedido!Comentarios) Then Txt_Comentarios.Text = Rs_Consulta_Pedido!Comentarios
        Rs_Consulta_Cat_Clientes.Close
        Btn_Imprimir.Enabled = True
        'Prepara el recordset para consultar los detalles del pedido
        Mi_SQL = "SELECT * FROM Ope_Pedidos_Detalles WHERE  Pedido_ID='" & Format(No_Pedido, "0000000000") & "'"
        Set Rs_Consulta_Detalles_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Grid_Detalle_Pedido.Rows = 0
        Grid_Detalle_Pedido.Cols = 4
        'Pone el encabezado en las columnas
        Grid_Detalle_Pedido.AddItem "Cantidad" & Chr(9) & "Descripcion" & Chr(9) & "Producto ID" & Chr(9) & "Cantidad_Envases"
        'Llenado del grid
        While Not Rs_Consulta_Detalles_Pedido.EOF
            Grid_Detalle_Pedido.AddItem Rs_Consulta_Detalles_Pedido!Cantidad & Chr(9) & Rs_Consulta_Detalles_Pedido!Descripcion & Chr(9) & Rs_Consulta_Detalles_Pedido!Producto_ID & Chr(9) & Rs_Consulta_Detalles_Pedido!Cantidad_Envases
            Grid_Detalle_Pedido.FixedRows = 1
            Rs_Consulta_Detalles_Pedido.MoveNext
        Wend
        Rs_Consulta_Pedido.Close
        Rs_Consulta_Detalles_Pedido.Close
        'Configura el grid
        Grid_Detalle_Pedido.ColWidth(0) = 800              'Ancho de las columnas
        Grid_Detalle_Pedido.ColWidth(1) = 9200
        Grid_Detalle_Pedido.ColAlignment(1) = 1            'Alínea la columna a la derecha
        Grid_Detalle_Pedido.ColWidth(2) = 0
        Grid_Detalle_Pedido.ColWidth(3) = 0
        Fra_Datos_Cliente.Enabled = False
        Fra_Datos_Pedido.Enabled = False
        Fra_Detalle_Pedido.Enabled = False
        Mi_SQL = " SELECT * FROM Ope_Pedidos_Detalles "
        Mi_SQL = Mi_SQL & " WHERE  Pedido_ID='" & Format(No_Pedido, "0000000000") & "' "
        Mi_SQL = Mi_SQL & " AND Estatus='SURTIDO' "
        Set Consulta_Estatus_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Consulta_Estatus_Detalles.EOF Then
           Btn_Modificar.Enabled = False
        Else
            Btn_Modificar.Enabled = True
        End If
        Consulta_Estatus_Detalles.Close
    Else
        MsgBox "Pedido inexistente", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Actualizar_Pedido
'DESCRIPCIÓN            : Actualiza los datos del Pedido
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 26- Diciembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************'
Public Sub Actualizar_Pedido()
Dim Mi_SQL As String
Dim Rs_Edita_Pedido As rdoResultset
Dim Rs_Elimina_Detalles As rdoResultset
Dim Rs_Alta_Pedidos_Detalles As rdoResultset

On Error GoTo Handler

    Conexion_Base.BeginTrans
    
    'Alta del pedido
    Set Rs_Edita_Pedido = Conectar_Ayudante.Recordset_Editar("SELECT * FROM Ope_Pedidos WHERE Pedido_ID='" & Format(Txt_No_Pedido.Text, "0000000000") & "'")
    'Edita la tabla de Pedidos con los datos contenidos en las cajas de textos
    With Rs_Edita_Pedido
        .Edit
            .rdoColumns("Cliente_ID") = Format(Cmb_Cliente.ItemData(Cmb_Cliente.ListIndex), "00000")
            .rdoColumns("Fecha_Pedido") = Format(Dtp_Fecha_Pedido.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Entrega") = Format(Dtp_Fecha_Entrega.Value, "MM/dd/yyyy")
            .rdoColumns("Comentarios") = Txt_Comentarios.Text
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now()
        .Update
    End With
    Rs_Edita_Pedido.Close
    
    'SE ELIMINAN LOS DETALLES PARA PODER ACTUALIZARLAS
    Mi_SQL = " SELECT * FROM Ope_Pedidos_Detalles WHERE Pedido_ID='" & Format(Txt_No_Pedido.Text, "0000000000") & "'"
    Set Rs_Elimina_Detalles = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
    While Not Rs_Elimina_Detalles.EOF
        With Rs_Elimina_Detalles
            .Delete
        End With
        Rs_Elimina_Detalles.MoveNext
    Wend
    Rs_Elimina_Detalles.Close
    
    Set Rs_Alta_Pedidos_Detalles = Conectar_Ayudante.Recordset_Agregar("Ope_Pedidos_Detalles")
    For Cont_Detalles_Factura = 1 To Grid_Detalle_Pedido.Rows - 1
        'Llena la tabla de Ope_Pedidos_Detalles con los datos contenidos en el grid
        With Rs_Alta_Pedidos_Detalles
            .AddNew
              .rdoColumns("Pedido_ID") = Format(Txt_No_Pedido.Text, "0000000000")
              .rdoColumns("Producto_ID") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 2)
              .rdoColumns("Cantidad") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 0)
              .rdoColumns("Descripcion") = Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 1)
              .rdoColumns("ESTATUS") = "PENDIENTE"
              .rdoColumns("Cantidad_Envases") = Val(Grid_Detalle_Pedido.TextMatrix(Cont_Detalles_Factura, 3))
            .Update
        End With
    Next Cont_Detalles_Factura
    Rs_Alta_Pedidos_Detalles.Close
  
    Conexion_Base.CommitTrans
    Btn_Modificar.Caption = "Modificar"
    If MsgBox("El pedido ha sido modificado" & Chr(13) & " ¿Desea enviarlo a Imprimir?", vbYesNo + vbInformation) = vbYes Then
       Imprime_Pedido 'Se imprime el pedido
    End If
    'Deshabilita controles y habilita los necesarios
    Call Btn_Salir_Click
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Imprime_Pedido
'DESCRIPCIÓN            : Imprime el pedido
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 26- Diciembre - 2010
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************'
Public Sub Imprime_Pedido()
Dim Fila As Integer
Dim Cantidad_Articulos As Double
Dim Cont_Renglon As Integer
    
On Error GoTo Handler
    Printer.FontSize = "14"
    Printer.Font = "COURIER NEW"
    Printer.FontBold = True
    Printer.Print
    Printer.Print
    Printer.FontSize = "10"
    Printer.Print Spc(80); "  PEDIDO"
    Printer.Print Spc(80); Txt_No_Pedido.Text
    Printer.FontBold = False
    Printer.Print
    Printer.Print "  Vendido a:"
    Printer.Print "  NOMBRE O RAZON: "; Mid(Cmb_Cliente.Text, 1, 50); Spc(70 - Len(Mid("NOMBRE O RAZON: " & Cmb_Cliente.Text, 1, 61))); "HORA      : "; Format(Dtp_Fecha_Pedido.Value, "HH:mm:ss")
    Printer.Print "  CALLE         : "; Mid(Txt_Direccion.Text, 1, 50); Spc(70 - Len(Mid("CALLE         : " & Txt_Direccion.Text, 1, 50))); "FECHA     : "; Format(Dtp_Fecha_Pedido.Value, "dd/MMM/yyyy")
    Printer.Print "  COLONIA       : "; Mid(Txt_Colonia.Text, 1, 50); Spc(70 - Len(Mid("COLONIA       : " & Txt_Colonia.Text, 1, 50)))
    Printer.Print "  C.P.          : "; Txt_CP.Text; "          R.F.C.: "; Txt_RFC.Text; Spc(70 - Len("C.P.          : " & Mid(Txt_CP.Text & "          R.F.C.: " & Txt_RFC.Text, 1, 50)))
    Printer.Print "  TELEFONO      : "; Txt_Telefono.Text; Spc(70 - Len(Mid("TELEFONO      : " & Txt_Telefono.Text, 1, 50))); Txt_Vendedor.Text
    Printer.Print "  CIUDAD        : "; Txt_Ciudad.Text
    Printer.Print
    Printer.Print "  CANT     DESCRIPCION                                                                            "
    Printer.Print "________________________________________________________________________________________________  "
    Printer.Print
    Printer.FontSize = "9"
    For Fila = 1 To Grid_Detalle_Pedido.Rows - 1
        If Val(Grid_Detalle_Pedido.TextMatrix(Fila, 3)) > 0 Then
            Cantidad_Articulos = Val(Cantidad_Articulos) + Val(Grid_Detalle_Pedido.TextMatrix(Fila, 1))
            If Trim(Grid_Detalle_Pedido.TextMatrix(Fila, 0)) <> "" Then
                Printer.Print Conectar_Ayudante.Alinea_Derecha(Grid_Detalle_Pedido.TextMatrix(Fila, 0), 6); "  "; _
                Spc(5); _
                Trim(Mid(Grid_Detalle_Pedido.TextMatrix(Fila, 1), 1, 80)); Spc(82 - Len(Mid(Grid_Detalle_Pedido.TextMatrix(Fila, 1), 1, 80))); Grid_Detalle_Pedido.TextMatrix(Fila, 3) & " " & "Envases"
            End If
        Else
            Cantidad_Articulos = Val(Cantidad_Articulos) + Val(Grid_Detalle_Pedido.TextMatrix(Fila, 1))
            If Trim(Grid_Detalle_Pedido.TextMatrix(Fila, 0)) <> "" Then
                Printer.Print Conectar_Ayudante.Alinea_Derecha(Grid_Detalle_Pedido.TextMatrix(Fila, 0), 6); "  "; _
                Spc(5); _
                Trim(Mid(Grid_Detalle_Pedido.TextMatrix(Fila, 1), 1, 100))
            End If
        End If
    Next Fila
    Printer.FontSize = "10"
    Printer.Print
    Printer.Print "________________________________________________________________________________________________  "
    Printer.Print
    Printer.Print
    Printer.Print "   COMENTARIOS: "
    Printer.Print
    Cont_Renglon = Imprime_Varias_Lineas(Conectar_Ayudante.Quitar_Caracter(UCase(Txt_Comentarios.Text), Chr(13)), 90, 10, 0.3)
    Printer.Print
    Printer.EndDoc
    MsgBox "Pedido enviado a impresion", vbInformation
    Exit Sub
Handler:
    For Each Er In rdoErrors
       'MsgBox Er.Description, vbInformation
    Next Er
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
