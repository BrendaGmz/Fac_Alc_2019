VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cfg_Impresion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Configuracion de Formatos"
   ClientHeight    =   7080
   ClientLeft      =   5355
   ClientTop       =   450
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8055
   Begin VB.CommandButton Btn_Modificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modificar"
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
      Left            =   2280
      Picture         =   "Frm_Cfg_Impresion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "M"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar_Formato 
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
      Height          =   660
      Left            =   4440
      Picture         =   "Frm_Cfg_Impresion.frx":3731
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "B"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
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
      Left            =   6600
      Picture         =   "Frm_Cfg_Impresion.frx":6CEB
      Style           =   1  'Graphical
      TabIndex        =   38
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
      Picture         =   "Frm_Cfg_Impresion.frx":A3EA
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "A"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Campos 
      Height          =   1815
      Left            =   240
      TabIndex        =   36
      Top             =   4320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   6
      BackColorBkg    =   16777215
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin VB.Frame Fra_Campos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Campos a Imprimir"
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
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   7815
      Begin VB.CommandButton Btn_Modificar_Detalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modificar"
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
         Left            =   3300
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frm_Cfg_Impresion.frx":D921
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1260
         Width           =   1230
      End
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
         Left            =   6480
         Picture         =   "Frm_Cfg_Impresion.frx":10BD7
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1260
         Width           =   1230
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
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frm_Cfg_Impresion.frx":13E89
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1260
         Width           =   1230
      End
      Begin VB.ComboBox Cmb_Formato_Campo 
         Height          =   315
         ItemData        =   "Frm_Cfg_Impresion.frx":1713F
         Left            =   960
         List            =   "Frm_Cfg_Impresion.frx":1714C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   900
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_Tipo_Campo 
         Height          =   315
         ItemData        =   "Frm_Cfg_Impresion.frx":17169
         Left            =   960
         List            =   "Frm_Cfg_Impresion.frx":17173
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   540
         Width           =   2895
      End
      Begin VB.TextBox Txt_Longitud_Campo 
         Height          =   285
         Left            =   5400
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Txt_Coordenada_Y 
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Txt_Coordenada_X 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_Nombre_Campo 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label Lbl_Formato 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formato"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Lbl_Tipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Lbl_Longitud 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitud"
         Height          =   195
         Index           =   4
         Left            =   4080
         TabIndex        =   9
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Lbl_Coord_Y 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenada Y"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Lbl_Coord_X 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenada X"
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Lbl_Nombre_Campo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Fra_Formato 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Frame Fra_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
         Enabled         =   0   'False
         Height          =   1035
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   7575
         Begin VB.ComboBox Cmb_Letra_Detalles 
            Height          =   315
            ItemData        =   "Frm_Cfg_Impresion.frx":17189
            Left            =   960
            List            =   "Frm_Cfg_Impresion.frx":1719C
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txt_Tamaño_Detalles 
            Height          =   315
            Left            =   6600
            TabIndex        =   28
            Top             =   180
            Width           =   855
         End
         Begin VB.ComboBox Cmb_Estilo_Detalles 
            Height          =   315
            ItemData        =   "Frm_Cfg_Impresion.frx":171DD
            Left            =   3720
            List            =   "Frm_Cfg_Impresion.frx":171E7
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txt_Separacion_Detalles 
            Height          =   315
            Left            =   960
            TabIndex        =   26
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Txt_No_Detalles 
            Height          =   315
            Left            =   3720
            TabIndex        =   25
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Txt_No_Columnas 
            Height          =   315
            Left            =   6600
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Lbl_Tamaño_Detalle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño"
            Height          =   195
            Index           =   35
            Left            =   5520
            TabIndex        =   35
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Lbl_Estilo_Detalle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estilo"
            Height          =   195
            Index           =   36
            Left            =   2760
            TabIndex        =   34
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Lbl_Fuente_Detalle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente"
            Height          =   195
            Index           =   38
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Lbl_Separacion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separacion"
            Height          =   195
            Index           =   39
            Left            =   120
            TabIndex        =   32
            Top             =   660
            Width           =   810
         End
         Begin VB.Label Lbl_No_Partidas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Partidas"
            Height          =   195
            Index           =   40
            Left            =   2760
            TabIndex        =   31
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Lbl_No_Columnas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Columnas"
            Height          =   195
            Index           =   41
            Left            =   5520
            TabIndex        =   30
            Top             =   660
            Width           =   990
         End
      End
      Begin VB.Frame Fra_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   7575
         Begin VB.ComboBox Cmb_Estilo_Generales 
            Height          =   315
            ItemData        =   "Frm_Cfg_Impresion.frx":171FC
            Left            =   3720
            List            =   "Frm_Cfg_Impresion.frx":17206
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txt_Tamaño_Generales 
            Height          =   315
            Left            =   6600
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox Cmb_Letra_Generales 
            Height          =   315
            ItemData        =   "Frm_Cfg_Impresion.frx":1721B
            Left            =   960
            List            =   "Frm_Cfg_Impresion.frx":1722E
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Lbl_Fuente_General 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Lbl_Estilo_General 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estilo"
            Height          =   195
            Index           =   34
            Left            =   2760
            TabIndex        =   21
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Lbl_Tamaño_General 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño"
            Height          =   195
            Index           =   0
            Left            =   5520
            TabIndex        =   20
            Top             =   300
            Width           =   585
         End
      End
      Begin VB.ComboBox Cmb_Nombre_Formato 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   300
         Width           =   6495
      End
      Begin VB.Label Lbl_Nombre_Formato 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "Frm_Cfg_Impresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Agregar_Click()
    
    If Txt_Nombre_Campo.Text <> "" And Val(Txt_Coordenada_X.Text) > 0 And Val(Txt_Coordenada_Y.Text) > 0 And Val(Txt_Longitud_Campo.Text) > 0 Then
        
        If Grid_Campos.Rows = 0 Then
            Grid_Campos.AddItem "Campo" & Chr(9) & "X" & Chr(9) & "Y" & Chr(9) & "Longitud" & Chr(9) & "Tipo" & Chr(9) & "Formato"
            Grid_Campos.ColWidth(0) = 2000
            Grid_Campos.ColWidth(1) = 1000
            Grid_Campos.ColWidth(2) = 1000
            Grid_Campos.ColWidth(3) = 1000
            Grid_Campos.ColWidth(4) = 1000
            Grid_Campos.ColWidth(5) = 1000
        End If
        
        Grid_Campos.AddItem Txt_Nombre_Campo.Text & Chr(9) & Txt_Coordenada_X.Text & Chr(9) & Txt_Coordenada_Y.Text & Chr(9) & Txt_Longitud_Campo.Text & Chr(9) & Cmb_Tipo_Campo.Text & Chr(9) & Cmb_Formato_Campo.Text
        Grid_Campos.FixedRows = 1
        Txt_Coordenada_X.Text = ""
        Txt_Coordenada_Y.Text = ""
        Txt_Nombre_Campo.Text = ""
        Txt_Longitud_Campo.Text = ""
    Else
        MsgBox "Faltan Datos para agregar el campo", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    
End Sub

Private Sub Btn_Eliminar_Click()
    If Grid_Campos.RowSel > 0 Then
        'Remueve del grid_Campos el campos seleccionado por el usuario
        If Grid_Campos.Rows = 2 Then
            Grid_Campos.Rows = 0
        Else
            Grid_Campos.RemoveItem Grid_Campos.RowSel
        End If
        Txt_Coordenada_X.Text = ""
        Txt_Coordenada_Y.Text = ""
        Txt_Nombre_Campo.Text = ""
        Txt_Longitud_Campo.Text = ""
    End If
End Sub

Private Sub Btn_Eliminar_Formato_Click()
Dim Valor As Integer
Dim Mi_SQL As String
Set Conectar_Ayudante = New Ayudante
    
    If Cmb_Nombre_Formato.Text <> "" And Grid_Campos.Rows > 1 Then
        Valor = MsgBox("¿Seguro de Eliminar el Formato?", vbYesNo + vbQuestion, UCase(MDIFrm_Apl_Principal.Caption))
        
        If Valor = 6 Then
            Conexion_Base.BeginTrans
            Mi_SQL = "DELETE FROM Cfg_Formatos_Detalles WHERE Nombre = '" & Cmb_Nombre_Formato.Text & "'"
            Conexion_Base.Execute Mi_SQL
            Mi_SQL = "DELETE FROM Cfg_Formatos WHERE Nombre = '" & Cmb_Nombre_Formato.Text & "'"
            Conexion_Base.Execute Mi_SQL
            Conexion_Base.CommitTrans
            Limpia_Controles
            Consulta_Formatos
            MsgBox "Formato Eliminado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    Else
        MsgBox "No existe un formato seleccionado para eliminar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
    
End Sub

Private Sub Btn_Modificar_Detalle_Click()
    
    If Grid_Campos.RowSel > 0 Then
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 0) = Txt_Nombre_Campo.Text
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 1) = Txt_Coordenada_X.Text
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 2) = Txt_Coordenada_Y.Text
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 3) = Txt_Longitud_Campo.Text
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 4) = Cmb_Tipo_Campo.Text
        Grid_Campos.TextMatrix(Grid_Campos.RowSel, 5) = Cmb_Formato_Campo.Text
        Txt_Coordenada_X.Text = ""
        Txt_Coordenada_Y.Text = ""
        Txt_Nombre_Campo.Text = ""
        Txt_Longitud_Campo.Text = ""
    End If
    
End Sub

Private Sub Cmb_Nombre_Formato_Click()
Dim Mi_SQL As String                                    'Obtiene los valores de la consulta
Dim Rs_Consulta_Cfg_Formatos As rdoResultset            'Consulta los valores del formato en la tabla Cfg_Formatos
Dim Rs_Consulta_Cfg_Formatos_Detalles As rdoResultset   'Consulta los valores de los detalles del formato en la tabla Cfg_Formatos_Detalles
Set Conectar_Ayudante = New Ayudante
    
    'Consulta los valores del formto
    Mi_SQL = "SELECT * " & " FROM Cfg_Formatos" & " WHERE Nombre = '" & Cmb_Nombre_Formato.Text & "'"
    Set Rs_Consulta_Cfg_Formatos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si encuentra los valores pertenecientes al formto entonces llena las cajas
    'de texto y el grid con los valores correspondientes al formato consultado
    
    If Not Rs_Consulta_Cfg_Formatos.EOF Then
        With Rs_Consulta_Cfg_Formatos
            Txt_Tamaño_Generales.Text = .rdoColumns("Tamaño_Generales")
            Txt_Tamaño_Detalles.Text = .rdoColumns("Tamaño_Detalles")
            Txt_Separacion_Detalles.Text = .rdoColumns("Separacion_Detalles")
            Txt_No_Detalles.Text = .rdoColumns("No_Detalles")
            Txt_No_Columnas.Text = .rdoColumns("No_Columnas")
            
            For I = 0 To Cmb_Letra_Generales.ListCount
                If Cmb_Letra_Generales.List(I) = .rdoColumns("Letra_Generales") Then
                   Cmb_Letra_Generales.ListIndex = I
                   I = Cmb_Letra_Generales.ListCount
                End If
            Next I
            
            For I = 0 To Cmb_Estilo_Generales.ListCount
                If Cmb_Estilo_Generales.List(I) = .rdoColumns("Estilo_Generales") Then
                   Cmb_Estilo_Generales.ListIndex = I
                   I = Cmb_Estilo_Generales.ListCount
                End If
            Next I
            
            For I = 0 To Cmb_Letra_Detalles.ListCount
                If Cmb_Letra_Detalles.List(I) = .rdoColumns("Letra_Detalles") Then
                   Cmb_Letra_Detalles.ListIndex = I
                   I = Cmb_Letra_Detalles.ListCount
                End If
            Next I
            
            For I = 0 To Cmb_Estilo_Detalles.ListCount
                If Cmb_Estilo_Detalles.List(I) = .rdoColumns("Estilo_Detalles") Then
                   Cmb_Estilo_Detalles.ListIndex = I
                   I = Cmb_Estilo_Detalles.ListCount
                End If
            Next I
        End With
        
        'Consulta los Detalles del formato
        Mi_SQL = "SELECT * " & " FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Nombre = '" & Cmb_Nombre_Formato & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Tipo DESC, Nombre"
        Set Rs_Consulta_Cfg_Formatos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Si encuentra los valores de los detalles del formto entonces los detalles
        'los agrega el grid_campos
        
        If Not Rs_Consulta_Cfg_Formatos_Detalles.EOF Then
            With Rs_Consulta_Cfg_Formatos_Detalles
                Grid_Campos.Rows = 0
                'Pone el encabezado del grid_campos y configura el grid_campos
                Grid_Campos.AddItem "Campo" & Chr(9) & "X" & Chr(9) & "Y" & Chr(9) & _
                                    "Longitud" & Chr(9) & "Tipo" & Chr(9) & "Formato"
                Grid_Campos.ColWidth(0) = 2000
                Grid_Campos.ColWidth(1) = 1000
                Grid_Campos.ColWidth(2) = 1000
                Grid_Campos.ColWidth(3) = 1000
                Grid_Campos.ColWidth(4) = 1000
                Grid_Campos.ColWidth(5) = 1000
                'Llena el grid_campos con los resultadosobtenidos de la consulta
                While Not Rs_Consulta_Cfg_Formatos_Detalles.EOF
                    Grid_Campos.AddItem .rdoColumns("Campo") & Chr(9) & .rdoColumns("X") & Chr(9) & .rdoColumns("Y") _
                                            & Chr(9) & .rdoColumns("Longitud") & Chr(9) & .rdoColumns("Tipo") & Chr(9) & .rdoColumns("Formato")
                    Grid_Campos.FixedRows = 1
                    Rs_Consulta_Cfg_Formatos_Detalles.MoveNext
                Wend
            End With
        End If
        Rs_Consulta_Cfg_Formatos.Close
        Rs_Consulta_Cfg_Formatos_Detalles.Close
    End If
    
End Sub

Private Sub Btn_Modificar_Click()
    
    If Cmb_Nombre_Formato.Text <> "" And Grid_Campos.Rows > 1 Then
        
        If Btn_Modificar.Caption = "Modificar" Then
            Fra_generales.Enabled = True
            Fra_detalles.Enabled = True
            Fra_Campos.Enabled = True
            Grid_Campos.Enabled = True
            Btn_Modificar.Caption = "Actualizar"
            Btn_Nuevo.Enabled = False
            Btn_Eliminar_Formato.Enabled = False
            Cmb_Nombre_Formato.Enabled = True
        Else
            
            If Cmb_Nombre_Formato.Text <> "" And Grid_Campos.Rows > 1 And Val(Txt_Tamaño_Generales.Text) > 0 And Val(Txt_Tamaño_Detalles.Text) > 0 _
                    And Val(Txt_Separacion_Detalles.Text) > 0 And Val(Txt_No_Detalles.Text) > 0 And Val(Txt_No_Columnas.Text) > 0 Then
                Modifica_Formato
            Else
                MsgBox "Faltan datos para modificar el formato", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        End If
    Else
        MsgBox "No existe un formato seleccionado para modificar", vbExclamation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
End Sub

Private Sub Btn_Nuevo_Click()
Set Conectar_Ayudante = New Ayudante
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        Limpia_Controles
        Fra_generales.Enabled = True
        Fra_detalles.Enabled = True
        Fra_Campos.Enabled = True
        Grid_Campos.Enabled = True
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Eliminar_Formato.Enabled = False
        Cmb_Nombre_Formato.Clear
        Cmb_Nombre_Formato.SetFocus
    Else
        
        If Cmb_Nombre_Formato.Text <> "" And Grid_Campos.Rows > 1 And Val(Txt_Tamaño_Generales.Text) > 0 And Val(Txt_Tamaño_Detalles.Text) > 0 And Val(Txt_Separacion_Detalles.Text) > 0 And Val(Txt_No_Detalles.Text) > 0 And Val(Txt_No_Columnas.Text) > 0 Then
            Alta_Formato
        Else
            MsgBox "Faltan datos para dar de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    End If
    
End Sub

Public Sub Alta_Formato()
Dim Rs_Alta_Cfg_Formatos As rdoResultset
Dim Rs_Alta_Cfg_Formatos_Detalles As rdoResultset
Dim I As Integer
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de bancos
    Set Rs_Alta_Cfg_Formatos = Conectar_Ayudante.Recordset_Agregar("Cfg_Formatos")
    With Rs_Alta_Cfg_Formatos
        .AddNew
            .rdoColumns("Nombre") = Trim(Cmb_Nombre_Formato.Text)
            .rdoColumns("Letra_Generales") = Trim(Cmb_Letra_Generales.Text)
            .rdoColumns("Estilo_Generales") = Trim(Cmb_Estilo_Generales.Text)
            .rdoColumns("Tamaño_Generales") = Txt_Tamaño_Generales.Text
            .rdoColumns("Letra_Detalles") = Cmb_Letra_Detalles.Text
            .rdoColumns("Estilo_Detalles") = Cmb_Estilo_Detalles.Text
            .rdoColumns("Tamaño_Detalles") = Txt_Tamaño_Detalles.Text
            .rdoColumns("Separacion_Detalles") = Txt_Separacion_Detalles.Text
            .rdoColumns("No_Detalles") = Txt_No_Detalles.Text
            .rdoColumns("No_Columnas") = Txt_No_Columnas.Text
            .rdoColumns("Usuario_Creo") = Usuario_Sistema
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    'Da de alta los campos
    
    Set Rs_Alta_Cfg_Formatos_Detalles = Conectar_Ayudante.Recordset_Agregar("Cfg_Formatos_Detalles")
    For I = 1 To Grid_Campos.Rows - 1
        With Rs_Alta_Cfg_Formatos_Detalles
            .AddNew
                .rdoColumns("Nombre") = Cmb_Nombre_Formato.Text
                .rdoColumns("Campo") = Grid_Campos.TextMatrix(I, 0)
                .rdoColumns("X") = Val(Grid_Campos.TextMatrix(I, 1))
                .rdoColumns("Y") = Val(Grid_Campos.TextMatrix(I, 2))
                .rdoColumns("Longitud") = Val(Grid_Campos.TextMatrix(I, 3))
                .rdoColumns("Tipo") = Grid_Campos.TextMatrix(I, 4)
                .rdoColumns("Formato") = Grid_Campos.TextMatrix(I, 5)
            .Update
        End With
    Next I
    Fra_generales.Enabled = False
    Fra_detalles.Enabled = False
    Fra_Campos.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar_Formato.Enabled = True
    Limpia_Controles
    Consulta_Formatos
    Conexion_Base.CommitTrans
    MsgBox "Formato dado de alta", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
    
HANDLER:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
    
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Formatos
    'DESCRIPCIÓN: Da de alta en la tabla Cfg_Formato el formato y en la tabla
    'Cfg_Gormato_Detalles, los detalles que contendra el formato a dar de alta
    'PARÁMETROS:
    'CREO:
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Sub Modifica_Formato()
Dim Mi_SQL As String                                'Obtiene los valores de la consulta
Dim Rs_Consulta_Cfg_Formatos As rdoResultset        'Consulta los datos que contiene Cfg_Formatos con respecto al formato consultado
Dim Rs_Alta_Cfg_Formatos_Detalles As rdoResultset   'Da de alta los nuevos valores del los detalles que contiene el formato
Dim I As Integer                                    'Contador de las filas del grid_campos
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta los datos del formato para despues actualizar los datos del formato
    Mi_SQL = "SELECT *" & " FROM Cfg_Formatos"
    Mi_SQL = Mi_SQL & " WHERE Nombre = '" & Cmb_Nombre_Formato.Text & "'"
    Set Rs_Consulta_Cfg_Formatos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    
    If Not Rs_Consulta_Cfg_Formatos.EOF Then
        'Modifica los datos del formato
        With Rs_Consulta_Cfg_Formatos
            .Edit
                .rdoColumns("Letra_Generales") = Cmb_Letra_Generales.Text
                .rdoColumns("Estilo_Generales") = Cmb_Estilo_Generales.Text
                .rdoColumns("Tamaño_Generales") = Txt_Tamaño_Generales.Text
                .rdoColumns("Letra_Detalles") = Cmb_Letra_Detalles.Text
                .rdoColumns("Estilo_Detalles") = Cmb_Estilo_Detalles.Text
                .rdoColumns("Tamaño_Detalles") = Txt_Tamaño_Detalles.Text
                .rdoColumns("Separacion_Detalles") = Txt_Separacion_Detalles.Text
                .rdoColumns("No_Detalles") = Txt_No_Detalles.Text
                .rdoColumns("No_Columnas") = Txt_No_Columnas.Text
                .rdoColumns("Usuario_Modifico") = Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    
    Rs_Consulta_Cfg_Formatos.Close
    'Elimina los detalles del formato que se tienen en la tabla Cfg_Formatos_Detalles
    Mi_SQL = "DELETE FROM Cfg_Formatos_Detalles  WHERE Nombre = '" & Cmb_Nombre_Formato.Text & "'"
    Conexion_Base.Execute Mi_SQL
    'Da de alta los campos
    
    For I = 1 To Grid_Campos.Rows - 1
        Set Rs_Alta_Cfg_Formatos_Detalles = Conectar_Ayudante.Recordset_Agregar("Cfg_Formatos_Detalles")
        With Rs_Alta_Cfg_Formatos_Detalles
            .AddNew
                .rdoColumns("Nombre") = Cmb_Nombre_Formato.Text
                .rdoColumns("Campo") = Grid_Campos.TextMatrix(I, 0)
                .rdoColumns("X") = Val(Grid_Campos.TextMatrix(I, 1))
                .rdoColumns("Y") = Val(Grid_Campos.TextMatrix(I, 2))
                .rdoColumns("Longitud") = Val(Grid_Campos.TextMatrix(I, 3))
                .rdoColumns("Tipo") = Grid_Campos.TextMatrix(I, 4)
                .rdoColumns("Formato") = Grid_Campos.TextMatrix(I, 5)
            .Update
        End With
    Next I
    Rs_Alta_Cfg_Formatos_Detalles.Close
    Fra_generales.Enabled = False
    Fra_detalles.Enabled = False
    Fra_Campos.Enabled = False
    Btn_Modificar.Caption = "Modificar"
    Btn_Nuevo.Enabled = True
    Btn_Eliminar_Formato.Enabled = True
    Conexion_Base.CommitTrans
    MsgBox "Formato Modificado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub

HANDLER:
    Correcto = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub

Private Sub Btn_Salir_Click()
    Unload Frm_Cfg_Impresion
End Sub

Private Sub Form_Load()
Set Conectar_Ayudante = New Ayudante
    Me.Top = 0
    Me.Width = 8280
    Me.Height = 7470
    Limpia_Controles
    Consulta_Formatos
End Sub

Public Sub Limpia_Controles()
    Cmb_Nombre_Formato.Text = ""
    Cmb_Letra_Generales.ListIndex = 0
    Cmb_Estilo_Generales.ListIndex = 0
    Txt_Tamaño_Generales.Text = ""
    Cmb_Letra_Detalles.ListIndex = 0
    Cmb_Estilo_Detalles.ListIndex = 0
    Txt_Tamaño_Detalles.Text = ""
    Txt_Separacion_Detalles.Text = ""
    Txt_No_Detalles.Text = ""
    Txt_No_Columnas.Text = ""
    Txt_Nombre_Campo.Text = ""
    Cmb_Formato_Campo.ListIndex = 0
    Cmb_Tipo_Campo.ListIndex = 0
    Txt_Coordenada_X.Text = ""
    Txt_Coordenada_Y.Text = ""
    Txt_Longitud_Campo.Text = ""
    Grid_Campos.Rows = 0
End Sub

Public Sub Consulta_Formatos()
Dim Mi_SQL As String
Dim Rs_Consulta_Cfg_Formatos As rdoResultset
Set Conectar_Ayudante = New Ayudante
    
    'Consulta del catalogo de Clientes
    Mi_SQL = "SELECT  Nombre " & " FROM Cfg_Formatos" & " ORDER BY Nombre "
    Set Rs_Consulta_Cfg_Formatos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Consulta del catalogo de Clientes
    Cmb_Nombre_Formato.Clear
    If Not Rs_Consulta_Cfg_Formatos.EOF Then
        While Not Rs_Consulta_Cfg_Formatos.EOF
            Cmb_Nombre_Formato.AddItem Rs_Consulta_Cfg_Formatos!Nombre
            Rs_Consulta_Cfg_Formatos.MoveNext
        Wend
    End If
    Rs_Consulta_Cfg_Formatos.Close

End Sub


Private Sub Grid_Campos_Click()
    
    If Grid_Campos.RowSel > 0 Then
        Txt_Nombre_Campo.Text = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 0)
        Txt_Coordenada_X.Text = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 1)
        Txt_Coordenada_Y.Text = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 2)
        Txt_Longitud_Campo.Text = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 3)
        For I = 0 To Cmb_Tipo_Campo.ListCount
            If Cmb_Tipo_Campo.List(I) = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 4) Then
               Cmb_Tipo_Campo.ListIndex = I
               I = Cmb_Tipo_Campo.ListCount
            End If
        Next I
        
        For I = 0 To Cmb_Formato_Campo.ListCount
            If Cmb_Formato_Campo.List(I) = Grid_Campos.TextMatrix(Grid_Campos.RowSel, 5) Then
               Cmb_Formato_Campo.ListIndex = I
               I = Cmb_Formato_Campo.ListCount
            End If
        Next I
        Btn_Modificar_Detalle.Enabled = True
    Else
        Btn_Modificar_Detalle.Enabled = False
    End If

End Sub

