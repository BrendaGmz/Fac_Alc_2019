VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Referencia_Iconos 
   Caption         =   "Varios"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   13425
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   7500
      Left            =   -15
      ScaleHeight     =   7440
      ScaleWidth      =   13335
      TabIndex        =   0
      Top             =   -165
      Width           =   13395
      Begin VB.CommandButton Btn_Crystal 
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
         Left            =   11925
         Picture         =   "Frm_Cat_Varios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "A"
         Top             =   6465
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
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
         Left            =   11925
         Picture         =   "Frm_Cat_Varios.frx":3560
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "A"
         Top             =   5670
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Agrega_clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   90
         TabIndex        =   27
         Top             =   2880
         Width           =   13170
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   8850
            TabIndex        =   36
            Top             =   255
            Width           =   795
         End
         Begin VB.TextBox Txt_Importe 
            Height          =   315
            Left            =   10635
            TabIndex        =   33
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox Txt_Carga 
            Height          =   315
            Left            =   6975
            TabIndex        =   32
            Top             =   270
            Width           =   795
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
            Left            =   11775
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Frm_Cat_Varios.frx":6BF7
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   1230
         End
         Begin VB.ComboBox Cmb_Cliente 
            Height          =   315
            Left            =   1215
            TabIndex        =   28
            Top             =   270
            Width           =   4890
         End
         Begin VB.Label Lbl_Cantidad 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cantidad"
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
            Left            =   8055
            TabIndex        =   37
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Lbl_importe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Importe"
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
            Left            =   9930
            TabIndex        =   35
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Lbl_Carga 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Carga"
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
            Left            =   6270
            TabIndex        =   34
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Lbl_Agrega_Clientes 
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
            Left            =   135
            TabIndex        =   29
            Top             =   285
            Width           =   1140
         End
      End
      Begin VB.CommandButton Btn_Regresar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anterior"
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
         Left            =   45
         Picture         =   "Frm_Cat_Varios.frx":9EAD
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "A"
         Top             =   6495
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Siguiente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Siguiente"
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
         Left            =   1530
         Picture         =   "Frm_Cat_Varios.frx":D4A0
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Reporte 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte"
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
         Left            =   3030
         Picture         =   "Frm_Cat_Varios.frx":10A97
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Pdf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Crear PDF"
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
         Left            =   4545
         Picture         =   "Frm_Cat_Varios.frx":14027
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Graficar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Graficar"
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
         Left            =   6030
         Picture         =   "Frm_Cat_Varios.frx":176AE
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Correo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviar Correo"
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
         Left            =   7530
         Picture         =   "Frm_Cat_Varios.frx":1AC58
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Generar_codigo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Codigo de Barras"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   9015
         Picture         =   "Frm_Cat_Varios.frx":1E1F6
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "A"
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Sincronizar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sincronización"
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
         Left            =   10470
         Picture         =   "Frm_Cat_Varios.frx":21737
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "A"
         Top             =   6465
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
         Left            =   10470
         Picture         =   "Frm_Cat_Varios.frx":24006
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "A"
         Top             =   5670
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Guardar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Guardar"
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
         Left            =   9015
         Picture         =   "Frm_Cat_Varios.frx":27705
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "A"
         Top             =   5685
         UseMaskColor    =   -1  'True
         Width           =   1350
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
         Height          =   660
         Left            =   7530
         Picture         =   "Frm_Cat_Varios.frx":2ADAC
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "B"
         Top             =   5685
         UseMaskColor    =   -1  'True
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
         Left            =   6030
         Picture         =   "Frm_Cat_Varios.frx":2E366
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "C"
         Top             =   5685
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Exportar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar"
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
         Left            =   4560
         Picture         =   "Frm_Cat_Varios.frx":318F2
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "A"
         Top             =   5670
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
         Left            =   3030
         Picture         =   "Frm_Cat_Varios.frx":3412C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "A"
         Top             =   5685
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
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
         Left            =   1530
         Picture         =   "Frm_Cat_Varios.frx":375F2
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "M"
         Top             =   5685
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
         Left            =   60
         Picture         =   "Frm_Cat_Varios.frx":3AD23
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "A"
         Top             =   5685
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Detalles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   75
         TabIndex        =   3
         Top             =   3750
         Width           =   13200
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
            Left            =   11850
            Picture         =   "Frm_Cat_Varios.frx":3E25A
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1455
            Width           =   1260
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   1170
            Left            =   135
            TabIndex        =   17
            Top             =   270
            Width           =   13005
            _ExtentX        =   22939
            _ExtentY        =   2064
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColorFixed  =   16777215
            BackColorBkg    =   16777215
         End
      End
      Begin VB.Frame Fra_generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   90
         TabIndex        =   1
         Top             =   765
         Width           =   13200
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1290
            TabIndex        =   13
            Top             =   1515
            Width           =   10710
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   1290
            TabIndex        =   12
            Top             =   915
            Width           =   10710
         End
         Begin VB.TextBox Txt_generales_nombre 
            Height          =   345
            Left            =   1290
            TabIndex        =   11
            Top             =   390
            Width           =   10695
         End
         Begin VB.Label Lbl_Ciudad 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ciudad"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   16
            Top             =   1545
            Width           =   885
         End
         Begin VB.Label Lbl_Direccion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   15
            Top             =   960
            Width           =   960
         End
         Begin VB.Label Lbl_nombre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   14
            Top             =   435
            Width           =   690
         End
      End
      Begin VB.Label Lbl_Nombre_Catalogo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Catálogo de Varios"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5055
         TabIndex        =   2
         Top             =   285
         Width           =   3510
      End
   End
End
Attribute VB_Name = "Frm_Cat_Referencia_Iconos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Sallir_Click()
Unload Frm_Cat_Varios
End Sub

Private Sub Form_Load()
    Me.Width = 13400
    Me.Height = 7500
End Sub

