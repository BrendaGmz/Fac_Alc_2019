VERSION 5.00
Begin VB.Form Frm_Apl_Registro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro del Software"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6540
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   375
      Width           =   855
   End
   Begin VB.CommandButton Btn_Registrar 
      Caption         =   "Registrar"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Fram_Datos_Configuracion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de Configuracion"
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
      Left            =   75
      TabIndex        =   22
      Top             =   2325
      Width           =   6375
      Begin VB.TextBox Txt_Usuario 
         Height          =   285
         Left            =   1095
         TabIndex        =   10
         Top             =   1080
         Width           =   2385
      End
      Begin VB.TextBox Txt_Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4680
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_Tipo_Base 
         Height          =   315
         ItemData        =   "Frm_Cfg_Registro_Sistema.frx":0000
         Left            =   1095
         List            =   "Frm_Cfg_Registro_Sistema.frx":000A
         TabIndex        =   8
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox Txt_Servidor 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox Txt_Base_Datos 
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         Height          =   195
         Index           =   9
         Left            =   3750
         TabIndex        =   27
         Top             =   825
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Base"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Servidor:"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Datos"
         Height          =   195
         Index           =   7
         Left            =   3750
         TabIndex        =   23
         Top             =   420
         Width           =   825
      End
   End
   Begin VB.Frame Fram_Datos_Empresa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la Empresa"
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
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   6330
      Begin VB.TextBox Txt_Estado 
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Txt_Ciudad 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Txt_CP 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Txt_Colonia 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Txt_Direccion 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox Txt_RFC 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R.F.C."
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
         Height          =   195
         Index           =   5
         Left            =   3600
         TabIndex        =   20
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   600
      End
   End
End
Attribute VB_Name = "Frm_Apl_Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Registrar_Click()
'Contador para especificar la linea en la cual se va a ir guardando los datos
Dim I As Integer
Set Conectar_Ayudante = New Ayudante
'Abre el documento llamado Config.ini y empiesa a escribir lo contenido
'en las cajas de texto
Open App.Path & "\Config.ini" For Output As #1
    Print #1, Txt_Nombre.Text
    Print #1, Txt_RFC.Text
    Print #1, Txt_Direccion.Text
    Print #1, Txt_Colonia.Text
    Print #1, Txt_CP.Text
    Print #1, Txt_Ciudad.Text
    Print #1, Txt_Estado.Text
    Print #1, Cmb_Tipo_Base.Text
    Print #1, Txt_Servidor.Text
    Print #1, Txt_Usuario.Text
    Print #1, Txt_Base_Datos.Text
    Print #1, Txt_Password.Text
Close #1
MsgBox "Datos Registrados"
Unload Frm_Apl_Registro
Conexion_Base.Close
Conectar_Ayudante.Conexion 'Manda llamar la función Conexion contenida en el Module1
End Sub




Private Sub Form_Load()
Dim I As Integer        'Contador que indica que linea se esta leyendo del documento
Dim Linea As String     'Guarda el valor de la linea

Me.Left = (Screen.Width - Me.Width) \ 2
Me.Top = (Screen.Height - Me.Height) \ 2

On Error GoTo Etiqueta
    Open App.Path & "\Config.ini" For Input As #1
    I = 0
    '1. Llena las cajas de texto de acuerdo a lo contenido en el documento Config.ini
    Do While Not EOF(1)
        Line Input #1, Linea
        If I <= 6 Then
            If I = 0 Then
                Txt_Nombre.Text = Linea
            ElseIf I = 1 Then
                Txt_RFC.Text = Linea
            ElseIf I = 2 Then
                Txt_Direccion = Linea
            ElseIf I = 3 Then
                Txt_Colonia.Text = Linea
            ElseIf I = 4 Then
                Txt_CP.Text = Linea
            ElseIf I = 5 Then
                Txt_Ciudad.Text = Linea
            ElseIf I = 6 Then
                Txt_Estado.Text = Linea
            End If
        Else
            If I <= 14 Then
                If I = 7 Then
                    Cmb_Tipo_Base.Text = Linea
                ElseIf I = 8 Then
                    Txt_Servidor.Text = Linea
                ElseIf I = 9 Then
                    Txt_Usuario.Text = Linea
                ElseIf I = 10 Then
                    Txt_Base_Datos.Text = Linea
                ElseIf I = 11 Then
                    Txt_Password.Text = Linea
                End If
            End If
        End If
    I = I + 1
    Loop
Close #1
Exit Sub

Etiqueta:
    MsgBox "El sistema no se encuentra registrado" & Chr(13) & "Favor de llenar sus datos", vbExclamation
End Sub
