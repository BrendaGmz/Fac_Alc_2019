VERSION 5.00
Begin VB.Form Frm_Apl_Enviando_Correo 
   ClientHeight    =   390
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Lbl_Mensaje_Esperando 
      Caption         =   "Enviando Correo......Favor de esperar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_Apl_Enviando_Correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Height = 510
    Me.Width = 4680
End Sub
