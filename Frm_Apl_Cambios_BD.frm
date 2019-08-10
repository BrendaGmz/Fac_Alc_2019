VERSION 5.00
Begin VB.Form Frm_Apl_Cambios_BD 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CAMBIOS BD"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3360
   Begin VB.Frame Fra_Registrar_Cambios 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registrar Cambios en BD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Btn_Registra_Cambios 
         Caption         =   "Registrar Cambios"
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frm_Apl_Cambios_BD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Registra_Nuevos_Campos()

    MDIFrm_Apl_Principal.MousePointer = 11
    
        Mi_SQL = "ALTER TABLE Adm_Clientes_Facturas ALTER COLUMN Metodo_Pago varchar(50)"
        Conexion_Base.Execute Mi_SQL
        
    MDIFrm_Apl_Principal.MousePointer = 0
    
    MsgBox "Cambios realizados satisfactoriamente", vbInformation
    
End Sub

Private Sub Btn_Registra_Cambios_Click()
    If MsgBox("¿Esta seguro de registrar los movimientos?", vbQuestion + vbYesNo, "CONFIRMACIÓN") = vbYes Then
        Registra_Nuevos_Campos
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 2595
    Me.Width = 3600
    Me.Top = 200
    Me.Left = 2000
End Sub
