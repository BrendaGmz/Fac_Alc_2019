VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Baja_Facturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dar de Baja Facturas"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Dar_Baja 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dar de Baja"
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
      Left            =   480
      Picture         =   "Frm_Baja_Facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "A"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Btn_Mas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   240
      Width           =   390
   End
   Begin VB.PictureBox Pic_Entradas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   0
      ScaleHeight     =   5280
      ScaleWidth      =   7125
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Facturas 
         Height          =   2025
         Left            =   120
         TabIndex        =   10
         Top             =   2070
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   3572
         _Version        =   393216
         Rows            =   0
         Cols            =   7
         FixedRows       =   0
         BackColorBkg    =   16777215
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.CommandButton Btn_Imprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ver"
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
         Left            =   2160
         Picture         =   "Frm_Baja_Facturas.frx":3537
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "A"
         Top             =   4320
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
         Left            =   3720
         Picture         =   "Frm_Baja_Facturas.frx":69FD
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "C"
         Top             =   4320
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
         Left            =   5280
         Picture         =   "Frm_Baja_Facturas.frx":9F89
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "S"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Datos_Factura 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos de Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   120
         TabIndex        =   1
         Top             =   135
         Width           =   6945
         Begin VB.TextBox Txt_Metodo_Pago 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1365
            Width           =   2880
         End
         Begin VB.TextBox Txt_Forma_Pago 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1035
            Width           =   2880
         End
         Begin VB.TextBox Txt_Cliente 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   652
            Width           =   5445
         End
         Begin VB.TextBox Txt_Serie 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   293
            Width           =   1125
         End
         Begin VB.TextBox Txt_Folio 
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   293
            Width           =   1125
         End
         Begin VB.TextBox Txt_Saldo 
            Height          =   285
            Left            =   5325
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1365
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Cotizacion 
            Height          =   315
            Left            =   5325
            TabIndex        =   2
            Top             =   1035
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd MMM yyyy"
            Format          =   107216899
            CurrentDate     =   40444
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4560
            TabIndex        =   12
            Top             =   1365
            Width           =   540
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Factura"
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
            Left            =   3000
            TabIndex        =   11
            Top             =   293
            Width           =   1125
         End
         Begin VB.Label Lbl_Numero_Entrda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Serie"
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
            Left            =   45
            TabIndex        =   7
            Top             =   293
            Width           =   1125
         End
         Begin VB.Label Lbl_Proveedor_Entradas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   45
            TabIndex        =   6
            Top             =   652
            Width           =   1710
         End
         Begin VB.Label Lbl_Fecha 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4560
            TabIndex        =   5
            Top             =   1035
            Width           =   540
         End
         Begin VB.Label Lbl_Observaciones 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Método Pago"
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
            Left            =   45
            TabIndex        =   4
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label Lbl_Tipo_Entrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   45
            TabIndex        =   3
            Top             =   1035
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "Frm_Baja_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Buscar_Click()
    Dim Factura As String
    Factura = InputBox("Teclee el número de factura electrónica a consultar", "Consulta factura")
'    If Trim(factura) <> "" Then
       Llena_tabla (Factura)
       If Grid_Consulta_Facturas.Rows = 1 Or Grid_Consulta_Facturas.Rows = 0 Then
            Llena_tabla ("")
       End If
'    End If
End Sub

Private Sub Btn_Dar_Baja_Click()
    Dim Rs_Modifica As rdoResultset
    Mi_SQL = "SELECT Saldo,pagada,abono FROM Adm_Clientes_Facturas WHERE No_Factura_Electronica='" & Format(Txt_Folio, "0000000000") & "' AND Serie='" & Trim(Txt_Serie.text) & "'"
    Set Rs_Modifica = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    With Rs_Modifica
        If Not .EOF Then
            .Edit
                .rdoColumns("saldo") = 0
                .rdoColumns("pagada") = "S"
                .rdoColumns("abono") = .rdoColumns("abono") + Val(Txt_Saldo.text)
            .Update
            MsgBox "La factura se ha dado de baja correctamente"
            Llena_tabla ("")
        End If
    End With
End Sub

Private Sub Btn_Mas_Click()
If Btn_Mas.Caption = "+" Then
        Me.Top = 0
        Me.Left = 0
        Me.Width = MDIFrm_Apl_Principal.ScaleWidth
        Me.Height = MDIFrm_Apl_Principal.ScaleHeight
        
        Pic_Entradas.Width = MDIFrm_Apl_Principal.ScaleWidth - 400
        Pic_Entradas.Height = MDIFrm_Apl_Principal.ScaleHeight - 900
      
        Grid_Consulta_Facturas.Width = MDIFrm_Apl_Principal.ScaleWidth - 900
        Grid_Consulta_Facturas.Height = MDIFrm_Apl_Principal.ScaleHeight - 4000

        Fra_Datos_Factura.Width = Pic_Entradas.Width - 400
        Fra_Datos_Factura.Height = Pic_Entradas.Height - 100
        
        Btn_Dar_Baja.Top = MDIFrm_Apl_Principal.ScaleHeight - 1800
        Btn_Imprimir.Top = MDIFrm_Apl_Principal.ScaleHeight - 1800
        Btn_Buscar.Top = MDIFrm_Apl_Principal.ScaleHeight - 1800
        Btn_Salir.Top = MDIFrm_Apl_Principal.ScaleHeight - 1800
        
        
        Btn_Mas.Caption = "-"
    Else
        Me.Left = 100
        Me.Height = 5685
        Me.Width = 7230
        Grid_Consulta_Facturas.Width = 6945
        Grid_Consulta_Facturas.Height = 2025
        Fra_Datos_Factura.Width = 6915
        Fra_Datos_Factura.Height = 6400
        
        Pic_Entradas.Width = 7125
        Pic_Entradas.Height = 5280
        
        Btn_Dar_Baja.Top = 4200
        Btn_Imprimir.Top = 4200
        Btn_Buscar.Top = 4200
        Btn_Salir.Top = 4200
        
        Btn_Mas.Caption = "+"
    End If
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Llena_tabla ("")
End Sub
Private Function Llena_tabla(Factura As String)
    Dim Rs_Consulta As rdoResultset
    Dim Forma_Pago, Metodo_Pago As String
    Grid_Consulta_Facturas.Rows = 0
    Grid_Consulta_Facturas.AddItem "No. Factura" & Chr(9) & "Serie" & Chr(9) & "Cliente" & Chr(9) & "Forma Pago" & Chr(9) & "Método Pago" & Chr(9) & "Fecha" & Chr(9) & "Saldo"
    If Factura <> "" Then
        Mi_SQL = "SELECT Serie, No_Factura_Electronica, Forma_Pago,Tipo_Pago, Saldo, Fecha, Clientes.Nombre AS Cliente FROM Adm_Clientes_Facturas AS Facturas, Cat_Clientes AS Clientes WHERE No_Factura_Electronica='" & Format(Factura, "0000000000") & "' AND Clientes.Cliente_ID=Facturas.Cliente_ID AND Saldo>0 AND Pagada='N' AND Cancelada='N' AND substring(Tipo_Pago,1,3)='PUE' AND (substring(Forma_Pago,1,2)='03' OR substring(Forma_Pago,1,2)='01' OR substring(Forma_Pago,1,2)='02')  ORDER BY No_Factura"
    Else
        Mi_SQL = "SELECT Serie, No_Factura_Electronica, Forma_Pago,Tipo_Pago, Saldo, Fecha, Clientes.Nombre AS Cliente FROM Adm_Clientes_Facturas AS Facturas, Cat_Clientes AS Clientes WHERE Clientes.Cliente_ID=Facturas.Cliente_ID AND Saldo>0 AND Pagada='N' AND Cancelada='N' AND substring(Tipo_Pago,1,3)='PUE' AND (substring(Forma_Pago,1,2)='03' OR substring(Forma_Pago,1,2)='01' OR substring(Forma_Pago,1,2)='02')  ORDER BY No_Factura"
    End If
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta
        While Not .EOF
             Select Case (Mid(.rdoColumns("forma_pago"), 1, 2))
                Case "03"
                    Forma_Pago = "03 Transferencia Electrónica de fondos"
                Case "01"
                    Forma_Pago = "01 Efectivo"
                Case "02"
                    Forma_Pago = "02 Cheque Nominativo"
             End Select
             If Mid(.rdoColumns("tipo_pago"), 1, 3) = "PUE" Then
                Metodo_Pago = "PUE Pago en una sóla exhibición"
             Else
                Metodo_Pago = "PPD Pago en Parcialidades o diferido"
             End If
             Grid_Consulta_Facturas.AddItem .rdoColumns("No_Factura_Electronica") & Chr(9) & .rdoColumns("Serie") & Chr(9) & .rdoColumns("Cliente") & Chr(9) & Forma_Pago & Chr(9) & Metodo_Pago & Chr(9) & .rdoColumns("fecha") & Chr(9) & .rdoColumns("saldo")
             .MoveNext
             
        Wend
    End With
    Rs_Consulta.Close
    If Factura <> "" Or Grid_Consulta_Facturas.Rows < 2 Then
        Limpia_campos
        MsgBox "No hay datos para mostrar"
    Else
        Grid_Consulta_Facturas.RowSel = 1
        Grid_Consulta_Facturas_Click
'        Grid_Consulta_Facturas.FixedCols = 1
'        Grid_Consulta_Facturas.FixedRows = 1
        Grid_Consulta_Facturas.ColWidth(0) = 1000 'No_Factura
        Grid_Consulta_Facturas.ColWidth(1) = 0    'Serie
        Grid_Consulta_Facturas.ColWidth(2) = 5500 'Cliente
        Grid_Consulta_Facturas.ColWidth(3) = 3000  'forma pago
        Grid_Consulta_Facturas.ColAlignment(3) = flexAlignLeftCenter
        Grid_Consulta_Facturas.ColWidth(4) = 2500  'método pago
        Grid_Consulta_Facturas.ColWidth(5) = 1000  'fecha
        Grid_Consulta_Facturas.ColWidth(6) = 1000  'saldo
    End If
    
    
End Function

Private Sub Grid_Consulta_Facturas_Click()
     If Grid_Consulta_Facturas.RowSel > 0 Then
        Txt_Serie.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 1)
        Txt_Folio.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 0)
        Txt_Cliente.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 2)
        Txt_Forma_Pago.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 3)
        Txt_Metodo_Pago.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 4)
        Dtp_Fecha_Cotizacion.Value = (Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 5))
        Txt_Saldo.text = Grid_Consulta_Facturas.TextMatrix(Grid_Consulta_Facturas.RowSel, 6)
     End If
End Sub

Private Function Limpia_campos()
     Txt_Cliente.text = ""
     Txt_Folio.text = ""
     Txt_Forma_Pago.text = ""
     Txt_Metodo_Pago.text = ""
     Txt_Saldo.text = ""
     Txt_Serie.text = ""
     Dtp_Fecha_Cotizacion.Value = Now
End Function
Private Sub Btn_Imprimir_Click()
    Muestra_PDF
End Sub

Private Sub Muestra_PDF()
Dim Nombre_Archivo As String   'Almacena el nombre del archivo

On Error GoTo errorHandler

    If Txt_Folio.text <> "" Then
        MDIFrm_Apl_Principal.MousePointer = 11
            'Asigna el nombre del archivo
            Nombre_Archivo = "CFDI_" & Trim(Txt_Serie.text) & "_" & Val(Txt_Folio.text) & ".pdf"
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Pdfs & "\" & Nombre_Archivo, "ARCHIVO") = True Then
                'Envia para abrir el archivo
                ShellExecute ByVal 0&, "open", Ruta_Pdfs & "\" & Nombre_Archivo, vbNullString, vbNullString, SW_SHOWMAXIMIZED
            Else 'Regenera el pdf
                MsgBox "No se encontró el archivo PDF"
            End If
    Else
        MsgBox "Seleccione la factura", vbExclamation
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    If Err.Number = 70 Then
        MsgBox "El archivo PDF se encuentra abierto actualmente, favor de verificar", vbExclamation
    End If
End Sub
