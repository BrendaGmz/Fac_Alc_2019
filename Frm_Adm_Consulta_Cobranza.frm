VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Cobranza_Consulta 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta de Cobranza"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   10065
   Begin VB.Frame Fra_Reimpresiones 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6390
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9990
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
         Left            =   135
         Picture         =   "Frm_Adm_Consulta_Cobranza.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "A"
         Top             =   5610
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Excel 
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
         Left            =   4215
         Picture         =   "Frm_Adm_Consulta_Cobranza.frx":3697
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "A"
         Top             =   5610
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
         Left            =   8460
         Picture         =   "Frm_Adm_Consulta_Cobranza.frx":5ED1
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "A"
         Top             =   5610
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Consultar 
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
         Left            =   8100
         Picture         =   "Frm_Adm_Consulta_Cobranza.frx":95D0
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "C"
         Top             =   930
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.TextBox Txt_Total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   2190
      End
      Begin VB.ComboBox Cmb_Con_Forma_Pago 
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Cobranza.frx":CB5C
         Left            =   7275
         List            =   "Frm_Adm_Consulta_Cobranza.frx":CB6C
         TabIndex        =   5
         Top             =   150
         Width           =   2190
      End
      Begin VB.ComboBox Cmb_Con_Banco 
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Cobranza.frx":CBA0
         Left            =   1275
         List            =   "Frm_Adm_Consulta_Cobranza.frx":CBA2
         TabIndex        =   2
         Top             =   525
         Width           =   4815
      End
      Begin VB.ComboBox Cmb_Con_Cliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   150
         Width           =   4800
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_1 
         Height          =   315
         Left            =   1275
         TabIndex        =   3
         Top             =   900
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   126550019
         CurrentDate     =   38038
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_2 
         Height          =   315
         Left            =   3900
         TabIndex        =   4
         Top             =   900
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   126550019
         CurrentDate     =   38038
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Cobranza 
         Height          =   3930
         Left            =   135
         TabIndex        =   7
         Top             =   1635
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   6932
         _Version        =   393216
         Rows            =   0
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   2
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin VB.Label Lbl_Total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   6225
         TabIndex        =   13
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Lbl_Forma_Pago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   6225
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Lbl_Banco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   675
         Width           =   465
      End
      Begin VB.Label Lbl_Fecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   3525
         TabIndex        =   9
         Top             =   1050
         Width           =   135
      End
      Begin VB.Label Lbl_Cliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   300
         Width           =   480
      End
   End
End
Attribute VB_Name = "Frm_Adm_Cobranza_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Cancelar_Click()
    Cancela_Movimientos
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cancela_Movimientos()
    'DESCRIPCIÓN: Cancela los movimientos que son realizados a las facturas
    'PARÁMETROS:
    'CREO:  Sergio Godínez Banda
    'FECHA_CREO:    14-Agosto-2007
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Cancela_Movimientos()
    Dim MiConsulta As New rdoQuery
    Dim MiConsulta2 As New rdoQuery
    Dim Rs_Movimientos As rdoResultset
    Dim Rs_MiTabla As rdoResultset
    Dim Rs_MiTabla_2 As rdoResultset
    Dim Rs_MiFactura As rdoResultset
    Dim Rs_MiComplemento As rdoResultset
    Dim MiAnticipo As rdoResultset
    Dim Importe As Double
    Dim Respuesta As Integer
    Dim Mi_SQL As String
    Dim No_Movimiento As String
    Dim no_pagos As Long
    Dim No_Complemento As String
    Dim I As Long
    Dim Mensaje As String
    Dim Cancel As Boolean

    On Error GoTo handler
    Conexion_Base.BeginTrans
        no_pagos = 0
        If Grid_Consulta_Cobranza.RowSel > 0 And (Trim(Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 8)) = "A" Or Trim(Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 8)) = "PC") Then
            Respuesta = MsgBox("¿Seguro de Cancelar el Cobro?", vbYesNo + vbQuestion, " MAGICAL BRIDE")
            If Respuesta = 6 Then
                'Consulta el número de complemento
                Mi_SQL = "SELECT No_Complemento_Pago "
                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 0) & "' and No_Complemento_Pago is not null"
                Set Rs_MiComplemento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_MiComplemento.EOF Then
                        No_Complemento = Rs_MiComplemento.rdoColumns("No_Complemento_Pago")
                End If
                Rs_MiComplemento.Close
                
                'Timbra la factura cancelada
                If No_Complemento <> "" Then
                    Set Conectar_Ayudante = New Ayudante
                    Mi_SQL = "SELECT *"
                    Mi_SQL = Mi_SQL & " FROM Complemento_Pago"
                    Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Format(No_Complemento, "0000000000") & "'"
                    Set Rs_MiComplemento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_MiComplemento.EOF Then
                        Rs_MiComplemento.Edit
                            If Trim(Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 8)) = "PC" Then
                                Mensaje = CFD_Cancela_Xml(Rs_MiComplemento.rdoColumns("Timbre_UUID"), True)
                            ElseIf Trim(Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 8)) = "A" Then
                                Mensaje = CFD_Cancela_Xml(Rs_MiComplemento.rdoColumns("Timbre_UUID"), False)
                            End If
                            If Mensaje Like "*cancelado*" Or Mensaje Like "*Cancelado*" Then
                                Cancel = True
                                Rs_MiComplemento.rdoColumns("Estatus") = "CANCELADA"
                            Else
                                Cancel = False
                                Rs_MiComplemento.rdoColumns("Estatus") = "CANCELACION PROCESO"
                            End If
                        Rs_MiComplemento.Update
                    End If
                    Rs_MiComplemento.Close
                End If
                
                'Modifica el movimiento
                Set Conectar_Ayudante = New Ayudante
                Mi_SQL = "SELECT *"
                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                Mi_SQL = Mi_SQL & " WHERE No_Complemento_Pago='" & Format(No_Complemento, "0000000000") & "' and Cliente_ID='" & Format(Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 7), "00000") & "'"
                Set Rs_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                With Rs_Movimientos
                    While Not .EOF
                        .Edit
                            If Cancel Then
                                Importe = Val(Conectar_Ayudante.Quitar_Caracter(.rdoColumns("Cantidad"), ","))
                                .rdoColumns("Cantidad") = 0
                                .rdoColumns("Estatus") = "C"
                                .rdoColumns("Usuario_Modifico") = Usuario_Sistema
                                .rdoColumns("Fecha_Modifico") = Now
                                No_Movimiento = .rdoColumns("No_Movimiento")
                                
                                'modifica la factura
                                Mi_SQL = "SELECT No_Factura, Abono, Saldo, Pagada, Fecha_Pago, Usuario_Modifico, Fecha_Modifico,No_Parcialidad"
                                Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas"
                                Mi_SQL = Mi_SQL & " WHERE No_Factura= '" & .rdoColumns("No_Factura") & "'"
                                Set Rs_MiTabla = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                                    If Not Rs_MiTabla.EOF Then
                                        Rs_MiTabla.Edit
                                            Rs_MiTabla.rdoColumns("Abono") = Rs_MiTabla.rdoColumns("Abono") - Importe
                                            Rs_MiTabla.rdoColumns("Saldo") = Rs_MiTabla.rdoColumns("Saldo") + Importe
                                            Rs_MiTabla.rdoColumns("Pagada") = "N"
                                            Rs_MiTabla.rdoColumns("No_Parcialidad") = Val(Rs_MiTabla.rdoColumns("No_Parcialidad")) - 1
                                            Rs_MiTabla.rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                            Rs_MiTabla.rdoColumns("Fecha_Modifico") = Now
                                        Rs_MiTabla.Update
                                    End If
                                Rs_MiTabla.Close
                            Else
'                                 .rdoColumns("Cantidad") = 0
                                .rdoColumns("Estatus") = "PC"
                                .rdoColumns("Usuario_Modifico") = Usuario_Sistema
                                .rdoColumns("Fecha_Modifico") = Now
                            End If
                            .rdoColumns("Mensaje_Cancelado") = Mensaje
                        .Update
                        .MoveNext
                    Wend
                End With
                    
                
                
       
                Conexion_Base.CommitTrans
'                Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 8) = "C"
'                Grid_Consulta_Cobranza.TextMatrix(Grid_Consulta_Cobranza.RowSel, 4) = 0
                Txt_Total.text = ""
                If Cancel Then
                    MsgBox "Cobranza Cancelada", vbExclamation, "MAGICAL BRIDE"
                Else
                    MsgBox "Cobranza en Proceso de Cancelación", vbExclamation, "MAGICAL BRIDE"
                End If
                Btn_Consultar_Click
            Else
                MsgBox "Operacion Cancelada", vbExclamation, "MAGICAL BRIDE"
            End If
        End If
    Exit Sub
handler:
    Conexion_Base.RollbackTrans
    If Err.Number = -255 Then
        MsgBox Err.Description
    Else
        For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
    End If
    
End Sub

Private Sub Btn_Consultar_Click()
'    If Not (Cmb_Con_Cliente.Text <> "" Or Cmb_Bancos.Text <> "" Or Cmb_Forma_Pago.Text <> "") and  Then
'        MsgBox "Seleccione alguna opción", vbInformation
'    Else
        Consulta_Movimientos
'    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Movimientos()
    'DESCRIPCIÓN: Consulta los movimientos registrados en la base de datos
    'PARÁMETROS:
    'CREO:  Sergio Godínez Banda
    'FECHA_CREO:    14-Agosto-2007
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Movimientos()
    Dim Rs_Facturas_Pendientes As rdoResultset
    Dim Rs_MiTabla As rdoResultset
    Dim MiProveedor As rdoResultset
    Dim Importe As Double
    Dim Mi_SQL As String
    Dim Texto As String
    
        'Consulta los vales pendientes de facturar
    On Error GoTo handler
  
        Set Conectar_Ayudante = New Ayudante
            Mi_SQL = "SELECT No_Movimiento,No_Complemento_Pago,Mensaje_Cancelado, Referencia, Fecha,Adm_Movimientos.Fecha_Creo, Cantidad, Concepto, Cat_Clientes.Nombre, Cat_Clientes.Cliente_ID,No_Remision, No_Factura, Forma_Pago, "
            Mi_SQL = Mi_SQL & " Adm_Movimientos.Cliente_ID, Adm_Movimientos.Estatus, Banco "
            Mi_SQL = Mi_SQL & " FROM Adm_Movimientos, Cat_Clientes "
            Mi_SQL = Mi_SQL & "WHERE Adm_Movimientos.Cliente_ID = Cat_Clientes.Cliente_ID"
            ''Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo_Pago ='I'"
'            Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Estatus='A'"
            
            If Cmb_Con_Banco.text <> "" Then
                Mi_SQL = Mi_SQL & " AND Banco = '" & Cmb_Con_Banco.text & "'"
            End If
            If Cmb_Con_Cliente.text <> "" Then
                If Cmb_Con_Cliente.ListIndex > -1 Then
                    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID = '" & Format(Cmb_Con_Cliente.ItemData(Cmb_Con_Cliente.ListIndex), "00000") & "'"
                Else
                    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID = ''"
                End If
            End If
            If Cmb_Con_Forma_Pago.text <> "" Then
                Mi_SQL = Mi_SQL & " AND Forma_Pago = '" & Cmb_Con_Forma_Pago.text & "'"
            End If
            Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo >= '" & Format(DTP_Fecha_1.Value, "yyyy-MM-dd 00:00:00") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo <= '" & Format(DTP_Fecha_2.Value, "yyyy-MM-dd 23:59:00") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Referencia, Fecha "
            Set Rs_Facturas_Pendientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Facturas_Pendientes.EOF Then
                With Rs_Facturas_Pendientes
                    Txt_Total.text = 0
                    Grid_Consulta_Cobranza.Rows = 0
                    Grid_Consulta_Cobranza.Cols = 11
                    'Cambia el encabezado de la columna referencia o no_cheque dependiendo la forma de pago
                    Texto = ""
                    If Rs_Facturas_Pendientes.rdoColumns("Forma_Pago") <> "" Or Rs_Facturas_Pendientes.rdoColumns("Forma_Pago") <> "Efectivo" Then
                            If Mid(Rs_Facturas_Pendientes.rdoColumns("Forma_Pago"), 1, 2) = "02" Then Texto = "No_Cheque"
                            If Mid(Rs_Facturas_Pendientes.rdoColumns("Forma_Pago"), 1, 2) = "03" Then Texto = "Referencia"
                            If Mid(Rs_Facturas_Pendientes.rdoColumns("Forma_Pago"), 1, 2) = "01" Then Texto = "Referencia"
                    End If
                    Grid_Consulta_Cobranza.AddItem "No. Movimiento" & Chr(9) & Texto & Chr(9) & "Fecha" & Chr(9) & "Cliente" _
                    & Chr(9) & "Cantidad" & Chr(9) & "Concepto" & Chr(9) & "No Documento" _
                    & Chr(9) & "Cliente_ID" & Chr(9) & "Estatus" & Chr(9) & "Banco" & Chr(9) & "Mensaje Cancelación"
                    
                    While Not Rs_Facturas_Pendientes.EOF
                        If Not IsNull(.rdoColumns("No_Factura")) Then
                            Grid_Consulta_Cobranza.AddItem .rdoColumns("No_Movimiento") & Chr(9) & .rdoColumns("Referencia") & Chr(9) _
                            & Format(.rdoColumns("Fecha_Creo"), "dd/MMM/yy") & Chr(9) & .rdoColumns("Nombre") & Chr(9) _
                            & Format(.rdoColumns("Cantidad"), "#,###,###.00") & Chr(9) & .rdoColumns("Concepto") & Chr(9) _
                            & .rdoColumns("No_Complemento_Pago") & Chr(9) & .rdoColumns("Cliente_ID") & Chr(9) _
                            & .rdoColumns("Estatus") & Chr(9) & .rdoColumns("Banco") & Chr(9) & .rdoColumns("Mensaje_Cancelado")
                            Txt_Total.text = Val(Txt_Total.text) + .rdoColumns("Cantidad")
                            Grid_Consulta_Cobranza.FixedRows = 1
                            Grid_Consulta_Cobranza.FixedCols = 1
                        Else
                            Grid_Consulta_Cobranza.AddItem .rdoColumns("No_Movimiento") & Chr(9) & .rdoColumns("Referencia") & Chr(9) _
                            & Format(.rdoColumns("Fecha_Creo"), "dd/MMM/yy") & Chr(9) & .rdoColumns("Nombre") & Chr(9) _
                            & Format(.rdoColumns("Cantidad"), "#,###,###.00") & Chr(9) & .rdoColumns("Concepto") & Chr(9) _
                            & .rdoColumns("No_Complemento_Pago") & Chr(9) & .rdoColumns("Cliente_ID") & Chr(9) _
                            & .rdoColumns("Estatus") & Chr(9) & .rdoColumns("Banco") & Chr(9) & .rdoColumns("Mensaje_Cancelado")
                            Txt_Total.text = Val(Txt_Total.text) + .rdoColumns("Cantidad")
                            Grid_Consulta_Cobranza.FixedRows = 1
                            Grid_Consulta_Cobranza.FixedCols = 1
                        End If
                        .MoveNext
                    Wend
                End With
                Rs_Facturas_Pendientes.Close
                'Txt_Total.Text = Format(Txt_Total.Text, "###,##0.00")
                Grid_Consulta_Cobranza.ColWidth(0) = 1200     'No_Movimiento
                Grid_Consulta_Cobranza.ColAlignment(0) = 3
                Grid_Consulta_Cobranza.ColWidth(1) = 1000     'No Cheque
                Grid_Consulta_Cobranza.ColAlignment(1) = 3
                Grid_Consulta_Cobranza.ColWidth(2) = 1000     'Fecha
                Grid_Consulta_Cobranza.ColAlignment(2) = 3
                Grid_Consulta_Cobranza.ColWidth(3) = 2000     'Cliente
                Grid_Consulta_Cobranza.ColWidth(4) = 800      'Monto
                Grid_Consulta_Cobranza.ColWidth(5) = 1500     'Concepto
                Grid_Consulta_Cobranza.ColWidth(6) = 1200     'No_Factura
                Grid_Consulta_Cobranza.ColAlignment(6) = 3
                Grid_Consulta_Cobranza.ColWidth(7) = 0        'Cliente_ID
                Grid_Consulta_Cobranza.ColWidth(8) = 600      'Estatus
                Grid_Consulta_Cobranza.ColAlignment(8) = 3
                Grid_Consulta_Cobranza.ColWidth(9) = 1000     'Banco
                Grid_Consulta_Cobranza.ColWidth(10) = 7000     'Cancelado
               
                Btn_Cancelar.Enabled = True
                Btn_Excel.Enabled = True
                Btn_Salir.Enabled = True
            Else
                Grid_Consulta_Cobranza.Rows = 0
                MsgBox "No Existen Cobranzas con esos datos", vbInformation
            End If
            Cmb_Con_Cliente.text = ""
            Cmb_Con_Banco.text = ""
            Cmb_Con_Forma_Pago.text = ""
            If Grid_Consulta_Cobranza.Rows = 0 Then Btn_Cancelar.Enabled = False
    Exit Sub
handler:
    cn.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Btn_Excel_Click()
    'DESCRIPCIÓN: Exporta los datos a una hoja de excel
    'PARÁMETROS:
    'CREO:  Sergio Godínez Banda
    'FECHA_CREO:    13-Agosto-2007
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Excel_Click()
    Dim RutaArchivo As String

      ' Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
        ' Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
        ' Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.Xls|"
        ' Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
        ' Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
        ' Display name of selected file

    RutaArchivo = MDIFrm_Apl_Principal.CommonDialog1.fileName
    
    Open RutaArchivo For Output As #1
        For I = 0 To Grid_Consulta_Cobranza.Rows - 1
            For J = 0 To Grid_Consulta_Cobranza.Cols - 1
                CABECERA = CABECERA & Grid_Consulta_Cobranza.TextMatrix(I, J) & Chr(9)
            Next J
            Print #1, CABECERA
            CABECERA = ""
        Next I
    Close #1
    MsgBox "Reporte Importado"
    Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub Btn_Salir_Click()
    Unload Frm_Adm_Cobranza_Consulta
End Sub



Private Sub Cmb_Con_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Con_Banco, 1, "Nombre")
    End If
End Sub


Private Sub Form_Load()
    Me.Top = 0
    Me.Height = 7050
    Me.Width = 10275
    Consulta_Bancos
    DTP_Fecha_1.Value = Now
    DTP_Fecha_2.Value = Now
    Btn_Cancelar.Enabled = False
    Call Cmb_Con_Banco_KeyPress(13)
    Call Cmb_Con_Cliente_KeyPress(13)
    Consulta_Cancelados
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Bancos
    'DESCRIPCIÓN: Consulta los datos de la tabla Cat_Bancos
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 13-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Bancos()
    Dim Mi_SQL As String
    Dim Rs_MiTabla As rdoResultset
    
    Set Conectar_Ayudante = New Ayudante
            'Consulta del catalogo de Bancos
        Mi_SQL = "SELECT DISTINCT Banco,Tipo_Pago "
        Mi_SQL = Mi_SQL & " FROM Adm_Movimientos "
        Mi_SQL = Mi_SQL & " WHERE Tipo_Pago= 'I' "
        Mi_SQL = Mi_SQL & " ORDER BY Banco "
        
        Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                'Consulta del catalogo de Clientes
            Cmb_Con_Banco.Clear
            If Not Rs_MiTabla.EOF Then
                While Not Rs_MiTabla.EOF
                    Cmb_Con_Banco.AddItem Rs_MiTabla!Banco
                    Rs_MiTabla.MoveNext
                Wend
                Rs_MiTabla.Close
            End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cmb_Con_Cliente_KeyPress()
    'DESCRIPCIÓN: Al hacer click en el combo muestra los nombres de los cliente segun la consulta a la tabla Cat_Clientes
    'PARÁMETROS:
    'CREO:  Sergio Godínez Banda
    'FECHA_CREO:    8-Agosto-2007
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Cmb_Con_Cliente_KeyPress(KeyAscii As Integer)
Dim Mi_SQL As String
Dim Rs_MiTabla As rdoResultset
    
    Set Conectar_Ayudante = New Ayudante
    'Consulta del catalogo de Clientes
    Mi_SQL = "SELECT Cliente_ID, Nombre "
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes "
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Cmb_Con_Cliente.text & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre "
    Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            'Consulta del catalogo de Clientes
        Cmb_Con_Cliente.Clear
        If Not Rs_MiTabla.EOF Then
            While Not Rs_MiTabla.EOF
                Cmb_Con_Cliente.AddItem Rs_MiTabla!Nombre
                Cmb_Con_Cliente.ItemData(Cmb_Con_Cliente.NewIndex) = Rs_MiTabla!Cliente_ID
                Rs_MiTabla.MoveNext
            Wend
        Rs_MiTabla.Close
        End If
End Sub



