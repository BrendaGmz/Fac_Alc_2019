VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Clientes_Facturas_Consulta 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta Facturas de Clientes"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9165
   Begin VB.Frame Fra_Busqueda_Facturas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Busqueda de Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   9090
      Begin VB.CheckBox Chk_Remisiones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remisiones"
         Height          =   195
         Left            =   2100
         TabIndex        =   23
         Top             =   2100
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox Chk_Historico 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Histórico Facturas"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2100
         Width           =   1965
      End
      Begin VB.TextBox Txt_Con_Numero_De_Remision 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   21
         Top             =   225
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CheckBox Chk_Numero_Remision 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Por No. de Remision"
         Height          =   195
         Left            =   4590
         TabIndex        =   20
         Top             =   270
         Width           =   1830
      End
      Begin VB.CommandButton Btn_Busqueda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar Pagos"
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
         Index           =   0
         Left            =   5985
         Picture         =   "Frm_Adm_Consulta_Facturas_Cliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1237
         Width           =   3045
      End
      Begin VB.CheckBox Chk_No_Factura 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Por No. de Factura"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Width           =   1965
      End
      Begin VB.CheckBox Chk_Cliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Por Cliente"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Width           =   1965
      End
      Begin VB.TextBox Txt_Con_No_Factura 
         Height          =   285
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox Cmb_Con_Cliente 
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   6915
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   390
      End
      Begin VB.TextBox Txt_Total 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1665
         Width           =   1215
      End
      Begin VB.TextBox Txt_Abonos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1665
         Width           =   1215
      End
      Begin VB.TextBox Txt_Saldos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7500
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1665
         Width           =   1530
      End
      Begin VB.CommandButton Btn_Excel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1545
         Width           =   840
      End
      Begin VB.CheckBox Chk_Estatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Por Estatus"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   1305
         Width           =   1665
      End
      Begin VB.ComboBox Cmb_Con_Estatus 
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Facturas_Cliente.frx":0295
         Left            =   2100
         List            =   "Frm_Adm_Consulta_Facturas_Cliente.frx":029F
         TabIndex        =   3
         Top             =   1275
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CheckBox Chk_Fecha_Factura 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Por Fecha Factura"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   945
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_Factura_1 
         Height          =   315
         Left            =   2115
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   106037251
         CurrentDate     =   38038
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Facturas 
         Height          =   4035
         Left            =   120
         TabIndex        =   10
         Top             =   2310
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   7117
         _Version        =   393216
         Rows            =   0
         Cols            =   12
         FixedRows       =   0
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_Factura_2 
         Height          =   315
         Left            =   5985
         TabIndex        =   15
         Top             =   885
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   106037251
         CurrentDate     =   38038
      End
      Begin VB.Label Lbl_Total 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
         Height          =   240
         Left            =   1530
         TabIndex        =   18
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Lbl_Abonos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Abonos"
         Height          =   240
         Left            =   3330
         TabIndex        =   17
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Lbl_Saldos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldos"
         Height          =   240
         Left            =   5985
         TabIndex        =   16
         Top             =   1710
         Width           =   615
      End
   End
End
Attribute VB_Name = "Frm_Adm_Clientes_Facturas_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN   : Btn_Consultar_Click
'DESCRIPCIÓN            : Consulta las facturas existentes en la base de datos
'PARÁMETROS             :
'CREO                   : Julio Cruz
'FECHA_CREO             : 19-Enero-2011
'MODIFICO               :
'FECHA_MODIFICO         :
'CAUSA_MODIFICACIÓN     :
'*******************************************************************************
Private Sub Btn_Busqueda_Click(Index As Integer)
Dim MiConsulta As New rdoQuery
Dim Rs_MiTabla As rdoResultset
Dim Columnas As Integer
Dim Renglon As Integer
Dim MiConsulta2 As New rdoQuery
Dim Rs_MisAnticipos As rdoResultset
Dim Mi_SQL As String
Dim Folio_Electronico As String
Dim Resultados As Boolean


    'Consulta las facturas
    Resultados = False
    Txt_Total.text = ""
    Txt_Abonos.text = ""
    Txt_Saldos.text = ""
    Grid_Consulta_Facturas.Rows = 0
    Grid_Consulta_Facturas.Cols = 8
    Columnas = 8
    Grid_Consulta_Facturas.AddItem "Fecha" & Chr(9) & "No. Documento" & Chr(9) & "Electrónica" & Chr(9) & "Cliente" & _
        Chr(9) & "Total" & Chr(9) & "Saldo" & Chr(9) & "Abono" & Chr(9) & "Cancelada"
    Renglon = 1
        If Not Chk_Numero_Remision.Value = 1 And Chk_Remisiones.Value = 0 Then
            'CONSULTA LAS FACTURAS
            With MiConsulta
            Set Conectar_Ayudante = New Ayudante
                Mi_SQL = "SELECT No_Factura, No_Factura_Electronica, Cancelada, Fecha, Total, Saldo, Abono,"
                Mi_SQL = Mi_SQL & " Adm_Clientes_Facturas.Cliente_ID, Cat_Clientes.Nombre"
                Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Facturas, Cat_Clientes "
                Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Facturas.Cliente_ID = Cat_Clientes.Cliente_ID "
                If Chk_No_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND No_Factura_Electronica = '" & Format(Txt_Con_No_Factura.text, "0000000000") & "'"
                If Chk_Cliente.Value = 1 Then
                    If Cmb_Con_Cliente.ListIndex > -1 Then
                        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cliente_ID = '" & Format(Cmb_Con_Cliente.ItemData(Cmb_Con_Cliente.ListIndex), "00000") & "'"
                    Else
                        Mi_SQL = Mi_SQL & " AND Adm_Clientes_Facturas.Cliente_ID = ''"
                    End If
                End If
                If Chk_Fecha_Factura.Value = 1 And Chk_Historico.Value = 0 Then Mi_SQL = Mi_SQL & " AND (Fecha >= '" & Format(DTP_Fecha_Factura_1.Value, "MM/dd/yyyy") & "' AND Fecha <= '" & Format(DTP_Fecha_Factura_2.Value, "MM/dd/yyyy") & "' AND Fecha >= '" & Format("01-01-2018", "MM/dd/yyyy") & "')"
                If Chk_Fecha_Factura.Value = 0 And Chk_Historico.Value = 0 Then Mi_SQL = Mi_SQL & " AND (Fecha >= '" & Format("01-01-2018", "MM/dd/yyyy") & "')"
                If Chk_Estatus.Value = 1 And Cmb_Con_Estatus.text = "Pagadas" Then Mi_SQL = Mi_SQL & " AND Pagada = 'S'"
                If Chk_Estatus.Value = 1 And Cmb_Con_Estatus.text = "Sin Pagar" Then Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
                
                Mi_SQL = Mi_SQL & " ORDER BY No_Factura_Electronica "
                Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            End With
            
                If Not Rs_MiTabla.EOF Then
                    With Rs_MiTabla
                        Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        While Not Rs_MiTabla.EOF
                            If Not IsNull(Rs_MiTabla.rdoColumns("No_Factura_Electronica")) Then
                                Folio_Electronico = Val(Rs_MiTabla.rdoColumns("No_Factura_Electronica"))
                            Else
                                Folio_Electronico = ""
                            End If
                            Grid_Consulta_Facturas.AddItem Format(Rs_MiTabla.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) _
                                & Rs_MiTabla.rdoColumns("No_Factura") & Chr(9) & Folio_Electronico & Chr(9) _
                                & Rs_MiTabla.rdoColumns("Nombre") & Chr(9) _
                                & Format(Rs_MiTabla.rdoColumns("Total"), "###,##0.00") & Chr(9) _
                                & Format(Rs_MiTabla.rdoColumns("Saldo"), "###,##0.00") & Chr(9) _
                                & Format(Rs_MiTabla.rdoColumns("Abono"), "###,##0.00") & Chr(9) _
                                & Rs_MiTabla.rdoColumns("Cancelada")
                            Txt_Total.text = Val(Txt_Total.text) + Rs_MiTabla.rdoColumns("Total")
                            If Trim(Rs_MiTabla.rdoColumns("Cancelada")) = "N" Then Txt_Saldos.text = Val(Txt_Saldos.text) + Rs_MiTabla.rdoColumns("Saldo")
                            Grid_Consulta_Facturas.FixedRows = 1
                            'CONSULTA LOS PAGOS DE LA FACTURA
                            With MiConsulta2
                                Set Conectar_Ayudante = New Ayudante
                                Mi_SQL = "SELECT No_Movimiento, Concepto, Referencia, Fecha, Cantidad, Forma_Pago, Banco "
                                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                                Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Tipo_Pago = 'I' "
                                Mi_SQL = Mi_SQL & " AND No_Factura = '" & Rs_MiTabla.rdoColumns("No_Factura") & "'"
                                Mi_SQL = Mi_SQL & " AND Estatus = 'A'"
                                Mi_SQL = Mi_SQL & " ORDER BY Fecha "
                            
                                Set Rs_MisAnticipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            End With
                        
                            With Rs_MisAnticipos
                                Set Conectar_Ayudante = New Ayudante
                                While Not Rs_MisAnticipos.EOF
                                    If Grid_Consulta_Facturas.Cols <= Columnas Then
                                        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 6
                                        Columnas = Grid_Consulta_Facturas.Cols
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 6) = "Concepto"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 5) = "Forma_Pago"
                                        If Rs_MisAnticipos.rdoColumns("Forma_Pago") <> "" Or Rs_MisAnticipos.rdoColumns("Forma_Pago") <> "Efectivo" Then
                                            If Rs_MisAnticipos.rdoColumns("Forma_Pago") = "Transferencia" Then
                                                Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "Referencia"
                                            Else
                                                Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "No_Cheque"
                                            End If
                                        End If
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 3) = "Banco"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 2) = "Fecha"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 1) = "Cantidad"
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 6) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 5) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 4) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 3) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 2) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 1) = 1200
                                    Else
                                        Columnas = Columnas + 6
                                    End If
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 6) = Rs_MisAnticipos.rdoColumns("Concepto")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 5) = Rs_MisAnticipos.rdoColumns("Forma_Pago")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = Rs_MisAnticipos.rdoColumns("Referencia")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 3) = Rs_MisAnticipos.rdoColumns("Banco")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 2) = Format(Rs_MisAnticipos.rdoColumns("Fecha"), "dd/MMM/yyyy")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 1) = Format(Rs_MisAnticipos.rdoColumns("Cantidad"), "###,##0.00")
                                    Txt_Abonos.text = Val(Txt_Abonos.text) + .rdoColumns("Cantidad")
                                    .MoveNext
                                Wend
                            End With
                            Columnas = 8
                            Renglon = Renglon + 1
                            Rs_MisAnticipos.Close
                            Rs_MiTabla.MoveNext
                            Resultados = True
                        Wend
                        Grid_Consulta_Facturas.ColWidth(0) = 1000   'Fecha
                        Grid_Consulta_Facturas.ColAlignment(0) = 3
                        Grid_Consulta_Facturas.ColWidth(1) = 1200   'Documento
                        Grid_Consulta_Facturas.ColAlignment(1) = 3
                        Grid_Consulta_Facturas.ColWidth(2) = 1100   'Electronica
                        Grid_Consulta_Facturas.ColAlignment(2) = 3
                        Grid_Consulta_Facturas.ColWidth(3) = 2000   'Cliente
                        Grid_Consulta_Facturas.ColAlignment(3) = 1
                        Grid_Consulta_Facturas.ColWidth(4) = 850    'Total
                        Grid_Consulta_Facturas.ColWidth(5) = 850    'Saldo
                        Grid_Consulta_Facturas.ColWidth(6) = 850    'Abono
                        Grid_Consulta_Facturas.ColWidth(7) = 700    'Cancelada
                        Txt_Abonos.text = Format(Val(Txt_Total.text) - Val(Txt_Saldos.text), "###,###,###.00")
                        Txt_Total.text = Format(Txt_Total.text, "###,###,###.00")
                        Txt_Saldos.text = Format(Txt_Saldos.text, "###,###,###.00")
                End With
                Rs_MiTabla.Close
            Else
'                MsgBox "No se encontraron documentos con esos datos", vbInformation
            End If
        End If
        
        
        If Not Chk_No_Factura.Value = 1 And Chk_Remisiones.Value = 1 Then
            'CONSULTA LAS REMISIONES
            With MiConsulta
                Set Conectar_Ayudante = New Ayudante
                    Mi_SQL = "SELECT No_Remision,Cancelada, Fecha, Total, Saldo,Abono, Adm_Clientes_Remisiones.Cliente_ID,Cat_Clientes.Nombre "
                    Mi_SQL = Mi_SQL & " FROM Adm_Clientes_Remisiones, Cat_Clientes "
                    Mi_SQL = Mi_SQL & " WHERE Adm_Clientes_Remisiones.Cliente_ID = Cat_Clientes.Cliente_ID "
                    If Chk_Numero_Remision.Value = 1 Then Mi_SQL = Mi_SQL & " AND No_Remision = '" & Format(Txt_Con_Numero_De_Remision.text, "0000000000") & "'"
                    If Chk_Cliente.Value = 1 Then
                        If Cmb_Con_Cliente.ListIndex > -1 Then
                            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cliente_ID = '" & Format(Cmb_Con_Cliente.ItemData(Cmb_Con_Cliente.ListIndex), "00000") & "'"
                        Else
                            Mi_SQL = Mi_SQL & " AND Adm_Clientes_Remisiones.Cliente_ID = ''"
                        End If
                    End If
                    If Chk_Fecha_Factura.Value = 1 Then Mi_SQL = Mi_SQL & " AND Fecha >= '" & Format(DTP_Fecha_Factura_1.Value, "MM/dd/yyyy") & "' AND Fecha <= '" & Format(DTP_Fecha_Factura_2.Value, "MM/dd/yyyy") & "'"
                    If Chk_Estatus.Value = 1 And Cmb_Con_Estatus.text = "Pagadas" Then Mi_SQL = Mi_SQL & " AND Pagada = 'S'"
                    If Chk_Estatus.Value = 1 And Cmb_Con_Estatus.text = "Sin Pagar" Then Mi_SQL = Mi_SQL & " AND Pagada = 'N'"
                    Mi_SQL = Mi_SQL & " ORDER BY No_Remision "
                    Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            End With
            If Not Rs_MiTabla.EOF Then
                With Rs_MiTabla
                    Set Rs_MiTabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        Renglon = 1
                        While Not Rs_MiTabla.EOF
                            Grid_Consulta_Facturas.AddItem Format(Rs_MiTabla.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Rs_MiTabla.rdoColumns("No_Remision") _
                            & Chr(9) & "" & Chr(9) & Rs_MiTabla.rdoColumns("Nombre") & Chr(9) & Format(Rs_MiTabla.rdoColumns("Total"), "###,##0.00") & Chr(9) & Format(Rs_MiTabla.rdoColumns("Saldo"), "###,##0.00") _
                            & Chr(9) & Format(Rs_MiTabla.rdoColumns("Abono"), "###,##0.00") & Chr(9) & Rs_MiTabla.rdoColumns("Cancelada")
                            Txt_Total.text = Val(Txt_Total.text) + Rs_MiTabla.rdoColumns("Total")
                            If Trim(Rs_MiTabla.rdoColumns("Cancelada")) = "N" Then Txt_Saldos.text = Val(Txt_Saldos.text) + Rs_MiTabla.rdoColumns("Saldo")
                            Grid_Consulta_Facturas.FixedRows = 1
                            'CONSULTA LOS PAGOS DE LA FACTURA
                            With MiConsulta2
                                Set Conectar_Ayudante = New Ayudante
                                Mi_SQL = "SELECT No_Movimiento, Concepto, Referencia, Fecha, Cantidad, Forma_Pago, Banco "
                                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                                Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Tipo_Pago = 'I' "
                                Mi_SQL = Mi_SQL & " AND No_Remision = '" & Rs_MiTabla.rdoColumns("No_Remision") & "'"
                                Mi_SQL = Mi_SQL & " AND Estatus = 'A'"
                                Mi_SQL = Mi_SQL & " ORDER BY Fecha "
                                Set Rs_MisAnticipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            End With
                            With Rs_MisAnticipos
                                Set Conectar_Ayudante = New Ayudante
                                While Not Rs_MisAnticipos.EOF
                                    If Grid_Consulta_Facturas.Cols <= Columnas Then
                                        Grid_Consulta_Facturas.Cols = Grid_Consulta_Facturas.Cols + 6
                                        Columnas = Grid_Consulta_Facturas.Cols
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 6) = "Concepto"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 5) = "Forma_Pago"
                                        If Rs_MisAnticipos.rdoColumns("Forma_Pago") <> "" Or Rs_MisAnticipos.rdoColumns("Forma_Pago") <> "Efectivo" Then
                                            If Rs_MisAnticipos.rdoColumns("Forma_Pago") = "Transferencia" Then
                                                Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "Referencia"
                                            Else
                                                Grid_Consulta_Facturas.TextMatrix(0, Columnas - 4) = "No_Cheque"
                                            End If
                                        End If
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 3) = "Banco"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 2) = "Fecha"
                                        Grid_Consulta_Facturas.TextMatrix(0, Columnas - 1) = "Cantidad"
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 6) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 5) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 4) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 3) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 2) = 1200
                                        Grid_Consulta_Facturas.ColWidth(Columnas - 1) = 1200
                                    Else
                                        Columnas = Columnas + 6
                                    End If
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 6) = Rs_MisAnticipos.rdoColumns("Concepto")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 5) = Rs_MisAnticipos.rdoColumns("Forma_Pago")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 4) = Rs_MisAnticipos.rdoColumns("Referencia")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 3) = Rs_MisAnticipos.rdoColumns("Banco")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 2) = Format(Rs_MisAnticipos.rdoColumns("Fecha"), "dd/MMM/yyyy")
                                    Grid_Consulta_Facturas.TextMatrix(Renglon, Columnas - 1) = Format(Rs_MisAnticipos.rdoColumns("Cantidad"), "###,##0.00")
                                    Txt_Abonos.text = Val(Txt_Abonos.text) + .rdoColumns("Cantidad")
                                    .MoveNext
                                Wend
                            End With
                            Columnas = 7
                            Renglon = Renglon + 1
                            Rs_MisAnticipos.Close
                            Rs_MiTabla.MoveNext
                            Resultados = True
                        Wend
                        Grid_Consulta_Facturas.ColWidth(0) = 1000   'Fecha
                        Grid_Consulta_Facturas.ColAlignment(0) = 3
                        Grid_Consulta_Facturas.ColWidth(1) = 1200   'Documento
                        Grid_Consulta_Facturas.ColAlignment(1) = 3
                        Grid_Consulta_Facturas.ColWidth(2) = 1100   'Electronica
                        Grid_Consulta_Facturas.ColAlignment(2) = 3
                        Grid_Consulta_Facturas.ColWidth(3) = 2000   'Cliente
                        Grid_Consulta_Facturas.ColAlignment(3) = 1
                        Grid_Consulta_Facturas.ColWidth(4) = 850    'Total
                        Grid_Consulta_Facturas.ColWidth(5) = 850    'Saldo
                        Grid_Consulta_Facturas.ColWidth(6) = 850    'Abono
                        Grid_Consulta_Facturas.ColWidth(7) = 700    'Cancelada
                        Txt_Abonos.text = Format(Val(Txt_Total.text) - Val(Txt_Saldos.text), "###,###,###.00")
                        Txt_Total.text = Format(Txt_Total.text, "###,###,###.00")
                        Txt_Saldos.text = Format(Txt_Saldos.text, "###,###,###.00")
                    End With
                    Rs_MiTabla.Close
                Else
'                    MsgBox "No se encontraron documentos con esos datos", vbInformation
                End If
        End If
        If Resultados = False Then
            MsgBox "No se encontraron documentos con esos datos", vbInformation
        End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Btn_Excel_Click()
    'DESCRIPCIÓN: Exporta los datos a una hoja de excel
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 8-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
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
        For I = 0 To Grid_Consulta_Facturas.Rows - 1
            For J = 0 To Grid_Consulta_Facturas.Cols - 1
                CABECERA = CABECERA & Grid_Consulta_Facturas.TextMatrix(I, J) & Chr(9)
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

Private Sub Chk_No_Factura_Click()
    If Chk_No_Factura.Value = 1 Then
        Txt_Con_No_Factura.Visible = True
        Txt_Con_No_Factura.SetFocus
    Else
        Txt_Con_No_Factura.Visible = False
    End If
End Sub

Private Sub Chk_Cliente_Click()
    If Chk_Cliente.Value = 1 Then
        Cmb_Con_Cliente.Visible = True
        Cmb_Con_Cliente.SetFocus
    Else
        Cmb_Con_Cliente.Visible = False
    End If
End Sub

Private Sub Chk_Estatus_Click()
    If Chk_Estatus.Value = 1 Then
        Cmb_Con_Estatus.Visible = True
        Cmb_Con_Estatus.SetFocus
    Else
        Cmb_Con_Estatus.Visible = False
    End If
End Sub

Private Sub Chk_Fecha_Factura_Click()
    If Chk_Fecha_Factura.Value = 1 Then
        DTP_Fecha_Factura_1.Visible = True
        DTP_Fecha_Factura_2.Visible = True
        DTP_Fecha_Factura_1.SetFocus
    Else
        DTP_Fecha_Factura_1.Visible = False
        DTP_Fecha_Factura_2.Visible = False
    End If
End Sub

Private Sub Chk_Numero_Remision_Click()
    If Chk_Numero_Remision.Value = 1 Then
        Txt_Con_Numero_De_Remision.Visible = True
        Txt_Con_Numero_De_Remision.SetFocus
    Else
        Txt_Con_Numero_De_Remision.Visible = False
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cmb_Con_Clientes_KeyPress()
    'DESCRIPCIÓN: Consulta
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 8-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
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

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Btn_Mas_Click()
    'DESCRIPCIÓN: Expande o reduce el tamaño de la ventana
    'PARÁMETROS :
    'CREO       : Sergio Godínez Banda
    'FECHA_CREO : 8-Agosto-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Mas_Click()
    If Btn_Mas.Caption = "+" Then
        Me.Top = 0
        Me.Left = 0
        Me.Width = MDIFrm_Apl_Principal.ScaleWidth
        Me.Height = MDIFrm_Apl_Principal.ScaleHeight
        Grid_Consulta_Facturas.Width = MDIFrm_Apl_Principal.ScaleWidth - 700
        Grid_Consulta_Facturas.Height = MDIFrm_Apl_Principal.ScaleHeight - 3500
        Fra_Busqueda_Facturas.Width = MDIFrm_Apl_Principal.ScaleWidth - 400
        Fra_Busqueda_Facturas.Height = MDIFrm_Apl_Principal.ScaleHeight - 1000
        Btn_Mas.Caption = "-"
    Else
        Me.Left = 1000
        Me.Height = 7065
        Me.Width = 7230
        Grid_Consulta_Facturas.Width = 6500
        Grid_Consulta_Facturas.Height = 4275
        Fra_Busqueda_Facturas.Width = 6915
        Fra_Busqueda_Facturas.Height = 6485
        Btn_Mas.Caption = "+"
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 1000
    Me.Height = 7065
    Me.Width = 9345
    DTP_Fecha_Factura_1.Value = Now
    DTP_Fecha_Factura_2.Value = Now
    If Rol_ID = "00001" Then
        Chk_Remisiones.Visible = True
    End If
End Sub


Private Sub Txt_Con_No_Factura_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Con_No_Factura.text, False)
End Sub
