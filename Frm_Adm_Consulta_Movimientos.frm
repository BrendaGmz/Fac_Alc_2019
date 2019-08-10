VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Movimientos_Consulta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Pagos"
   ClientHeight    =   6570
   ClientLeft      =   2790
   ClientTop       =   2475
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   12885
   Begin MSComDlg.CommonDialog Cmd_Exportar 
      Left            =   180
      Top             =   6975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Height          =   1755
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   12750
      Begin VB.CommandButton Btn_Imprimir_Reporte 
         Caption         =   "Imprimir Reporte"
         Enabled         =   0   'False
         Height          =   150
         Left            =   2685
         TabIndex        =   15
         Tag             =   "C"
         Top             =   1710
         Visible         =   0   'False
         Width           =   360
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
         Left            =   11280
         Picture         =   "Frm_Adm_Consulta_Movimientos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "C"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CheckBox Chk_Estatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Estatus"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   1365
         Width           =   1065
      End
      Begin VB.CheckBox Chk_Ordenar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ordenar Por"
         Height          =   210
         Left            =   9075
         TabIndex        =   7
         Top             =   637
         Width           =   1380
      End
      Begin VB.CheckBox Chk_Forma_Pago 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forma Pago"
         Height          =   210
         Left            =   9075
         TabIndex        =   5
         Top             =   262
         Width           =   1350
      End
      Begin VB.CheckBox Chk_Fechas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   990
         Width           =   1065
      End
      Begin VB.CheckBox Chk_Banco 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Banco"
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   637
         Width           =   1065
      End
      Begin VB.CheckBox Chk_Proveedor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proveedor"
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   262
         Width           =   1065
      End
      Begin VB.ComboBox Cmb_Con_Estatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Movimientos.frx":358C
         Left            =   1275
         List            =   "Frm_Adm_Consulta_Movimientos.frx":3596
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   2040
      End
      Begin VB.TextBox Txt_Total 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   4065
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1365
         Width           =   4125
      End
      Begin VB.ComboBox Cmb_Con_Ordenar 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Movimientos.frx":35AF
         Left            =   10515
         List            =   "Frm_Adm_Consulta_Movimientos.frx":35BF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   585
         Width           =   2145
      End
      Begin VB.ComboBox Cmb_Con_Proveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Top             =   210
         Width           =   7725
      End
      Begin VB.ComboBox Cmb_Con_Forma_Pago 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Movimientos.frx":35E8
         Left            =   10515
         List            =   "Frm_Adm_Consulta_Movimientos.frx":35FB
         TabIndex        =   3
         Top             =   210
         Width           =   2145
      End
      Begin VB.ComboBox Cmb_Con_Banco 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Adm_Consulta_Movimientos.frx":3640
         Left            =   1275
         List            =   "Frm_Adm_Consulta_Movimientos.frx":3642
         TabIndex        =   6
         Top             =   585
         Width           =   7740
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_1 
         Height          =   315
         Left            =   1275
         TabIndex        =   10
         Top             =   945
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58785795
         CurrentDate     =   38038
      End
      Begin MSComCtl2.DTPicker DTP_Fecha_2 
         Height          =   315
         Left            =   6810
         TabIndex        =   11
         Top             =   945
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58785795
         CurrentDate     =   38038
      End
      Begin VB.CommandButton Btn_Reimprimir 
         Caption         =   "Reimprimir Cheque"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2190
         TabIndex        =   14
         Tag             =   "C"
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Lbl_Total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3345
         TabIndex        =   17
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   4740
         TabIndex        =   16
         Top             =   1005
         Width           =   135
      End
   End
   Begin VB.Frame Fra_detalles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   90
      TabIndex        =   20
      Top             =   1740
      Width           =   12750
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
         Left            =   11310
         Picture         =   "Frm_Adm_Consulta_Movimientos.frx":3644
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "A"
         Top             =   4050
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
         Left            =   5850
         Picture         =   "Frm_Adm_Consulta_Movimientos.frx":6D43
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "A"
         Top             =   4050
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
         Left            =   135
         Picture         =   "Frm_Adm_Consulta_Movimientos.frx":957D
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "A"
         Top             =   4050
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Pagos 
         Height          =   3750
         Left            =   75
         TabIndex        =   21
         Top             =   240
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   6615
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "Frm_Adm_Movimientos_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Formato As String    'Almacena el nombre del formato de impresion del cheque

'**************************************************************************************
    'Nombre Funcion : Cancela movimientos
    'Descripcion    : Realiza cancelacion de los movimientos que se vayan seleccionando
    'Creo           : Joel Romero
    'Fecha Creo     :
    'Fecha Modificacion  : 10-Agosto-2006
    'Usuario Modifico    : Jorge Razo
    'Motivo Modificacion : Estandarizacion de codigo
    'Parametros     :
'**************************************************************************************
Public Sub Cancela_Movimientos()
Dim Rs_Actualiza_Factura As rdoResultset                    'Manejador de datos para actualizar la factura
Dim Rs_Actualiza_Movimiento As rdoResultset                 'Manejador de actualizacion de movimientos
Dim Rs_Consulta_Anticipos_Proveedores As rdoResultset       'Manejador de editar anticipo
Dim Importe As Double                                       'Almacena el importe del movimiento eliminado para activar la factura
Dim Concepto As String                                      'almacena el concepto de cancelacion del movimiento
Dim Banco As String                                         'Almacena el banco del movimiento a cancelar
Dim Fecha As Date                                           'Almacena la fecha del movimiento a cancelar
Dim Cheque As String

On Error GoTo Handler
    If Grid_Consulta_Pagos.Rows > 1 Then
        If Grid_Consulta_Pagos.RowSel > 0 And Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 9) = "A" Then
            If MsgBox("Esta Operación eliminará todas las afectaciones del movimiento" & Chr(13) & Chr(13) & "¿Seguro de Cancelar el movimiento?", 3) = 6 Then
                Conexion_Base.BeginTrans
                'Modifica el moviemiento
                Mi_SQL = "SELECT * FROM Adm_Movimientos"
                Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 0)) & "' "
                Set Rs_Actualiza_Movimiento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                With Rs_Actualiza_Movimiento
                    .Edit
                        If Not IsNull(.rdoColumns("Referencia")) Then Cheque = .rdoColumns("Referencia")
                        .rdoColumns("Estatus") = "C"
                        .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                        .rdoColumns("Fecha_Modifico") = Now()
                         Concepto = .rdoColumns("Concepto")
                    .Update
                End With
                'Set Rs_Actualiza_Movimiento = Nothing
                If Rs_Actualiza_Movimiento.rdoColumns("Concepto") = "ANTICIPO" Then
                    'Cambia el estatus del anticipo
                    Mi_SQL = "SELECT Estatus, No_Movimiento, Proveedor_ID, Estatus,Usuario_Modifico,Fecha_Modifico FROM Adm_Proveedores_Anticipos"
                    Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 0) & "'"
                    Mi_SQL = Mi_SQL & " AND Proveedor_ID = '" & Format(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 8), "00000") & "'"
                    Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    With Rs_Consulta_Anticipos_Proveedores
                        If Not Rs_Consulta_Anticipos_Proveedores.EOF Then
                            If .rdoColumns("Estatus") = "ACTIVO" Then
                                .Edit
                                    .rdoColumns("Estatus") = "CANCELADO"
                                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                    .rdoColumns("Fecha_Modifico") = Now()
                                .Update
                            Else
                                .Edit
                                    .rdoColumns("Estatus") = "ACTIVO"
                                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                    .rdoColumns("Fecha_Modifico") = Now()
                                .Update
                            End If
                        End If
                    End With
                    Rs_Consulta_Anticipos_Proveedores.Close
                End If
                Rs_Actualiza_Movimiento.Close
                
                'Actualiza las facturas
                Mi_SQL = "SELECT No_Factura,Abono, Saldo, Pagada, Fecha_Recepcion, Usuario_Modifico, Fecha_Modifico, Moneda, Tipo_Cambio  "
                Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas "
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Trim(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 7)) & "'"
                Mi_SQL = Mi_SQL & " AND Proveedor_ID = '" & Trim(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 8)) & "'"
                Set Rs_Actualiza_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                With Rs_Actualiza_Factura
                    If Not .EOF Then
                        If .rdoColumns("Moneda") = "Dolares" Then
                            Importe = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4), ",")) / .rdoColumns("Tipo_Cambio")
                        Else
                            Importe = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4), ","))
                        End If
                        .Edit
                            .rdoColumns("Pagada") = "N"
                            .rdoColumns("Saldo") = .rdoColumns("Saldo") + Importe
                            .rdoColumns("Abono") = .rdoColumns("Abono") - Importe
                            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                            .rdoColumns("Fecha_Modifico") = Now()
                        .Update
                        'Call Actualiza_Estatus_Documento("Ope_Ordenes_Carga WHERE No_Carga='" & .rdoColumns("No_Carga") & "' ", "No_Carga,Entregada", "Entregada", "No")
                    End If
                End With
                Rs_Actualiza_Factura.Close
                Conexion_Base.CommitTrans
                Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 9) = "C"
                Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4) = "0"
                MsgBox "Movimiento Cancelado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            Else
                MsgBox "Operación Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
            End If
        End If
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'**************************************************************************************
    'Nombre Funcion : Cancela notas de credito
    'Descripcion    : Realiza cancelacion de las notas de credito
    'Creo           : Joel Romero
    'Fecha Creo     :
    'Fecha Modificacion  : 10-Agosto-2006
    'Usuario Modifico    : Jorge Razo
    'Motivo Modificacion : Estandarizacion de codigo
    'Parametros     :
'**************************************************************************************
Public Sub Cancela_Notas_Credito()
Dim rs_Elimina_Nota_Credito As rdoResultset   'Manejador de datos para eliminar nota de credito
Dim Rs_Actualiza_Factura As rdoResultset      'Manejador de datos para actualizar la factura
Dim Importe As Double                         'Almacena el monto de la nota de credito y regresar el saldo a la factura


On Error GoTo Handler
If Grid_Consulta_Pagos.RowSel > 0 Then
    If MsgBox("Esta Operación eliminará todas las afectaciones de la Nota de Crédito y la Nota de Crédito" & Chr(13) & Chr(13) & "¿Seguro de Eliminar la Nota de Crédito?", vbYesNo + vbQuestion) = 6 Then
        Conexion_Base.BeginTrans
        'Elimna la nota de credito
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Notas_Credito"
        Mi_SQL = Mi_SQL & " WHERE No_Nota_Credito = '" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 0) & "' AND Proveedor_ID ='" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 5) & "'"
        Set rs_Elimina_Nota_Credito = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not rs_Elimina_Nota_Credito.EOF Then
            rs_Elimina_Nota_Credito.Edit
                rs_Elimina_Nota_Credito("Importe") = 0
            rs_Elimina_Nota_Credito.Update
        End If
        rs_Elimina_Nota_Credito.Close
        'Actualiza las facturas de proveedor con el saldo correcto al cancelar la nota de credito
        Mi_SQL = "SELECT No_Factura, Abono, Saldo, Pagada, Fecha_Pago, Usuario_Modifico, Fecha_Modifico, Moneda, Tipo_Cambio  "
        Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Facturas "
        Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4) & "'"
        Mi_SQL = Mi_SQL & " AND Proveedor_ID = '" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 5) & "'"
        Set Rs_Actualiza_Factura = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        With Rs_Actualiza_Factura
            If Not .EOF Then
                Importe = Val(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 3), ","))
                .Edit
                    .rdoColumns("Pagada") = "N"
                    .rdoColumns("Saldo") = .rdoColumns("Saldo") + Importe
                    .rdoColumns("Abono") = .rdoColumns("Abono") - Importe
                    .rdoColumns("Usuario_Modifico") = Usuario
                    .rdoColumns("Fecha_Modifico") = Now
                .Update
            End If
        End With
        Rs_Actualiza_Factura.Close
        
        'Actualiza el movmiento de la nota de credito para cancelar el movimiento administrativo
        Mi_SQL = "UPDATE Adm_Movimientos SET Cantidad =0, Estatus ='C', Referencia ='Cancelado', Concepto = 'CANCELADO  ' + Concepto WHERE Referencia ='" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4) & "' and Tipo ='N' and Proveedor_Cliente ='" & Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 5) & "'"
        Conexion_Base.Execute Mi_SQL
        
        Conexion_Base.CommitTrans
        'remueve del otro grid la factura
        If Grid_Consulta_Pagos.Rows = 2 Then
            Grid_Consulta_Pagos.Rows = 0
        Else
            Grid_Consulta_Pagos.RemoveItem (Grid_Consulta_Pagos.RowSel)
        End If
        MsgBox "Nota de Crédito Eliminada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Else
        MsgBox "Operación Cancelada", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'**************************************************************************************
    'Nombre Funcion : Consulta movimientos
    'Descripcion    : Realiza la consulta de los movimientos segun los criterios de busqueda
    'Creo           : Joel Romero
    'Fecha Creo     :
    'Fecha Modificacion  : 10-Agosto-2006
    'Usuario Modifico    : Jorge Razo
    'Motivo Modificacion : Estandarizacion de codigo
    'Parametros     :
'**************************************************************************************
Public Sub Consulta_Movimientos(Optional ByVal No_Factura As String = "")
Dim Rs_Consulta_Movimientos As rdoResultset  'Manejador para consulta general
Dim Rs_Cat_Proveedores As rdoResultset
Dim Texto As String                          'Almacena la cadena de No_Cheque ó Transferencia dependiendo la forma de pago
Dim Cadena As String                         'Almacena el valor de la consulta de los campos Referencia o No_cheque dependiendo la forma de pago
    
'Consulta los vales pendientes de facturar
On Error GoTo Handler
    Mi_SQL = "SELECT No_Movimiento, Referencia, Fecha, Cantidad, Concepto, No_Factura, Proveedor_ID, Estatus, Banco_ID, Forma_Pago, Banco"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Tipo='E'"
    If Chk_Proveedor.Value = 1 And Cmb_Con_Proveedor.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & Format(Cmb_Con_Proveedor.ItemData(Cmb_Con_Proveedor.ListIndex), "00000") & "'"
    End If
    If Chk_Banco.Value = 1 And Cmb_Con_Banco.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Banco_ID='" & Format(Cmb_Con_Banco.ItemData(Cmb_Con_Banco.ListIndex), "00000") & "'"
    End If
    If Chk_Forma_Pago.Value = 1 And Cmb_Con_Forma_Pago.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Forma_Pago='" & Cmb_Con_Forma_Pago.Text & "' "
    End If
    If Chk_Estatus.Value = 1 Then
        If Cmb_Con_Estatus = "Activos" Then Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Estatus= 'A' "
        If Cmb_Con_Estatus = "Cancelados" Then Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Estatus= 'C' "
    End If
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha >= " & Par_Fecha & Format(DTP_Fecha_1.Value, "MM/dd/yyyy") & Par_Fecha & ""
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha <= " & Par_Fecha & Format(DTP_Fecha_2.Value, "MM/dd/yyyy") & Par_Fecha & " "
    End If
    If No_Factura <> "" Then Mi_SQL = Mi_SQL & " AND Adm_Movimientos.No_Factura='" & No_Factura & "' "
    If Chk_Ordenar.Value = 1 Then
        If Cmb_Con_Ordenar.Text = "No. Cheque" Then Mi_SQL = Mi_SQL & " ORDER BY Referencia "
        If Cmb_Con_Ordenar.Text = "Proveedor" Then Mi_SQL = Mi_SQL & " ORDER BY Proveedor_ID "
        If Cmb_Con_Ordenar.Text = "Fecha" Then Mi_SQL = Mi_SQL & " ORDER BY Adm_Movimientos.Fecha"
        If Cmb_Con_Ordenar.Text = "Banco" Then Mi_SQL = Mi_SQL & " ORDER BY Banco "
    End If
    Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Realiza la consulta de los movimientos en base a los parametros especificados en los criterios
    With Rs_Consulta_Movimientos
        Grid_Consulta_Pagos.Rows = 0
        Grid_Consulta_Pagos.Cols = 12
        Grid_Consulta_Pagos.AddItem "No. Movimiento" & Chr(9) & "Referencia" & Chr(9) _
            & "Fecha" & Chr(9) & "Proveedor" _
            & Chr(9) & "Cantidad" & Chr(9) & "Concepto" & Chr(9) & "Leyenda" & Chr(9) & "No Factura" _
            & Chr(9) & "Proveedor_ID" & Chr(9) & "Estatus" & Chr(9) & "Banco" & Chr(9) & "Banco_ID"
        Txt_Total.Text = 0
        While Not .EOF
            Mi_SQL = "SELECT * FROM Cat_Proveedores WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
            Set Rs_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Cat_Proveedores.EOF Then
                Grid_Consulta_Pagos.AddItem .rdoColumns("No_Movimiento") & Chr(9) & .rdoColumns("Referencia") & Chr(9) _
                & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Rs_Cat_Proveedores.rdoColumns("Nombre") & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,###.00") _
                & Chr(9) & .rdoColumns("Concepto") & Chr(9) & "" & Chr(9) & .rdoColumns("No_Factura") & Chr(9) & _
                Rs_Cat_Proveedores.rdoColumns("Proveedor_ID") & Chr(9) & .rdoColumns("Estatus") & Chr(9) & .rdoColumns("Banco") & Chr(9) & .rdoColumns("Banco_ID")
                Txt_Total.Text = Val(Txt_Total.Text) + .rdoColumns("Cantidad")
            Else
                Grid_Consulta_Pagos.AddItem .rdoColumns("No_Movimiento") & Chr(9) & .rdoColumns("Referencia") _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & "" & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,###.00") _
                    & Chr(9) & .rdoColumns("Concepto") & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & .rdoColumns("Estatus") & Chr(9) & .rdoColumns("Banco") & Chr(9) & .rdoColumns("Banco_ID")
                Txt_Total.Text = Val(Txt_Total.Text) + .rdoColumns("Cantidad")
            End If
            Rs_Cat_Proveedores.Close
            .MoveNext
            Grid_Consulta_Pagos.FixedRows = 1
            Grid_Consulta_Pagos.FixedCols = 1
        Wend
        Txt_Total.Text = Format(Txt_Total.Text, "###,###,###.00")
        Grid_Consulta_Pagos.ColWidth(0) = 1200                      'Columna No Movimiento
        Grid_Consulta_Pagos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Consulta_Pagos.ColWidth(1) = 1000                      'Columna Referencia ó Cheque
        Grid_Consulta_Pagos.ColAlignment(1) = 3
        Grid_Consulta_Pagos.ColWidth(2) = 950                       'Columna Fecha
        Grid_Consulta_Pagos.ColAlignment(2) = flexAlignCenterCenter
        Grid_Consulta_Pagos.ColWidth(3) = 2400                      'Columna Proveedor
        Grid_Consulta_Pagos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Consulta_Pagos.ColWidth(4) = 1050                      'Columna Cantidad
        Grid_Consulta_Pagos.ColWidth(5) = 2000                      'Columna Concepto
        Grid_Consulta_Pagos.ColWidth(6) = 0                         'Columna Leyenda
        Grid_Consulta_Pagos.ColWidth(7) = 1000                      'Columna No Factura
        Grid_Consulta_Pagos.ColAlignment(7) = flexAlignCenterCenter
        Grid_Consulta_Pagos.ColAlignment(8) = flexAlignCenterCenter
        Grid_Consulta_Pagos.ColWidth(8) = 0                         'Columna Proveedor_ID
        Grid_Consulta_Pagos.ColWidth(9) = 1000                      'Columna Estatus
        Grid_Consulta_Pagos.ColWidth(10) = 1600                     'Columna Banco
        Grid_Consulta_Pagos.ColWidth(11) = 0                        'Columna Banco_ID
    End With
    Rs_Consulta_Movimientos.Close
    If Grid_Consulta_Pagos.Rows > 1 Then
        Grid_Consulta_Pagos.FixedRows = 1
        Btn_Excel.Enabled = True
        Btn_Imprimir_Reporte.Enabled = True
        Btn_Reimprimir.Enabled = True
    Else
        Btn_Excel.Enabled = False
        Btn_Imprimir_Reporte.Enabled = False
        Btn_Reimprimir.Enabled = False
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'**************************************************************************************
    'Nombre Funcion : Consulta Notas de credito
    'Descripcion    : Realiza la consulta de las notas de credito segun los criterios de busqueda
    'Creo           : Joel Romero
    'Fecha Creo     :
    'Fecha Modificacion  : 10-Agosto-2006
    'Usuario Modifico    : Jorge Razo
    'Motivo Modificacion : Estandarizacion de codigo
    'Parametros     :
'**************************************************************************************
Public Sub Consulta_Notas_Credito()
Dim RS_Consulta_Notas_Credito As rdoResultset    'Manejador de consulta de notas de credito

    'Consulta los vales pendientes de facturar
    On Error GoTo Handler
    Mi_SQL = "SELECT No_Nota_Credito, Fecha, Importe, Nombre, No_Factura, Cat_Proveedores.Proveedor_ID "
    Mi_SQL = Mi_SQL & " FROM Adm_Proveedores_Notas_Credito, Cat_Proveedores "
    Mi_SQL = Mi_SQL & " WHERE Adm_Proveedores_Notas_Credito.Proveedor_ID = Cat_Proveedores.Proveedor_ID "
    Mi_SQL = Mi_SQL & " AND Fecha >= '" & Format(DTP_Fecha_1.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Fecha <= '" & Format(DTP_Fecha_2.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY No_Nota_Credito "
    Set RS_Consulta_Notas_Credito = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'realiza la consulta de las notas de credito en la base de datos segun los criterios
    With RS_Consulta_Notas_Credito
        Grid_Consulta_Pagos.Rows = 0
        Grid_Consulta_Pagos.Cols = 6
        Grid_Consulta_Pagos.AddItem "No. Nota" & Chr(9) & "Fecha" & Chr(9) & "Proveedor" _
        & Chr(9) & "Cantidad" & Chr(9) & "No Factura" & Chr(9) & "Proveedor_ID"
        While Not .EOF
            Grid_Consulta_Pagos.AddItem .rdoColumns("No_Nota_Credito") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & _
            .rdoColumns("Nombre") & Chr(9) & Format(.rdoColumns("Importe"), "#,###,###.00") & Chr(9) & .rdoColumns("No_Factura") _
            & Chr(9) & .rdoColumns("Proveedor_ID")
            .MoveNext
            Grid_Consulta_Pagos.FixedRows = 1
            Grid_Consulta_Pagos.FixedCols = 1
        Wend
        Grid_Consulta_Pagos.ColWidth(0) = 1300
        Grid_Consulta_Pagos.ColAlignment(0) = 3
        Grid_Consulta_Pagos.ColWidth(1) = 1300
        Grid_Consulta_Pagos.ColWidth(2) = 3000
        Grid_Consulta_Pagos.ColWidth(3) = 1400
        Grid_Consulta_Pagos.ColWidth(4) = 1200
        Grid_Consulta_Pagos.ColWidth(5) = 0
    End With
    RS_Consulta_Notas_Credito.Close
Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Cancelar_Click()
    'Realiza la validacion del usuario para que pueda cancelar los movimientos
        ''If UCase(Rol_ID) = "00001" Then 'UCase("Administrador") Or UCase(Rol) = UCase("Supervisor") Then
            If Cmb_Con_Forma_Pago.Text = "Notas Credito" Then
                Cancela_Notas_Credito
            Else
                Cancela_Movimientos
            End If
        ''Else
            ''MsgBox "No tiene privilegios para cancelar movimientos", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        ''End If
End Sub

Public Sub Btn_Consultar_Click()
    'hace el llamado a las consultas de movimientos
    If Cmb_Con_Forma_Pago.Text = "Notas Credito" Then
        Consulta_Notas_Credito
    Else
        Consulta_Movimientos
    End If
End Sub

Private Sub Btn_Excel_Click()
Dim RutaArchivo As String
    'Realiza la exportacion de la consulta
    ' Set CancelError is True
    Cmd_Exportar.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    Cmd_Exportar.Flags = cdlOFNHideReadOnly
    ' Set filters
    Cmd_Exportar.Filter = "Archivos de Excel |*.Xls|"
    ' Specify default filter
    Cmd_Exportar.FilterIndex = 2
    ' Display the Open dialog box
    Cmd_Exportar.ShowSave
    ' Display name of selected file
    RutaArchivo = Cmd_Exportar.FileName
    Open RutaArchivo For Output As #1
        For I = 0 To Grid_Consulta_Pagos.Rows - 1
            For J = 0 To Grid_Consulta_Pagos.Cols - 1
                CABECERA = CABECERA & Grid_Consulta_Pagos.TextMatrix(I, J) & Chr(9)
            Next J
            Print #1, CABECERA
            CABECERA = ""
        Next I
    Close #1
    MsgBox "Reporte exportado", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub Btn_Imprimir_Reporte_Click()
    'Imprime los datos que hay en pantalla
    If Grid_Consulta_Pagos.Rows > 1 Then
        Printer.FontSize = 10
        Printer.Font = "COURIER NEW"
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 1), 6);
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 2), 6);
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 3), 31);
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 4), 10);
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 7), 11);
        Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(0, 11), 15)
        For I = 1 To Grid_Consulta_Pagos.Rows - 1
            Printer.Print Alinea_Derecha(Grid_Consulta_Pagos.TextMatrix(I, 1), 5); Spc(1);
            Printer.Print Format(Grid_Consulta_Pagos.TextMatrix(I, 2), "dd/MMM/yy"); Spc(1);
            Printer.Print Mid(Grid_Consulta_Pagos.TextMatrix(I, 3), 1, 25); Spc(26 - Len(Mid(Grid_Consulta_Pagos.TextMatrix(I, 3), 1, 25)));
            Printer.Print Alinea_Derecha(Format(Grid_Consulta_Pagos.TextMatrix(I, 4), "###,###,##0.00"), 15); Spc(3);
            Printer.Print Mid(Grid_Consulta_Pagos.TextMatrix(I, 7), 1, 10); Spc(11 - Len(Mid(Grid_Consulta_Pagos.TextMatrix(I, 7), 1, 10)));
            Printer.Print Mid(Grid_Consulta_Pagos.TextMatrix(I, 11), 1, 15)
        Next I
        Printer.EndDoc
        MsgBox "Reporte enviado a impresión", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
End Sub

'**************************************************************************************
    'Nombre Funcion : reimprimir cheque
    'Descripcion    : Realiza la reimpresion de los cheques
    'Creo           : Joel Romero
    'Fecha Creo     :
    'Fecha Modificacion  : 10-Agosto-2006
    'Usuario Modifico    : Jorge Razo
    'Motivo Modificacion : Estandarizacion de codigo
    'Parametros     :
'**************************************************************************************
Private Sub Btn_Reimprimir_Click()
Dim Rs_Consulta_Formato  As rdoResultset   'Manejador de consulta para datos del formato de impresion
Dim Rs_Consulta_Generales  As rdoResultset 'Manejador de consulta para datos generales del formato de impresion
Dim Rs_Consulta_Detalles As rdoResultset   'Manejador de consulta para detalles del formato de impresion
Dim Longitud As Integer                    'Almacena la longitud de cada campo a imprimir
Dim Inicio As Integer                      'Indica donde empieza a imprimir los detalles
Dim Salto As Double                        'Indica el espacio entre partida y partida

    If Grid_Consulta_Pagos.RowSel > 0 Then
        'Consulta el encabezado del formato
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cfg_Formatos"
        Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
        Set Rs_Consulta_Formato = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Consulta los datos tipo generales del formato
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
        Mi_SQL = Mi_SQL & " AND Tipo = 'General'"
        Set Rs_Consulta_Generales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Consulta los datos tipo Detalle del formato
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cfg_Formatos_Detalles"
        Mi_SQL = Mi_SQL & " WHERE  Nombre = '" & Formato & "'"
        Mi_SQL = Mi_SQL & " AND Tipo = 'Detalle'"
        Set Rs_Consulta_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Formato.EOF Then
            With Rs_Consulta_Formato
                'Comienza la impresion de la factura
                Printer.ScaleMode = vbCentimeters
                'Configura la fuente de la factura
                Printer.FontSize = .rdoColumns("Tamaño_Generales")
                Printer.Font = .rdoColumns("Letra_Generales")
                If .rdoColumns("Estilo_Generales") = "Negrita" Then
                    Printer.FontBold = True
                Else
                    Printer.FontBold = False
                End If
                Salto = .rdoColumns("Separacion_Detalles")
            End With
            'Inicia la impresión
            'Imprime la fecha de la factura y la ciudad
            With Rs_Consulta_Generales
                While Not .EOF
                    Printer.CurrentX = .rdoColumns("X")
                    Printer.CurrentY = .rdoColumns("Y")
                    Longitud = .rdoColumns("Longitud")
                    If .rdoColumns("Campo") = "Lugar" Then Printer.Print "IRAPUATO, GTO."
                    If .rdoColumns("Campo") = "Fecha" Then Printer.Print Format(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 2), "dd-MMM-yyyy")
                    If .rdoColumns("Campo") = "Nombre" Then Printer.Print Mid(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 3), 1, Longitud)
                    If .rdoColumns("Campo") = "Cantidad" Then Printer.Print Format(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4), "###,###,###.00")
                    If .rdoColumns("Campo") = "Cantidad_Letra" Then Printer.Print Conectar_Ayudante.Convierte_Cantidad_Letras(Conectar_Ayudante.Quitar_Caracter(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 4), ","))
                    If .rdoColumns("Campo") = "Concepto" Then Printer.Print Mid(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 5), 1, Longitud)
                    If .rdoColumns("Campo") = "No_Cheque" Then Printer.Print "No. Cheque " & Mid(Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 10), 1, Longitud)
                    If .rdoColumns("Campo") = "Leyenda" Then
                        If Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 6) = "SI" Then Printer.Print "PARA ABONO A CUENTA DEL BENEFICIARIO"
                    End If
                    .MoveNext
                Wend
            End With
    '        'Imprime los detalles de la poliza
    '        With MiFormato
    '            Printer.FontSize = .rdoColumns("Tamaño_Detalles")
    '            Printer.Font = .rdoColumns("Letra_Detalles")
    '            If .rdoColumns("Estilo_Detalles") = "Negrita" Then
    '                Printer.FontBold = True
    '            Else
    '                Printer.FontBold = False
    '            End If
    '        End With
    '        Cont_Renglon = 0
    '        While Not MisDetallesPoliza.EOF
    '            Cont_Renglon = Cont_Renglon + Salto
    '            While Not MisDetalles.EOF
    '                Printer.CurrentX = MisDetalles.rdoColumns("X")
    '                Printer.CurrentY = MisDetalles.rdoColumns("Y") + Cont_Renglon
    '                Longitud = MisDetalles.rdoColumns("Longitud")
    '                If MisDetalles.rdoColumns("Campo") = "Cuenta" Then Printer.Print Mid(MisDetallesPoliza.rdoColumns("Cuenta"), 1, Longitud)
    '                If MisDetalles.rdoColumns("Campo") = "Debe" And MisDetallesPoliza.rdoColumns("Debe") > 0 Then
    '                    Printer.Print Conectar_ayudante.Alinea_Derecha(Format(MisDetallesPoliza.rdoColumns("Debe"), "###,###,###.00"), 13)
    '                End If
    '                If MisDetalles.rdoColumns("Campo") = "Haber" And MisDetallesPoliza.rdoColumns("Haber") > 0 Then
    '                    Printer.Print Conectar_ayudante.Alinea_Derecha(Format(MisDetallesPoliza.rdoColumns("Haber"), "###,###,###.00"), 13)
    '                End If
    '                MisDetalles.MoveNext
    '            Wend
    '            MisDetalles.MoveFirst
    '            MisDetallesPoliza.MoveNext
    '        Wend
            Printer.EndDoc
            MsgBox "Cheque enviado a impresión", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        Else
            MsgBox "No existe formato predefinido" & Chr(13) & "Seleccione el banco", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
        End If
    Else
        MsgBox "Seleccione el cheque a imprimir", vbInformation, UCase(MDIFrm_Apl_Principal.Caption)
    End If
End Sub

Private Sub Cbo_Forma_Pago_Click()
Cmb_Con_Banco.Visible = True
'Label6.Visible = True
If Cbo_Forma_Pago.Text = "Cheque" Then
    Btn_Reimprimir.Visible = True
Else
    Btn_Reimprimir.Visible = False
    If Cbo_Forma_Pago.Text = "Notas Credito" Then
        Cmb_Con_Banco.Visible = False
'        Label6.Visible = False
    End If
End If
End Sub

Private Sub Chk_Banco_Click()
    Cmb_Con_Banco.Enabled = Chk_Banco.Value
End Sub

Private Sub Chk_Estatus_Click()
    Cmb_Con_Estatus.Enabled = Chk_Estatus.Value
End Sub

Private Sub Chk_Fechas_Click()
    DTP_Fecha_1.Enabled = Chk_Fechas.Value
    DTP_Fecha_2.Enabled = Chk_Fechas.Value
End Sub

Private Sub Chk_Forma_Pago_Click()
    Cmb_Con_Forma_Pago.Enabled = Chk_Forma_Pago.Value
End Sub

Private Sub Chk_Orden_Click()
    Cmb_Con_Ordenar.Enabled = Chk_Orden.Value
End Sub

Private Sub Chk_Proveedor_Click()
    Cmb_Con_Proveedor.Enabled = Chk_Proveedor.Value
End Sub

Private Sub Cmb_Con_Banco_Click()
''Dim Rs_Consulta_Bancos As rdoResultset 'Manejador de consulta de datos del banco
''    Formato = ""
''    'Consulta el formato de impresion del cheque
''    Mi_SQL = "SELECT Formato, Banco_ID"
''    Mi_SQL = Mi_SQL & " FROM Cat_Bancos"
''    Mi_SQL = Mi_SQL & " WHERE Banco_ID = '" & Mid(Cmb_Con_Banco.Text, 1, 5) & "'"
''    Set Rs_Consulta_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''    With Rs_Consulta_Bancos
''        If Not .EOF Then
''            Formato = .rdoColumns("Formato")
''        End If
''    End With
''    Rs_Consulta_Bancos.Close
End Sub

Private Sub Cmb_Con_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Con_Banco, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Con_Proveedor_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    'Consulta del catalogo de proveedores
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID, Nombre", "Cat_Proveedores", Cmb_Con_Proveedor, 1, "Nombre")
    End If
End Sub

Private Sub Btn_Salir_Click()
    Unload Frm_Adm_Movimientos_Consulta
End Sub

Private Sub Form_Load()
    Me.Height = 7110
    Me.Width = 12975
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID, Nombre", "Cat_Bancos", Cmb_Con_Banco, 1, "Nombre")
    Call Cmb_Con_Proveedor_KeyPress(13)
    DTP_Fecha_1.Value = Now
    DTP_Fecha_2.Value = Now
    Cmb_Con_Ordenar.ListIndex = 0
End Sub



Private Sub Grid_Consulta_Pagos_Click()
    'Valida que el cheque no este cancelado
    'si ya esta cancelado nos bloquea el boton de cancelar movimiento para no aplicar
    'este movimiento dos veces
    Btn_Cancelar.Enabled = False
    If Grid_Consulta_Pagos.Rows > 1 Then
        If Grid_Consulta_Pagos.TextMatrix(Grid_Consulta_Pagos.RowSel, 9) = "A" Then
            Btn_Cancelar.Enabled = True
        End If
    End If
End Sub

