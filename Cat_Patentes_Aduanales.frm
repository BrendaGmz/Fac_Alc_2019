VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Cat_Patentes_Aduanales 
   Caption         =   "Catalogo de Patentes Aduanales"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra_Vigencia 
      Caption         =   "Fechas de vigencia"
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
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   6975
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   150142977
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   150142977
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio de Vigencia"
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
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1560
      Picture         =   "Cat_Patentes_Aduanales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "M"
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5880
      Picture         =   "Cat_Patentes_Aduanales.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3000
      Picture         =   "Cat_Patentes_Aduanales.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "B"
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4440
      Picture         =   "Cat_Patentes_Aduanales.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "C"
      Top             =   5040
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Cat_Patentes_Aduanales.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "A"
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.TextBox Txt_Codigo_Patente_Aduanal 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid Grid_Patentes 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   6975
      Begin VB.Label Label2 
         Caption         =   "Clave de patente aduanal"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATENTES ADUANALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Frm_Cat_Patentes_Aduanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Consulta_Patentes(Cadena As String)
Dim Rs_Consulta_Cat_Patentes As rdoResultset   'Manejo de registro
        Grid_Patentes.Rows = 0
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Patentes_Aduanales"
    Mi_SQL = Mi_SQL & " WHERE (Codigo_Patente_Aduanal ='" & Cadena & "')"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Patentes.EOF Then
        'Pone un encabezado en el grid
        Grid_Patentes.AddItem "Clave" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Patentes
                    Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                    If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
              
        End With
        While Not Rs_Consulta_Cat_Patentes.EOF
        Grid_Patentes.AddItem Rs_Consulta_Cat_Patentes.rdoColumns("Codigo_Patente_Aduanal") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Estado")
            Grid_Patentes.FixedRows = 1
            Rs_Consulta_Cat_Patentes.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Patentes.FixedCols = 1
        Grid_Patentes.ColWidth(0) = 1300
        Grid_Patentes.ColAlignment(0) = flexAlignCenterCenter
        Grid_Patentes.ColWidth(1) = 1800
        Grid_Patentes.ColAlignment(1) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(2) = 1800
        Grid_Patentes.ColAlignment(2) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(3) = 1500
        Grid_Patentes.ColAlignment(3) = flexAlignLeftCenter
                
    Else
        MsgBox "El código no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Patentes.Close
End Sub
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Codigo_Patente_Aduanal.text <> "" Then
        Consulta_Patentes (Txt_Codigo_Patente_Aduanal.text)
        
    Else
        MsgBox "Ingrese el código de la patente aduanal", vbInformation
    End If
End Sub

Private Sub Txt_Codigo_Patente_Aduanal_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Btn_Eliminar_Click()
If Txt_Codigo_Patente_Aduanal <> "" Then
        Cambiar_Estado (Txt_Codigo_Patente_Aduanal.text)
        
    Else
        MsgBox "Ingrese el código de la patente aduanal", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Patentes As rdoResultset   'Manejo de registro
    
    Grid_Patentes.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Patentes_Aduanales WHERE (Codigo_Patente_Aduanal ='" & Cadena & "')"
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Patentes.rdoColumns("Estado") = "ACTIVO" Or Rs_Consulta_Cat_Patentes.rdoColumns("Estado") = "" Then
        Mi_SQL = "UPDATE Cat_Patentes_Aduanales SET Estado='INACTIVO' WHERE (Codigo_Patente_Aduanal ='" & Cadena & "')"
    Else
        Mi_SQL = "UPDATE Cat_Patentes_Aduanales SET Estado='ACTIVO' WHERE (Codigo_Patente_Aduanal='" & Cadena & "')"
        End If
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Patentes.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
             If Trim(Txt_Codigo_Patente_Aduanal.text) <> "" Then
                Btn_Modificar.Caption = "Actualizar"
                Btn_Salir.Caption = "Regresar"
                Btn_Modificar.Enabled = True
                Btn_Consultar.Enabled = False
                Fra_Vigencia.Enabled = True
                                    
            Else
                MsgBox "Ingrese código de la patente aduanal", vbExclamation
                
                Exit Sub
            End If
        
    ElseIf Btn_Modificar.Caption = "Actualizar" Then
            If Trim(Txt_Codigo_Patente_Aduanal.text) <> "" Then
                    Modifica_Patentes
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        End If
End Sub
Public Sub Modifica_Patentes()
Dim Rs_Modificacion_Cat_Unidades As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Patentes_Aduanales"
    Mi_SQL = Mi_SQL & " WHERE Codigo_Patente_Aduanal='" & Txt_Codigo_Patente_Aduanal.text & "'"
    Set Rs_Modificacion_Cat_Patentes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Patentes.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Patentes
            .Edit
                
                .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
                If Chc_Vigencia.Value = 1 Then
                    .rdoColumns("Fecha_Fin_Vigencia") = Null
                 Else
                .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
                  End If
            .Update
        End With
        
    Else
        MsgBox "El código no existe", vbExclamation
        Exit Sub
    End If
    Rs_Modificacion_Cat_Patentes.Close
    MsgBox "La patente ha sido modificada", vbInformation
     Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
    Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Patentes.Rows <> 0 Then
                    Grid_Patentes.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Fra_Vigencia.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Trim(Txt_Codigo_Patente_Aduanal.text) <> "" Then
                Alta_Patentes
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub

Public Sub Alta_Patentes()
Dim Rs_Alta_Cat_Patente As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    ''Valida si ya existe el codigo
    
    Mi_SQL = " SELECT Codigo_Patente_Aduanal FROM Cat_Patentes_Aduanales where Codigo_Patente_Aduanal ='" & Trim(Txt_Codigo_Patente_Aduanal.text) & "'"
    Set Rs_Consulta_Codigo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Codigo.EOF Then
        MsgBox "El Codigo que quiere dar de alta ya existe", vbInformation
        Txt_Codigo_Patente_Aduanal.SetFocus
        Exit Sub
    End If
    Rs_Consulta_Codigo.Close
    
    'Alta de Producto
    Set Rs_Alta_Cat_Patente = Conectar_Ayudante.Recordset_Agregar("Cat_Patentes_Aduanales")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Patente
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Unidades_Medida", "Clave"), "00")
            .rdoColumns("Codigo_Patente_Aduanal") = Txt_Codigo_Patente_Aduanal.text
            .rdoColumns("Clave") = Clave
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Estado") = UCase("ACTIVO")
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Patente.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Patentes.Rows = 0 Then
        Grid_Patentes.AddItem "Clave" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Patentes.AddItem UCase(Txt_Codigo_Patente_Aduanal.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Patentes.AddItem UCase(Txt_Codigo_Patente_Aduanal.text) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
           End If
    
        Grid_Patentes.FixedCols = 1
        Grid_Patentes.ColWidth(0) = 1300
        Grid_Patentes.ColAlignment(0) = flexAlignCenterCenter
        Grid_Patentes.ColWidth(1) = 1800
        Grid_Patentes.ColAlignment(1) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(2) = 1800
        Grid_Patentes.ColAlignment(2) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(3) = 1500
        Grid_Patentes.ColAlignment(3) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        Btn_Modificar.Caption = "Modificar"
        Btn_Modificar.Enabled = False
        Btn_Consultar.Enabled = True
        Btn_Eliminar.Enabled = False
        Fra_Vigencia.Enabled = False
        Chc_Vigencia = 1
        Chc_Vigencia_Click
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        If Grid_Patentes.Rows <> 0 Then
            Grid_Patentes.Rows = 0
            End If
        Consulta
    End If
End Sub

Private Sub Chc_Vigencia_Click()
 If Chc_Vigencia.Value = 1 Then
        Dtp_Fin_Vigencia.Enabled = False
    Else
        Dtp_Fin_Vigencia.Enabled = True
        Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
        Dtp_Fin_Vigencia.Value = Dtp_Inicio_Vigencia.Value
    End If
End Sub

Private Sub Form_Load()
    'Set Conexion = New Conectar
    'Conexion.ConectarBD
    Consulta
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Patentes As rdoResultset   'Manejo de registro
    
    Grid_Patentes.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Patentes_Aduanales"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Patentes.EOF Then
        'Pone un encabezado en el grid
        Grid_Patentes.AddItem "Clave" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Patentes.EOF
            Grid_Patentes.AddItem Rs_Consulta_Cat_Patentes.rdoColumns("Codigo_Patente_Aduanal") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & Rs_Consulta_Cat_Patentes.rdoColumns("Estado")
            Grid_Patentes.FixedRows = 1
            Rs_Consulta_Cat_Patentes.MoveNext
        Wend
        'Tamaño de las columnas en el grid
        Grid_Patentes.FixedCols = 1
        Grid_Patentes.ColWidth(0) = 1300
        Grid_Patentes.ColAlignment(0) = flexAlignCenterCenter
        Grid_Patentes.ColWidth(1) = 1800
        Grid_Patentes.ColAlignment(1) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(2) = 1800
        Grid_Patentes.ColAlignment(2) = flexAlignLeftCenter
        Grid_Patentes.ColWidth(3) = 1500
        Grid_Patentes.ColAlignment(3) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Patentes.Close
End Sub

Private Sub Grid_Patentes_Click()
Dim Rs_Consulta_Cat_Patentes As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Patentes.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Patentes_Aduanales"
        Mi_SQL = Mi_SQL & " WHERE Codigo_Patente_Aduanal='" & Grid_Patentes.TextMatrix(Grid_Patentes.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Patentes.EOF Then
            With Rs_Consulta_Cat_Patentes
                Txt_Codigo_Patente_Aduanal.text = .rdoColumns("Codigo_Patente_Aduanal")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
                
            End With
        End If
        Rs_Consulta_Cat_Patentes.Close
    End If
End Sub
