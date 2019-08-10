VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Pedimentos_Operados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de número de pedimientos operados"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1560
      Picture         =   "Cat_Pedimientos_Operados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "M"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6120
      Picture         =   "Cat_Pedimientos_Operados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3120
      Picture         =   "Cat_Pedimientos_Operados.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "B"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4680
      Picture         =   "Cat_Pedimientos_Operados.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "C"
      Top             =   4680
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Cat_Pedimientos_Operados.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "A"
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1350
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
      TabIndex        =   15
      Top             =   2520
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid Grid_Pedimentos 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         BackColorBkg    =   16777215
      End
   End
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
      Height          =   1815
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   3495
      Begin VB.CheckBox Chc_Vigencia 
         Caption         =   "Sin Definir"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp_Inicio_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   42947
      End
      Begin MSComCtl2.DTPicker Dtp_Fin_Vigencia 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   42947
      End
      Begin VB.Label Label2 
         Caption         =   "Fin de Vigencia"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   14
         Top             =   840
         Width           =   1935
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
         TabIndex        =   13
         Top             =   360
         Width           =   2175
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
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin VB.TextBox Txt_Clave 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_Ejercicio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Text            =   "Cmb_Ejercicio"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_Patente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "Cmb_Patente"
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_Aduanas 
         Height          =   315
         ItemData        =   "Cat_Pedimientos_Operados.frx":0552
         Left            =   1920
         List            =   "Cat_Pedimientos_Operados.frx":0554
         TabIndex        =   7
         Text            =   "Cmb_Aduanas"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cantidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Clave"
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
         Index           =   6
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Ejercicio"
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
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
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
         Index           =   3
         Left            =   960
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Aduana"
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Patente"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Label Lbl_Almacenes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PEDIMENTOS OPERADOS"
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
      TabIndex        =   6
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Frm_Cat_Pedimentos_Operados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Consultar_Click()
Btn_Salir.Caption = "Regresar"
    If Txt_Clave.text <> "" Then
        Consulta_Pedimento (Txt_Clave.text)
        
    Else
        MsgBox "Ingrese la clave", vbInformation
    End If
End Sub
Public Sub Consulta_Pedimento(Cadena As String)
Dim Rs_Consulta_Cat_Pedimentos As rdoResultset   'Manejo de registro
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = True
    Cmb_Patente.Enabled = True
    Cmb_Ejercicio.Enabled = True
    Txt_Cantidad.Enabled = True
    Fra_Vigencia.Enabled = True
    Grid_Pedimentos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Numero_Pedimentos_Operados"
    Mi_SQL = Mi_SQL & " WHERE (Clave =" & Cadena & ")"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Pedimentos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Pedimentos.EOF Then
        'Pone un encabezado en el grid
        Grid_Pedimentos.AddItem "Clave" & Chr(9) & "Aduana" & Chr(9) & "Patente" & Chr(9) & "Ejercicio" & Chr(9) & "Cantidad" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Asignar valores
        With Rs_Consulta_Cat_Pedimentos
               For i = 0 To Cmb_Aduanas.ListCount - 1
                     Cmb_Aduanas.ListIndex = i
                        If Cmb_Aduanas.text = .rdoColumns("Aduana_ID") Then
                            Exit For
                        End If
                Next
                 For i = 0 To Cmb_Patente.ListCount - 1
                     Cmb_Patente.ListIndex = i
                        If Cmb_Patente.text = .rdoColumns("Patente") Then
                            Exit For
                        End If
                Next
                For i = 0 To Cmb_Ejercicio.ListCount - 1
                     Cmb_Ejercicio.ListIndex = i
                        If Cmb_Ejercicio.text = .rdoColumns("Ejercicio") Then
                            Exit For
                        End If
                Next
                Txt_Cantidad.text = .rdoColumns("Cantidad")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
                If IsNull(.rdoColumns("Fecha_Fin_Vigencia")) Then
                    Chc_Vigencia.Value = 1
                    Chc_Vigencia_Click
                    Else
                    Dtp_Fin_Vigencia.Value = .rdoColumns("Fecha_Fin_Vigencia")
                    Chc_Vigencia.Value = 0
                    Chc_Vigencia_Click
                    End If
       
        'Llenado del grid
        While Not Rs_Consulta_Cat_Pedimentos.EOF
            Grid_Pedimentos.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Aduana_ID") & Chr(9) & .rdoColumns("Patente") & Chr(9) & .rdoColumns("Ejercicio") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
            Grid_Pedimentos.FixedRows = 1
            Rs_Consulta_Cat_Pedimentos.MoveNext
        Wend
         End With
        'Tamaño de las columnas en el grid
        Grid_Pedimentos.FixedCols = 1
        Grid_Pedimentos.ColWidth(0) = 800
        Grid_Pedimentos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Pedimentos.ColWidth(1) = 800
        Grid_Pedimentos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(2) = 1000
        Grid_Pedimentos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(3) = 1000
        Grid_Pedimentos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(4) = 800
        Grid_Pedimentos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(5) = 1200
        Grid_Pedimentos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(6) = 1200
        Grid_Pedimentos.ColAlignment(6) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(7) = 1000
        Grid_Pedimentos.ColAlignment(7) = flexAlignLeftCenter
        
        
    Else
        MsgBox "La clave no existe", vbExclamation
        Btn_Salir.Caption = "Regresar"
        Btn_Salir_Click
    End If
    Rs_Consulta_Cat_Pedimentos.Close
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Clave.text <> "" Then
        Cambiar_Estado (Txt_Clave.text)
        
    Else
        MsgBox "Ingrese la clave", vbInformation
    End If
End Sub
Public Sub Cambiar_Estado(Cadena As String)
Dim Rs_Estado_Cat_Patentes As rdoResultset   'Manejo de registro
    
    Grid_Pedimentos.Rows = 0
    Mi_SQL = "SELECT Estado from Cat_Numero_Pedimentos_Operados WHERE (Clave =" & Cadena & ")"
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Rs_Consulta_Cat_Patentes.rdoColumns("Estado") = "ACTIVO" Then
        Mi_SQL = "UPDATE Cat_Numero_Pedimentos_Operados SET Estado='INACTIVO' WHERE (Clave =" & Cadena & ")"
    Else
        Mi_SQL = "UPDATE Cat_Numero_Pedimentos_Operados SET Estado='ACTIVO' WHERE (Clave=" & Cadena & ")"
        End If
    Set Rs_Consulta_Cat_Patentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    MsgBox "El estado se ha cambiado con éxito", vbExclamation
    Rs_Consulta_Cat_Patentes.Close
    Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Actualizar" And Txt_Clave.text <> "" And Cmb_Aduanas.ListIndex > -1 And Cmb_Patente.ListIndex > -1 And Cmb_Ejercicio.ListIndex > -1 And Txt_Cantidad.text <> "" Then
                Modifica_Pedimentos
                
            Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
            End If
        
End Sub
Public Sub Modifica_Pedimentos()
Dim Rs_Modificacion_Cat_Pedimentos As rdoResultset   'Manejo de registro de la tabla Cat_Productos
Dim Rs_Consulta_Producto As rdoResultset
Dim Extension As String

    'Consulta el producto seleccionado
    Mi_SQL = "SELECT * FROM Cat_Numero_Pedimentos_Operados"
    Mi_SQL = Mi_SQL & " WHERE Clave=" & Txt_Clave.text & ""
    Set Rs_Modificacion_Cat_Pedimentos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Pedimentos.EOF Then
        'Modifica los datos de la tabla Cat_Productos
        With Rs_Modificacion_Cat_Pedimentos
            .Edit
                .rdoColumns("Aduana_ID") = Cmb_Aduanas.text
                .rdoColumns("Patente") = Cmb_Patente.text
                .rdoColumns("Ejercicio") = Cmb_Ejercicio.text
                .rdoColumns("Cantidad") = Txt_Cantidad.text
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
    Rs_Modificacion_Cat_Pedimentos.Close
    MsgBox "Modificación exitosa", vbInformation
     Btn_Salir.Caption = "Regresar"
    Btn_Salir_Click
    Exit Sub
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
        Cmb_Patente.Enabled = False
        Cmb_Ejercicio.Enabled = False
        Cmb_Aduanas.Enabled = False
        Txt_Cantidad.Enabled = False
        Txt_Clave.Enabled = True
        Dtp_Inicio_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Dtp_Fin_Vigencia.Value = Format(Now(), "yyyy/MMM/dd")
        Grid_Pedimentos.Rows = 0
        Consulta
    End If

End Sub

Private Sub Cmb_Aduanas_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Cmb_Patente_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Cmb_Ejercicio_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789" + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Btn_Nuevo_Click()
 If Btn_Nuevo.Caption = "Nuevo" Then
                If Grid_Pedimentos.Rows <> 0 Then
                    Grid_Pedimentos.Rows = 0
                End If
                Btn_Consultar.Enabled = False
                Btn_Eliminar.Enabled = False
                Cmb_Patente.Enabled = True
                Cmb_Ejercicio.Enabled = True
                Txt_Cantidad.Enabled = True
                Fra_Vigencia.Enabled = True
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Btn_Salir.Caption = "Regresar"
                Btn_Nuevo.Caption = "Dar de Alta"
                Btn_Modificar.Enabled = False
                Chc_Vigencia.Value = 1
                Txt_Clave.text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Numero_Pedimentos_Operados", "Clave"), "00")
                Txt_Clave.Enabled = False
    ElseIf Btn_Nuevo.Caption = "Dar de Alta" And Cmb_Ejercicio.ListIndex > -1 And Trim(Txt_Cantidad.text) <> "" And Cmb_Aduanas.ListIndex > -1 And Cmb_Patente.ListIndex > -1 Then
                Alta_Pedimentos
                
    Else
                MsgBox "Faltan datos para dar de alta", vbInformation
        End If
End Sub
Public Sub Alta_Pedimentos()
Dim Rs_Alta_Cat_Pedimentos As rdoResultset           'Manejo del registro de Cat_Productos, da de alta al producto
Dim Mi_SQL As String
Dim Rs_Consulta_Codigo As rdoResultset
Dim Extension As String

On Error GoTo handler
    'Alta de Producto
    Set Rs_Alta_Cat_Pedimentos = Conectar_Ayudante.Recordset_Agregar("Cat_Numero_Pedimentos_Operados")
    'Llena la tabla de Cat_Aduanas con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Pedimentos
        .AddNew
            
            Clave = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Numero_Pedimentos_Operados", "Clave"), "00")
            .rdoColumns("Clave") = Clave
            .rdoColumns("Patente") = Cmb_Patente.text
            .rdoColumns("Aduana_ID") = Cmb_Aduanas.text
            .rdoColumns("Fecha_Inicio_Vigencia") = Format(Dtp_Inicio_Vigencia.Value, "yyyy/MM/dd")
            If Chc_Vigencia.Value = 1 Then
            .rdoColumns("Fecha_Fin_Vigencia") = Null
            Else
            .rdoColumns("Fecha_Fin_Vigencia") = Format(Dtp_Fin_Vigencia.Value, "yyyy/MM/dd")
            End If
            .rdoColumns("Ejercicio") = Cmb_Ejercicio.text
            .rdoColumns("Cantidad") = Txt_Cantidad.text
        
        .Update
    End With
    'Cierra el manejador del registro
    Rs_Alta_Cat_Pedimentos.Close
    'Deshabilita controles y habilita otros que será necesarios
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Actualizar"
    'Coloca un encabezado en la primera fila del grid
    If Grid_Pedimentos.Rows = 0 Then
        Grid_Pedimentos.AddItem "Clave" & Chr(9) & "Aduana" & Chr(9) & "Patente" & Chr(9) & "Ejercicio" & Chr(9) & "Cantidad" & Chr(9) & "Inicio Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
    End If
    'Llena el grid con los datos capturados
    If Chc_Vigencia.Value = 1 Then
            Grid_Pedimentos.AddItem Clave & Chr(9) & UCase(Cmb_Aduanas.text) & Chr(9) & UCase(Trim(Cmb_Patente.text)) & Chr(9) & UCase(Trim(Cmb_Ejercicio.text)) & Chr(9) & UCase(Trim(Txt_Cantidad.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase("") & Chr(9) & UCase("ACTIVO")
            Else
            Grid_Pedimentos.AddItem Clave & Chr(9) & UCase(Cmb_Aduanas.text) & Chr(9) & UCase(Trim(Cmb_Patente.text)) & Chr(9) & UCase(Trim(Cmb_Ejercicio.text)) & Chr(9) & UCase(Trim(Txt_Cantidad.text)) & Chr(9) & UCase(Dtp_Inicio_Vigencia.Value) & Chr(9) & UCase(Dtp_Fin_Vigencia.Value) & Chr(9) & UCase("ACTIVO")
            End If
    
         Grid_Pedimentos.FixedCols = 1
        Grid_Pedimentos.ColWidth(0) = 800
        Grid_Pedimentos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Pedimentos.ColWidth(1) = 800
        Grid_Pedimentos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(2) = 1000
        Grid_Pedimentos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(3) = 1000
        Grid_Pedimentos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(4) = 800
        Grid_Pedimentos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(5) = 1200
        Grid_Pedimentos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(6) = 1200
        Grid_Pedimentos.ColAlignment(6) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(7) = 1000
        Grid_Pedimentos.ColAlignment(7) = flexAlignLeftCenter
    MsgBox "Registro exitoso", vbInformation
    Exit Sub
'Ante error realiza un rollback en la transacción y no hace cambios en la base de datos
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
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
    Llena_Combos
    Chc_Vigencia.Value = 1
    Dtp_Inicio_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
    Dtp_Fin_Vigencia.Value = Format(Now, "yyyy/MMM/dd")
End Sub
Public Sub Consulta()
Dim Rs_Consulta_Cat_Pedimentos As rdoResultset   'Manejo de registro
    
    Grid_Pedimentos.Rows = 0
    'Consulta el producto de acuerdo a la descripción proporcionada
    'Cadena = Conectar_Ayudante.Quitar_Caracter(Cadena, "'")
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Numero_Pedimentos_Operados ORDER BY Clave"
    'Le envía la consulta al ayudante para que realice la consulta
    Set Rs_Consulta_Cat_Pedimentos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Revisa que no sea fin de archivo para llenar el grid
    If Not Rs_Consulta_Cat_Pedimentos.EOF Then
        'Pone un encabezado en el grid
        Grid_Pedimentos.AddItem "Clave" & Chr(9) & "Aduana" & Chr(9) & "Patente" & Chr(9) & "Ejercicio" & Chr(9) & "Cantidad" & Chr(9) & "Inicio de Vigencia" & Chr(9) & "Fin de Vigencia" & Chr(9) & "Estado"
        'Llenado del grid
        While Not Rs_Consulta_Cat_Pedimentos.EOF
            With Rs_Consulta_Cat_Pedimentos
                Grid_Pedimentos.AddItem .rdoColumns("Clave") & Chr(9) & .rdoColumns("Aduana_ID") & Chr(9) & .rdoColumns("Patente") & Chr(9) & .rdoColumns("Ejercicio") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Fecha_Inicio_Vigencia") & Chr(9) & .rdoColumns("Fecha_Fin_Vigencia") & Chr(9) & .rdoColumns("Estado")
                Grid_Pedimentos.FixedRows = 1
                Rs_Consulta_Cat_Pedimentos.MoveNext
            End With
        Wend
        'Tamaño de las columnas en el grid
        Grid_Pedimentos.FixedCols = 1
        Grid_Pedimentos.ColWidth(0) = 800
        Grid_Pedimentos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Pedimentos.ColWidth(1) = 800
        Grid_Pedimentos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(2) = 1000
        Grid_Pedimentos.ColAlignment(2) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(3) = 1000
        Grid_Pedimentos.ColAlignment(3) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(4) = 800
        Grid_Pedimentos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(5) = 1200
        Grid_Pedimentos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(6) = 1200
        Grid_Pedimentos.ColAlignment(6) = flexAlignLeftCenter
        Grid_Pedimentos.ColWidth(7) = 1000
        Grid_Pedimentos.ColAlignment(7) = flexAlignLeftCenter
        
    End If
    Rs_Consulta_Cat_Pedimentos.Close
End Sub
Public Sub Llena_Combos()
Dim i As Integer
Cmb_Ejercicio.text = ""
    Call Conectar_Ayudante.Llena_Combo_Item("Clave, Aduana_ID", "Cat_Aduanas ORDER BY Clave", Cmb_Aduanas, 0, "")
    Call Conectar_Ayudante.Llena_Combo_Item("Clave, Codigo_Patente_Aduanal", "Cat_Patentes_Aduanales ORDER BY Clave", Cmb_Patente, 0, "")
    For i = 2000 To Year(Now)
        Cmb_Ejercicio.AddItem (i)
    Next
End Sub

Private Sub Grid_Pedimentos_Click()
Dim Rs_Consulta_Cat_Pedimentos As rdoResultset
    
    'Si el grid tiene filas, entonces hace la consulta
    
    If Grid_Pedimentos.Rows > 1 Then
        Btn_Salir.Caption = "Regresar"
        Cmb_Patente.Enabled = True
        Cmb_Ejercicio.Enabled = True
        Txt_Cantidad.Enabled = True
        Fra_Vigencia.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = True
        'Selecciona los campos de la tabla de Cat_Almacenes
        Mi_SQL = "SELECT * FROM Cat_Numero_Pedimentos_Operados"
        Mi_SQL = Mi_SQL & " WHERE Clave=" & Grid_Pedimentos.TextMatrix(Grid_Pedimentos.RowSel, 0) & ""
        Set Rs_Consulta_Cat_Pedimentos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto de acuerdo a los datos obtenidos de la consulta
        If Not Rs_Consulta_Cat_Pedimentos.EOF Then
            With Rs_Consulta_Cat_Pedimentos
                For i = 0 To Cmb_Aduanas.ListCount - 1
                     Cmb_Aduanas.ListIndex = i
                        If Cmb_Aduanas.text = .rdoColumns("Aduana_ID") Then
                            Exit For
                        End If
                Next
                 For i = 0 To Cmb_Patente.ListCount - 1
                     Cmb_Patente.ListIndex = i
                        If Cmb_Patente.text = .rdoColumns("Patente") Then
                            Exit For
                        End If
                Next
                For i = 0 To Cmb_Ejercicio.ListCount - 1
                     Cmb_Ejercicio.ListIndex = i
                        If Cmb_Ejercicio.text = .rdoColumns("Ejercicio") Then
                            Exit For
                        End If
                Next
                Txt_Cantidad.text = .rdoColumns("Cantidad")
                Txt_Clave.text = .rdoColumns("Clave")
                Dtp_Inicio_Vigencia.Value = .rdoColumns("Fecha_Inicio_Vigencia")
                Dtp_Fin_Vigencia.MinDate = Dtp_Inicio_Vigencia.Value
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
        Rs_Consulta_Cat_Pedimentos.Close
    End If
End Sub
Private Sub Txt_Cantidad_KeyPress(KeyAscii As Integer)
Dim Cadena As String

    'SOLO PERMITE ESCRIBIR LOS CARACTERES DE LA CADENA
    Cadena = "0123456789." + Chr(8)       'chr(8) = delete, es decir admitimos borrar
    If InStr(Cadena, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
