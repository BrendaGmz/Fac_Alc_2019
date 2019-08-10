VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conectar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ConectarBD()
   Conexion
End Function

Public Sub Conexion()
Dim I As Integer        'Cuenta cada una de las lineas del documento Confing.ini
Dim linea As String     'Guarda lo que contenga la linea a la cual se esta apuntando
Dim Servidor As String  'Guarda el nombre del servidor
Dim Base As String      'Guarda el nombre de la base de datos
Dim Tipo  As String     'Identifica en que esta hecha la BD Access o SQL
Dim Password As String  'Guarda la contraseña para el acceso al servidor
Dim Usuario_Conexion As String   'Guarda el nombre válido del usuario para el acceso a la BD
Dim BaseD As String
    
On Error GoTo Etiqueta
    I = 0
    'LEE EL ARCHIVO DE INICIO DEL SISTEMA
    Open App.Path & "\Config.ini" For Input As #1
        Do While Not EOF(1)
            Line Input #1, linea
            If I = 0 Then Empresa = Trim(linea)
            If I = 1 Then RFC = Trim(linea)
            If I = 2 Then Direccion = Trim(linea)
            If I = 3 Then Direccion = Direccion & " " & Trim(linea)
            If I = 4 Then CP = Trim(linea)
            If I = 5 Then Telefono = Trim(linea)
            If I = 6 Then Ciudad_Edo = Trim(linea)
            If I = 7 Then Ciudad_Edo = Ciudad_Edo & " " & Trim(linea)
            If I = 8 Then Tipo = Trim(linea)
            If I = 9 Then Servidor = Trim(linea)
            If I = 10 Then Usuario_Conexion = Trim(linea)
            If I = 11 Then Base = Trim(linea)
            If I = 12 Then Password = Trim(linea)
            If I = 13 Then Regimen_Fiscal = Trim(linea)
            I = I + 1
        Loop
    Close #1
    BaseD = Base
    Conexion_Base.CursorDriver = rdUseOdbc
    'Cadena de conexión
    Conexion_Base.Connect = "uid=" & Usuario_Conexion & ";pwd=" & Password & ";server=" & Servidor & ";" _
                          & "driver={SQL Server};protocol=6.3;readonly=false;database=" & Base & ";dsn=;"
    Par_Fecha = "'"
    Server = Servidor
    Database = Base
    User_Password = Password
    User_Conexion = Usuario_Conexion
    Base_Datos = "S"
    Conexion_Base.EstablishConnection
    Conexion_Base.QueryTimeout = 80
    'Load Frm_Apl_Login
    MsgBox "Conexion exitosa", vbExclamation
Exit Sub
Etiqueta:
    End
End Sub

