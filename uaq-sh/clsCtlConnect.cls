VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPuntovta As Integer
Private vBackendServer As String
Private vBackendPort As String

Private vTablaAuditoria As String

Private Sub class_Initialize()

    vPuntovta = 0
    
End Sub

Public Property Get puntovta() As Variant

    puntovta = vPuntovta
    
End Property

Public Property Get backend_server() As String

    backend_server = vBackendServer
    
End Property

Public Property Get backend_port() As String

    backend_port = vBackendPort
    
End Property

Public Sub configureDB()

    loadParametersDB
    
    db.au = au
    
    configureAU
    
End Sub

Private Sub configureAU()

    au.dsn = "auditoria"
    
    loadParametersAU

End Sub

Private Sub loadParametersDB()
Dim lngSuccess As Long
    
Dim Reply As ICMP_ECHO_REPLY

    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DSN") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DB") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "UID") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "IP") = "" Then defaultParametersDB
    
    db.dsn = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DSN")
    db.database = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DB")
    db.uid = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "UID")
    db.pwd = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "PWD")
    db.ip = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "IP")
    
    ' Verifica la IP
    lngSuccess = ping(db.ip, Reply)
    If lngSuccess <> 0 Then
        MsgBox "ERROR: Negocio '" & db.ip & "' " & evaluatePingResponse(lngSuccess)
        End
    End If
    
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "PuntoVta") = "" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "PuntoVta", ""
    vPuntovta = Val(GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "PuntoVta"))
    
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_server") = "" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_server", "10.147.20.24"
    vBackendServer = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_server")
    
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_port") = "" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_port", "8080"
    vBackendPort = GetSettingString(HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "backend_port")
    
End Sub

Private Sub defaultParametersDB()
    
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DSN", "pys-uaq-sa"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "DB", "pysuaqsa"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "UID", "sistemas"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "PWD", "iveco"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\pys-uaq-sa", "IP", "127.0.0.1"

End Sub

Private Sub loadParametersAU()
    
    au.database = "auditoria"
    au.uid = db.uid
    au.pwd = db.pwd
    au.ip = db.ip
    
End Sub

Public Function fechaServer() As Date
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    fechaServer = Date
    
    strSQL = "SELECT CurDate() AS fecha;"
    
    Set rstQuery = db.query(strSQL)
    
    If Not rstQuery.EOF Then fechaServer = rstQuery!fecha
    
    rstQuery.Close

End Function

Public Sub saveAuditoria(ByVal pSQL As String)
Dim strSQL As String

    makeTable
    
    strSQL = "INSERT INTO " & vTablaAuditoria
    strSQL = strSQL & " SET aud_usuario = '" & gUsuario.nombre & "'"
    strSQL = strSQL & ", aud_que = '" & replaceSpecialSymbols(pSQL) & "'"
    strSQL = strSQL & ", aud_ip = '" & modIP.getIPAddress & "'"
    strSQL = strSQL & ", aud_cuando = Now()"
    strSQL = strSQL & ";"
    
    au.execute strSQL
    
End Sub

Public Sub makeTable()
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    makeTableName

    strSQL = "CREATE TABLE IF NOT EXISTS " & vTablaAuditoria
    strSQL = strSQL & " (Aud_ID int(4) auto_increment primary key,"
    strSQL = strSQL & " Aud_Usuario varchar(15),"
    strSQL = strSQL & " Aud_IP varchar(20),"
    strSQL = strSQL & " Aud_Cuando timestamp,"
    strSQL = strSQL & " Aud_Que text,"
    strSQL = strSQL & " index (aud_usuario),"
    strSQL = strSQL & " index (aud_cuando));"
    
    au.execute strSQL
    
End Sub

Public Sub makeTableName()
Dim datFecha As Date

Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    datFecha = Date
    
    strSQL = "SELECT CurDate() AS Fecha;"
    
    Set rstQuery = au.query(strSQL)
    If Not rstQuery.EOF Then datFecha = rstQuery!fecha
    rstQuery.Close
    
    vTablaAuditoria = "uash" & Trim(Str(Year(datFecha) * 100 + Month(datFecha)))

End Sub

Public Sub configureNegocio(pNeg As clsDAONegocio, pDB As clsDB)

    With pDB
        .dsn = pNeg.dsn
        .ip = pNeg.ip
        .database = pNeg.database
        .uid = pNeg.user
        .pwd = "iveco"
    End With
        
End Sub




