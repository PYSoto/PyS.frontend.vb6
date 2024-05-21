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

Private vTablaAuditoria As String

Public Sub configure_dbmdzsh()

    loadParametersDB "pys-mdz-sa", dbmdzsh, "pysmdzsa"
    
End Sub

Public Sub configure_dbmdzsrl()

    loadParametersDB "pys-mdz-srl", dbmdzsrl, "pysmdzsa"
    
End Sub

Public Sub configure_dbuaqsrl()

    loadParametersDB "pys-uaq-srl", dbuaqsrl, "pysuaqsrl"
    
End Sub

Private Sub loadParametersDB(dsn_name As String, db As clsDB, defaultDatabase As String)
Dim success As Long
    
Dim Reply As ICMP_ECHO_REPLY

    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DSN") = "" Then defaultParametersDB dsn_name, defaultDatabase
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DB") = "" Then defaultParametersDB dsn_name, defaultDatabase
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "UID") = "" Then defaultParametersDB dsn_name, defaultDatabase
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "IP") = "" Then defaultParametersDB dsn_name, defaultDatabase
    
    db.dsn = GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DSN")
    db.database = GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DB")
    db.uid = GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "UID")
    db.pwd = GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "PWD")
    db.ip = GetSettingString(HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "IP")
    
    ' Verifica la IP
    success = ping(db.ip, Reply)
    If success <> 0 Then
        MsgBox "ERROR: Negocio '" & db.ip & "' " & evaluatePingResponse(success)
        End
    End If
    
End Sub

Private Sub defaultParametersDB(dsn_name As String, defaultDatabase As String)
    
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DSN", dsn_name
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "DB", defaultDatabase
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "UID", "sistemas"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "PWD", "iveco"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\" & dsn_name, "IP", "127.0.0.1"

End Sub

Public Sub saveAuditoria(ByVal pSQL As String)
Dim strSQL As String

    makeTable
    
    strSQL = "INSERT INTO " & vTablaAuditoria
    strSQL = strSQL & " SET aud_usuario = 'integrador'"
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
    
    vTablaAuditoria = "integ" & Trim(Str(Year(datFecha) * 100 + Month(datFecha)))

End Sub


