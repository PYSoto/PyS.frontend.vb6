VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private vDBConn As New ADODB.Connection
Private vSConection As String
Private vDsn As String
Private vUid As String
Private vPwd As String
Private vIp As String
Private vDatabase As String
Private vAU As clsDB

Private vCounterTransaction As Integer

Private Sub class_Initialize()
    
    vDsn = ""
    vUid = ""
    vPwd = ""
    vIp = ""
    vDatabase = ""
    Set vAU = Nothing

    vCounterTransaction = 0

End Sub

Property Get connected() As Boolean
    
    If vDBConn Is Nothing Then
        connected = False
    ElseIf vDBConn.State = adStateClosed Then
        connected = False
    Else
        connected = True
    End If

End Property

Property Get sconection() As String
    
    sconection = vSConection

End Property

Public Property Let dsn(ByVal pDSN As String)
    
    vDsn = pDSN

End Property

Public Property Get dsn() As String

    dsn = vDsn

End Property

Public Property Let uid(ByVal pUID As String)
    
    vUid = pUID

End Property

Public Property Get uid() As String

    uid = vUid

End Property

Public Property Let pwd(ByVal pPWD As String)
    
    vPwd = pPWD

End Property

Public Property Get pwd() As String

    pwd = vPwd

End Property

Public Property Let ip(ByVal pIP As String)
    
    vIp = pIP

End Property

Public Property Get ip() As String

    ip = vIp

End Property

Public Property Let database(ByVal pDatabase As String)
    
    vDatabase = pDatabase

End Property

Public Property Get database() As String

    database = vDatabase

End Property

Public Sub makeConnectionString()
        
    vSConection = "DSN=" & vDsn & ";UID=" & vUid & ";PWD=" & vPwd & ";DATABASE=" & vDatabase & ";"

End Sub

Public Property Let au(ByVal vNewValue As Variant)

    Set vAU = vNewValue
    
End Property

Public Sub openDB()

On Error GoTo Mensaje
    
    Screen.MousePointer = 11
    
    If connected Then Exit Sub
    
    registerODBC5 vDsn, "MySQL ODBC 5.1 Driver", vDatabase, vIp
    
    Me.makeConnectionString
    
    With vDBConn
        .ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & vIp & ";Option=3;Database=" & vDatabase & ";User=" & vUid & ";Password=" & vPwd & ";"
        .CursorLocation = adUseServer
        .Open
    End With

salir:
    Screen.MousePointer = 0
    Exit Sub
Mensaje:
    MsgBox "ERROR: " & Now & " - " & Err.Number & " - " & Err.Description
    Resume salir

End Sub

Public Sub closeDB()
    
    If connected Then vDBConn.Close
    Set vDBConn = Nothing

End Sub

Public Function execute(pSQL As String) As Boolean

On Error GoTo iError:

    If Not connected Then Me.openDB
    If Not connected Then Exit Function
    
    Screen.MousePointer = 11
    
    If Right(pSQL, 1) <> ";" Then
        MsgBox "ERROR : Cadena SQL dañada" & Chr(13) & Chr(13) & "Favor INFORMAR" & Chr(13) & Chr(13) & pSQL
        Exit Function
    End If
    
    vDBConn.execute pSQL
    
    execute = True
    
    If Not vAU Is Nothing Then gCon.saveAuditoria pSQL

fin:
    Screen.MousePointer = 0
    Exit Function
iError:
    rollBack
    execute = False
    MsgBox Err.Number & " - " & Err.Description & Chr(13) & Chr(13) & pSQL, , "ERROR"
    Resume fin
    
End Function

Public Function query(pSQL As String) As ADODB.Recordset
Dim rstQuery As New ADODB.Recordset

On Error GoTo Mensaje:
    
    If Not connected Then Me.openDB
    If Not connected Then Exit Function
    
    rstQuery.Open pSQL, vDBConn, adOpenForwardOnly
    Set query = rstQuery
    Exit Function

Mensaje:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next

End Function

Property Get activa() As ADODB.Connection
    
    Set activa = vDBConn

End Property

Public Sub beginTrans()
    
    vDBConn.execute "START TRANSACTION;"

End Sub

Public Sub commitTrans()
    
    vDBConn.execute "COMMIT;"

End Sub

Public Sub rollBack()
    
    vDBConn.execute "ROLLBACK;"
    
    MsgBox "ERROR: La TRANSACCION NO fue realizada . . ."

End Sub

