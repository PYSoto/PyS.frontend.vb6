VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOValorMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNegocioID As Integer
Private vValorID As Integer
Private vClienteID As Long
Private vProveedorID As Long
Private vFechaEmision As Variant
Private vFechaVencimiento As Variant
Private vComprobanteID As Integer
Private vNroComprobante As Long
Private vImporte As Currency
Private vFechaRegistracion As Date
Private vClientemovimientoID As Long
Private vProveedormovimientoID As Long
Private vTitular As String
Private vBanco As String
Private vValormovimientoID As Long
Private vCreated As Date
' Atributos de memoria
Private vKeyCollection As Long

Private clon As clsDAOValorMov

Public ddl As New clsDDLValorMov

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vValormovimientoID = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vNegocioID = 0
    vValorID = 0
    vClienteID = 0
    vProveedorID = 0
    vFechaEmision = Null
    vFechaVencimiento = Null
    vComprobanteID = 0
    vNroComprobante = 0
    vImporte = 0
    vFechaRegistracion = Date
    vClientemovimientoID = 0
    vProveedormovimientoID = 0
    vTitular = ""
    vBanco = ""
    vCreated = Date
    vKeyCollection = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter negocioID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID() As Variant

    negocioID = vNegocioID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter negocioID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let negocioID(ByVal vNewValue As Variant)

On Error Resume Next

    vNegocioID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter valorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorID() As Variant

    valorID = vValorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter valorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valorID(ByVal vNewValue As Variant)

On Error Resume Next

    vValorID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter proveedorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID() As Variant

    proveedorID = vProveedorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedorID(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedorID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter fechaEmision
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaEmision() As Variant

    fechaEmision = vFechaEmision

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter fechaEmision
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaEmision(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaEmision = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter fechaVencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaVencimiento() As Variant

    fechaVencimiento = vFechaVencimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter fechaVencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaVencimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comprobanteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID() As Variant

    comprobanteID = vComprobanteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter comprobanteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let comprobanteID(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter nroComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nroComprobante() As Variant

    nroComprobante = vNroComprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter nroComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nroComprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter importe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importe() As Variant

    importe = vImporte

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter importe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter fechaRegistracion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaRegistracion() As Variant

    fechaRegistracion = vFechaRegistracion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter fechaRegistracion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaRegistracion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaRegistracion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter clientemovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clientemovimientoID() As Variant

    clientemovimientoID = vClientemovimientoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter clientemovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clientemovimientoID(ByVal vNewValue As Variant)

On Error Resume Next

    vClientemovimientoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter proveedormovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedormovimientoID() As Variant

    proveedormovimientoID = vProveedormovimientoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedormovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedormovimientoID(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedormovimientoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter titular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get titular() As Variant

    titular = vTitular

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter titular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let titular(ByVal vNewValue As Variant)

On Error Resume Next

    vTitular = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter banco
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get banco() As Variant

    banco = vBanco

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter banco
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let banco(ByVal vNewValue As Variant)

On Error Resume Next

    vBanco = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter valormovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valormovimientoID() As Variant

    valormovimientoID = vValormovimientoID

End Property

Public Property Get key() As String

    key = "k." & valormovimientoID
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter valormovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valormovimientoID(ByVal vNewValue As Variant)

On Error Resume Next

    vValormovimientoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As Variant

    created = vCreated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get keyCollection() As Variant

    keyCollection = vKeyCollection

End Property

Public Property Let keyCollection(ByVal vNewValue As Variant)

On Error Resume Next

    vKeyCollection = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vNegocioID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vValormovimientoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vNegocioID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vValormovimientoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOValorMov
Dim objObjeto As New clsDAOValorMov

    With objObjeto
        .negocioID = vNegocioID
        .valorID = vValorID
        .clienteID = vClienteID
        .proveedorID = vProveedorID
        .fechaEmision = vFechaEmision
        .fechaVencimiento = vFechaVencimiento
        .comprobanteID = vComprobanteID
        .nroComprobante = vNroComprobante
        .importe = vImporte
        .fechaRegistracion = vFechaRegistracion
        .clientemovimientoID = vClientemovimientoID
        .proveedormovimientoID = vProveedormovimientoID
        .titular = vTitular
        .banco = vBanco
        .valormovimientoID = vValormovimientoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(pObjeto As clsDAOValorMov, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .negocioID = pQuery.Fields(.ddl.negocioID)
        .valorID = pQuery.Fields(.ddl.valorID)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .proveedorID = pQuery.Fields(.ddl.proveedorID)
        .fechaEmision = pQuery.Fields(.ddl.fechaEmision)
        .fechaVencimiento = pQuery.Fields(.ddl.fechaVencimiento)
        .comprobanteID = pQuery.Fields(.ddl.comprobanteID)
        .nroComprobante = pQuery.Fields(.ddl.nroComprobante)
        .importe = pQuery.Fields(.ddl.importe)
        .fechaRegistracion = pQuery.Fields(.ddl.fechaRegistracion)
        .clientemovimientoID = pQuery.Fields(.ddl.clientemovimientoID)
        .proveedormovimientoID = pQuery.Fields(.ddl.proveedormovimientoID)
        .titular = pQuery.Fields(.ddl.titular)
        .banco = pQuery.Fields(.ddl.banco)
        .valormovimientoID = pQuery.Fields(.ddl.valormovimientoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Construye el SQL de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.valormovimientoID & " = " & vValormovimientoID

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findSQL(pSQL As String, pDB As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = pDB.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB buscando el último autonumérico agregado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLastID(pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.valormovimientoID
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vValormovimientoID = rstQuery.Fields(ddl.valormovimientoID)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.valormovimientoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Elimina 1 registro de la pDB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(pDB As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = pDB.execute(sql)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(pDB As clsDB) As Boolean

    If exist(pDB) Then
        save = update(pDB)
    Else
        save = add(pDB)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(pDB As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(pDB As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.negocioID & " = " & vNegocioID
    sql = sql & ", " & ddl.valorID & " = " & vValorID
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & ", " & ddl.fechaEmision & " = " & fechaDB(vFechaEmision)
    sql = sql & ", " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
    sql = sql & ", " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & ", " & ddl.nroComprobante & " = " & vNroComprobante
    sql = sql & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    sql = sql & ", " & ddl.fechaRegistracion & " = " & fechaDB(vFechaRegistracion)
    sql = sql & ", " & ddl.clientemovimientoID & " = " & vClientemovimientoID
    sql = sql & ", " & ddl.proveedormovimientoID & " = " & vProveedormovimientoID
    sql = sql & ", " & ddl.titular & " = '" & replaceSpecialSymbols(vTitular) & "'"
    sql = sql & ", " & ddl.banco & " = '" & replaceSpecialSymbols(vBanco) & "'"
    sql = sql & ";"

    add = pDB.execute(sql)

    findLastID pDB

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la pDB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(pDB As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vNegocioID <> clon.negocioID Or IsNull(vNegocioID) <> IsNull(clon.negocioID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negocioID & " = " & vNegocioID
        blnChange = True
    End If
    If vValorID <> clon.valorID Or IsNull(vValorID) <> IsNull(clon.valorID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.valorID & " = " & vValorID
        blnChange = True
    End If
    If vClienteID <> clon.clienteID Or IsNull(vClienteID) <> IsNull(clon.clienteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteID & " = " & vClienteID
        blnChange = True
    End If
    If vProveedorID <> clon.proveedorID Or IsNull(vProveedorID) <> IsNull(clon.proveedorID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.proveedorID & " = " & vProveedorID
        blnChange = True
    End If
    If vFechaEmision <> clon.fechaEmision Or IsNull(vFechaEmision) <> IsNull(clon.fechaEmision) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaEmision & " = " & fechaDB(vFechaEmision)
        blnChange = True
    End If
    If vFechaVencimiento <> clon.fechaVencimiento Or IsNull(vFechaVencimiento) <> IsNull(clon.fechaVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
        blnChange = True
    End If
    If vComprobanteID <> clon.comprobanteID Or IsNull(vComprobanteID) <> IsNull(clon.comprobanteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.comprobanteID & " = " & vComprobanteID
        blnChange = True
    End If
    If vNroComprobante <> clon.nroComprobante Or IsNull(vNroComprobante) <> IsNull(clon.nroComprobante) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroComprobante & " = " & vNroComprobante
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vFechaRegistracion <> clon.fechaRegistracion Or IsNull(vFechaRegistracion) <> IsNull(clon.fechaRegistracion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaRegistracion & " = " & fechaDB(vFechaRegistracion)
        blnChange = True
    End If
    If vClientemovimientoID <> clon.clientemovimientoID Or IsNull(vClientemovimientoID) <> IsNull(clon.clientemovimientoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clientemovimientoID & " = " & vClientemovimientoID
        blnChange = True
    End If
    If vProveedormovimientoID <> clon.proveedormovimientoID Or IsNull(vProveedormovimientoID) <> IsNull(clon.proveedormovimientoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.proveedormovimientoID & " = " & vProveedormovimientoID
        blnChange = True
    End If
    If vTitular <> clon.titular Or IsNull(vTitular) <> IsNull(clon.titular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.titular & " = '" & replaceSpecialSymbols(vTitular) & "'"
        blnChange = True
    End If
    If vBanco <> clon.banco Or IsNull(vBanco) <> IsNull(clon.banco) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.banco & " = '" & replaceSpecialSymbols(vBanco) & "'"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update = pDB.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Busca un nuevo ID según la Primary Key y lo carga al objeto si se le indica como parámetro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub newID(pFind As Boolean, pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vValormovimientoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.valormovimientoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vValormovimientoID = 1 + rstQuery.Fields(ddl.valormovimientoID)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Arma una colección con 1 registro para verificar si algo existe en la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAny(pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, pDB)

End Function

Public Function collectionByProveedorMovimientoID(proveedormovimientoID As Long, pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedormovimientoID & " = " & proveedormovimientoID
    sql = sql & ";"

    Set collectionByProveedorMovimientoID = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la pDB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOValorMov

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOValorMov
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, pDB As clsDB)

    fillComboCollection pCombo, collectionAll(pDB)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsDAOValorMov

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, pDB As clsDB)

    fillListCollection pList, collectionAll(pDB)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsDAOValorMov

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

