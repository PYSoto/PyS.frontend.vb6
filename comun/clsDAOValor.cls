VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vAutoID As Long
Private vNegocioID As Integer
Private vValorID As Integer
Private vConcepto As String
Private vNumerable As Integer
Private vDuplicados As Integer
Private vFechaEmision As Integer
Private vFechaVencimiento As Integer
Private vTitular As Integer
Private vBanco As Integer
Private vChequeTercero As Integer
Private vCuentaCorriente As Integer
Private vCreated As Date

Private clon As clsDAOValor

Public ddl As New clsDDLValor

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vAutoID = 0

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
    vConcepto = ""
    vNumerable = 0
    vDuplicados = 0
    vFechaEmision = 0
    vFechaVencimiento = 0
    vTitular = 0
    vBanco = 0
    vChequeTercero = 0
    vCuentaCorriente = 0
    vCreated = Date

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
' Description: Getter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

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

Public Property Get key() As String

    key = "k." & valorID
    
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
' Description: Getter concepto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get concepto() As Variant

    concepto = vConcepto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter concepto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let concepto(ByVal vNewValue As Variant)

On Error Resume Next

    vConcepto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter numerable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerable() As Variant

    numerable = vNumerable

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter numerable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let numerable(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerable = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter duplicados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get duplicados() As Variant

    duplicados = vDuplicados

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter duplicados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let duplicados(ByVal vNewValue As Variant)

On Error Resume Next

    vDuplicados = vNewValue

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
' Description: Getter chequeTercero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get chequeTercero() As Variant

    chequeTercero = vChequeTercero

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter chequeTercero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let chequeTercero(ByVal vNewValue As Variant)

On Error Resume Next

    vChequeTercero = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter cuentaCorriente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaCorriente() As Variant

    cuentaCorriente = vCuentaCorriente

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter cuentaCorriente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cuentaCorriente(ByVal vNewValue As Variant)

On Error Resume Next

    vCuentaCorriente = vNewValue

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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = concepto & " (" & valorID & ")"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = valorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = concepto & " (" & valorID & ")"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = valorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOValor
Dim objObjeto As New clsDAOValor

    With objObjeto
        .autoID = vAutoID
        .negocioID = vNegocioID
        .valorID = vValorID
        .concepto = vConcepto
        .numerable = vNumerable
        .duplicados = vDuplicados
        .fechaEmision = vFechaEmision
        .fechaVencimiento = vFechaVencimiento
        .titular = vTitular
        .banco = vBanco
        .chequeTercero = vChequeTercero
        .cuentaCorriente = vCuentaCorriente
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
Private Sub fillObject(pObjeto As clsDAOValor, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .autoID = pQuery.Fields(.ddl.autoID)
        .negocioID = pQuery.Fields(.ddl.negocioID)
        .valorID = pQuery.Fields(.ddl.valorID)
        .concepto = pQuery.Fields(.ddl.concepto)
        .numerable = pQuery.Fields(.ddl.numerable)
        .duplicados = pQuery.Fields(.ddl.duplicados)
        .fechaEmision = pQuery.Fields(.ddl.fechaEmision)
        .fechaVencimiento = pQuery.Fields(.ddl.fechaVencimiento)
        .titular = pQuery.Fields(.ddl.titular)
        .banco = pQuery.Fields(.ddl.banco)
        .chequeTercero = pQuery.Fields(.ddl.chequeTercero)
        .cuentaCorriente = pQuery.Fields(.ddl.cuentaCorriente)
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
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando 1 registro
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
' Description: Consulta la DB buscando el último autonumérico agregado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLastID(pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.autoID
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vAutoID = rstQuery.Fields(ddl.autoID)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByValorID(pDB As clsDB)
Dim sql As String

Dim intValorID As Integer

    intValorID = vValorID

    clean

    vValorID = intValorID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.valorID & " = " & vValorID
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Elimina 1 registro de la DB según la Primary Key
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
    sql = sql & ", " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
    sql = sql & ", " & ddl.numerable & " = " & vNumerable
    sql = sql & ", " & ddl.duplicados & " = " & vDuplicados
    sql = sql & ", " & ddl.fechaEmision & " = " & vFechaEmision
    sql = sql & ", " & ddl.fechaVencimiento & " = " & vFechaVencimiento
    sql = sql & ", " & ddl.titular & " = " & vTitular
    sql = sql & ", " & ddl.banco & " = " & vBanco
    sql = sql & ", " & ddl.chequeTercero & " = " & vChequeTercero
    sql = sql & ", " & ddl.cuentaCorriente & " = " & vCuentaCorriente
    sql = sql & ";"

    add = pDB.execute(sql)

    findLastID pDB

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
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
    If vConcepto <> clon.concepto Or IsNull(vConcepto) <> IsNull(clon.concepto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
        blnChange = True
    End If
    If vNumerable <> clon.numerable Or IsNull(vNumerable) <> IsNull(clon.numerable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.numerable & " = " & vNumerable
        blnChange = True
    End If
    If vDuplicados <> clon.duplicados Or IsNull(vDuplicados) <> IsNull(clon.duplicados) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.duplicados & " = " & vDuplicados
        blnChange = True
    End If
    If vFechaEmision <> clon.fechaEmision Or IsNull(vFechaEmision) <> IsNull(clon.fechaEmision) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaEmision & " = " & vFechaEmision
        blnChange = True
    End If
    If vFechaVencimiento <> clon.fechaVencimiento Or IsNull(vFechaVencimiento) <> IsNull(clon.fechaVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaVencimiento & " = " & vFechaVencimiento
        blnChange = True
    End If
    If vTitular <> clon.titular Or IsNull(vTitular) <> IsNull(clon.titular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.titular & " = " & vTitular
        blnChange = True
    End If
    If vBanco <> clon.banco Or IsNull(vBanco) <> IsNull(clon.banco) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.banco & " = " & vBanco
        blnChange = True
    End If
    If vChequeTercero <> clon.chequeTercero Or IsNull(vChequeTercero) <> IsNull(clon.chequeTercero) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.chequeTercero & " = " & vChequeTercero
        blnChange = True
    End If
    If vCuentaCorriente <> clon.cuentaCorriente Or IsNull(vCuentaCorriente) <> IsNull(clon.cuentaCorriente) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuentaCorriente & " = " & vCuentaCorriente
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

    vAutoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vAutoID = 1 + rstQuery.Fields(ddl.autoID)

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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOValor

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOValor
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
Dim elemento As clsDAOValor

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
Dim elemento As clsDAOValor

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

