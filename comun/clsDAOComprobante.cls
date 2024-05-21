VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNegocioID As Integer
Private vComprobanteID As Long
Private vDescripcion As String
Private vTipoTransaccionID As Integer
Private vCuentaCorriente As Integer
Private vDebita As Integer
Private vIva As Integer
Private vAplicaPendiente As Integer
Private vAplicable As Integer
Private vAplicacion As Integer
Private vLibroIva As Integer
Private vLetraComprobante As String
Private vRecibo As Integer
Private vContado As Integer
Private vOrdenPago As Integer
Private vPuntoVenta As Integer
Private vTipoAfip As Integer
Private vFacturaElectronica As Integer
Private vAutoID As Long
Private vCreated As Date

Private clon As clsDAOComprobante

Public ddl As New clsDDLComprobante

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vNegocioID = 0
    vComprobanteID = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vDescripcion = ""
    vTipoTransaccionID = 0
    vCuentaCorriente = 0
    vDebita = 0
    vIva = 0
    vAplicaPendiente = 0
    vAplicable = 0
    vAplicacion = 0
    vLibroIva = 0
    vLetraComprobante = ""
    vRecibo = 0
    vContado = 0
    vOrdenPago = 0
    vPuntoVenta = 0
    vTipoAfip = 0
    vFacturaElectronica = 0
    vAutoID = 0
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
' Description: Getter comprobanteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID() As Variant

    comprobanteID = vComprobanteID

End Property

Public Property Get key() As String

    key = "k." & comprobanteID
    
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
' Description: Getter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter tipoTransaccionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoTransaccionID() As Variant

    tipoTransaccionID = vTipoTransaccionID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter tipoTransaccionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tipoTransaccionID(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoTransaccionID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter aplicaPendiente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicaPendiente() As Variant

    aplicaPendiente = vAplicaPendiente

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter aplicaPendiente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicaPendiente(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicaPendiente = vNewValue

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
' Description: Getter debita
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get debita() As Variant

    debita = vDebita

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter debita
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let debita(ByVal vNewValue As Variant)

On Error Resume Next

    vDebita = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva() As Variant

    iva = vIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter aplicable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicable() As Variant

    aplicable = vAplicable

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter aplicable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicable(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicable = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter aplicacion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicacion() As Variant

    aplicacion = vAplicacion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter aplicacion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicacion(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicacion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter libroIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get libroIva() As Variant

    libroIva = vLibroIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter libroIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let libroIva(ByVal vNewValue As Variant)

On Error Resume Next

    vLibroIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter letraComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letraComprobante() As Variant

    letraComprobante = vLetraComprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter letraComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let letraComprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vLetraComprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter contado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contado() As Variant

    contado = vContado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter contado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let contado(ByVal vNewValue As Variant)

On Error Resume Next

    vContado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter ordenPago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ordenPago() As Variant

    ordenPago = vOrdenPago

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter ordenPago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ordenPago(ByVal vNewValue As Variant)

On Error Resume Next

    vOrdenPago = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter puntoVenta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get puntoVenta() As Variant

    puntoVenta = vPuntoVenta

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter puntoVenta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let puntoVenta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVenta = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter tipoAfip
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoAfip() As Variant

    tipoAfip = vTipoAfip

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter tipoAfip
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tipoAfip(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoAfip = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter facturaElectronica
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaElectronica() As Variant

    facturaElectronica = vFacturaElectronica

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter facturaElectronica
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let facturaElectronica(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaElectronica = vNewValue

End Property

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

    comboText = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vComprobanteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vComprobanteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOComprobante
Dim objObjeto As New clsDAOComprobante

    With objObjeto
        .negocioID = vNegocioID
        .comprobanteID = vComprobanteID
        .descripcion = vDescripcion
        .tipoTransaccionID = vTipoTransaccionID
        .cuentaCorriente = vCuentaCorriente
        .debita = vDebita
        .iva = vIva
        .aplicaPendiente = vAplicaPendiente
        .aplicable = vAplicable
        .aplicacion = vAplicacion
        .libroIva = vLibroIva
        .letraComprobante = vLetraComprobante
        .recibo = vRecibo
        .contado = vContado
        .ordenPago = vOrdenPago
        .puntoVenta = vPuntoVenta
        .tipoAfip = vTipoAfip
        .facturaElectronica = vFacturaElectronica
        .autoID = vAutoID
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
Private Sub fillObject(pObjeto As clsDAOComprobante, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .negocioID = pQuery.Fields(.ddl.negocioID)
        .comprobanteID = pQuery.Fields(.ddl.comprobanteID)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .tipoTransaccionID = pQuery.Fields(.ddl.tipoTransaccionID)
        .cuentaCorriente = pQuery.Fields(.ddl.cuentaCorriente)
        .debita = pQuery.Fields(.ddl.debita)
        .iva = pQuery.Fields(.ddl.iva)
        .aplicaPendiente = pQuery.Fields(.ddl.aplicaPendiente)
        .aplicable = pQuery.Fields(.ddl.aplicable)
        .aplicacion = pQuery.Fields(.ddl.aplicacion)
        .libroIva = pQuery.Fields(.ddl.libroIva)
        .letraComprobante = pQuery.Fields(.ddl.letraComprobante)
        .recibo = pQuery.Fields(.ddl.recibo)
        .contado = pQuery.Fields(.ddl.contado)
        .ordenPago = pQuery.Fields(.ddl.ordenPago)
        .puntoVenta = pQuery.Fields(.ddl.puntoVenta)
        .tipoAfip = pQuery.Fields(.ddl.tipoAfip)
        .facturaElectronica = pQuery.Fields(.ddl.facturaElectronica)
        .autoID = pQuery.Fields(.ddl.autoID)
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
    sql = sql & " WHERE " & ddl.negocioID & " = " & vNegocioID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID

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
' Description: Consulta la pDB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.negocioID & " = " & vNegocioID
    sql = sql & " ORDER BY " & ddl.comprobanteID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByAutoID(pDB As clsDB)
Dim sql As String

Dim lngAutoID As Long

    lngAutoID = vAutoID

    clean

    vAutoID = lngAutoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID
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
    sql = sql & " SET " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    sql = sql & ", " & ddl.tipoTransaccionID & " = " & vTipoTransaccionID
    sql = sql & ", " & ddl.cuentaCorriente & " = " & vCuentaCorriente
    sql = sql & ", " & ddl.debita & " = " & vDebita
    sql = sql & ", " & ddl.iva & " = " & vIva
    sql = sql & ", " & ddl.aplicaPendiente & " = " & vAplicaPendiente
    sql = sql & ", " & ddl.aplicable & " = " & vAplicable
    sql = sql & ", " & ddl.aplicacion & " = " & vAplicacion
    sql = sql & ", " & ddl.libroIva & " = " & vLibroIva
    sql = sql & ", " & ddl.letraComprobante & " = '" & replaceSpecialSymbols(vLetraComprobante) & "'"
    sql = sql & ", " & ddl.recibo & " = " & vRecibo
    sql = sql & ", " & ddl.contado & " = " & vContado
    sql = sql & ", " & ddl.ordenPago & " = " & vOrdenPago
    sql = sql & ", " & ddl.puntoVenta & " = " & vPuntoVenta
    sql = sql & ", " & ddl.tipoAfip & " = " & vTipoAfip
    sql = sql & ", " & ddl.facturaElectronica & " = " & vFacturaElectronica
    sql = sql & ", " & ddl.negocioID & " = " & vNegocioID
    sql = sql & ", " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & ";"

    add = pDB.execute(sql)

    findByPrimaryKey pDB

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

    If vDescripcion <> clon.descripcion Or IsNull(vDescripcion) <> IsNull(clon.descripcion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
        blnChange = True
    End If
    If vTipoTransaccionID <> clon.tipoTransaccionID Or IsNull(vTipoTransaccionID) <> IsNull(clon.tipoTransaccionID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipoTransaccionID & " = " & vTipoTransaccionID
        blnChange = True
    End If
    If vCuentaCorriente <> clon.cuentaCorriente Or IsNull(vCuentaCorriente) <> IsNull(clon.cuentaCorriente) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuentaCorriente & " = " & vCuentaCorriente
        blnChange = True
    End If
    If vDebita <> clon.debita Or IsNull(vDebita) <> IsNull(clon.debita) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.debita & " = " & vDebita
        blnChange = True
    End If
    If vIva <> clon.iva Or IsNull(vIva) <> IsNull(clon.iva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva & " = " & vIva
        blnChange = True
    End If
    If vAplicaPendiente <> clon.aplicaPendiente Or IsNull(vAplicaPendiente) <> IsNull(clon.aplicaPendiente) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicaPendiente & " = " & vAplicaPendiente
        blnChange = True
    End If
    If vAplicable <> clon.aplicable Or IsNull(vAplicable) <> IsNull(clon.aplicable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicable & " = " & vAplicable
        blnChange = True
    End If
    If vAplicacion <> clon.aplicacion Or IsNull(vAplicacion) <> IsNull(clon.aplicacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicacion & " = " & vAplicacion
        blnChange = True
    End If
    If vLibroIva <> clon.libroIva Or IsNull(vLibroIva) <> IsNull(clon.libroIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.libroIva & " = " & vLibroIva
        blnChange = True
    End If
    If vLetraComprobante <> clon.letraComprobante Or IsNull(vLetraComprobante) <> IsNull(clon.letraComprobante) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.letraComprobante & " = '" & replaceSpecialSymbols(vLetraComprobante) & "'"
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vContado <> clon.contado Or IsNull(vContado) <> IsNull(clon.contado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.contado & " = " & vContado
        blnChange = True
    End If
    If vOrdenPago <> clon.ordenPago Or IsNull(vOrdenPago) <> IsNull(clon.ordenPago) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ordenPago & " = " & vOrdenPago
        blnChange = True
    End If
    If vPuntoVenta <> clon.puntoVenta Or IsNull(vPuntoVenta) <> IsNull(clon.puntoVenta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntoVenta & " = " & vPuntoVenta
        blnChange = True
    End If
    If vTipoAfip <> clon.tipoAfip Or IsNull(vTipoAfip) <> IsNull(clon.tipoAfip) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipoAfip & " = " & vTipoAfip
        blnChange = True
    End If
    If vFacturaElectronica <> clon.facturaElectronica Or IsNull(vFacturaElectronica) <> IsNull(clon.facturaElectronica) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.facturaElectronica & " = " & vFacturaElectronica
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

    vComprobanteID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.negocioID & " = " & vNegocioID
    sql = sql & " ORDER BY " & ddl.comprobanteID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vComprobanteID = 1 + rstQuery.Fields(ddl.comprobanteID)

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

Public Function collectionByTipoTransaccionID(tipoTransaccionID As Integer, pDB As clsDB, Optional debita As Integer = 2, Optional letraComprobante As String = "") As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipoTransaccionID & " = " & tipoTransaccionID
    If debita < 2 Then sql = sql & " AND " & ddl.debita & " = " & debita
    If letraComprobante <> "" Then sql = sql & " AND " & ddl.letraComprobante & " = '" & letraComprobante & "'"
    sql = sql & ";"

    Set collectionByTipoTransaccionID = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la pDB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOComprobante

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOComprobante
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

Public Sub fillComboTipo(combo As ComboBox, tipoTransaccionID As Integer, pDB As clsDB, Optional debita As Integer = 2, Optional letraComprobante As String = "")

    fillComboCollection combo, collectionByTipoTransaccionID(tipoTransaccionID, pDB, debita, letraComprobante)
    
End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsDAOComprobante

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
Dim elemento As clsDAOComprobante

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

