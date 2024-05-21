VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOProveedorMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vProveedorID As Long
Private vComprobanteID As Integer
Private vPrefijo As Integer
Private vNroComprobante As Long
Private vEmpresaID As Integer
Private vNegocioID As Integer
Private vFechaComprobante As Date
Private vFechaVencimiento As Variant
Private vTotal As Currency
Private vTotalCancelado As Currency
Private vNeto As Currency
Private vImporteIva1 As Currency
Private vImporteIva2 As Currency
Private vImporteIva3 As Currency
Private vPercepcionIva As Currency
Private vPercepcionIIBB As Currency
Private vGastosNoGravados As Currency
Private vAjustes As Currency
Private vMonotributo As Integer
Private vCuentaMovimientoID As Currency
Private vObservaciones As String
Private vProveedormovimientoID As Long
Private vCreated As Date
' Atributos de memoria
Private vPago As Currency

Private clon As clsDAOProveedorMov

Public ddl As New clsDDLProveedorMov

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vProveedorID = 0
    vComprobanteID = 0
    vPrefijo = 0
    vNroComprobante = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vEmpresaID = 0
    vNegocioID = 0
    vFechaComprobante = Date
    vFechaVencimiento = Null
    vTotal = 0
    vTotalCancelado = 0
    vNeto = 0
    vImporteIva1 = 0
    vImporteIva2 = 0
    vImporteIva3 = 0
    vPercepcionIva = 0
    vPercepcionIIBB = 0
    vGastosNoGravados = 0
    vAjustes = 0
    vMonotributo = 0
    vCuentaMovimientoID = 0
    vObservaciones = ""
    vProveedormovimientoID = 0
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
' Description: Getter prefijo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijo() As Variant

    prefijo = vPrefijo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter prefijo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let prefijo(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijo = vNewValue

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
' Description: Getter empresaID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get empresaID() As Variant

    empresaID = vEmpresaID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter empresaID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let empresaID(ByVal vNewValue As Variant)

On Error Resume Next

    vEmpresaID = vNewValue

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
' Description: Getter fechaComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaComprobante() As Variant

    fechaComprobante = vFechaComprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter fechaComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaComprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaComprobante = vNewValue

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
' Description: Getter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total() As Variant

    total = vTotal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter totalCancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get totalCancelado() As Variant

    totalCancelado = vTotalCancelado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter totalCancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let totalCancelado(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalCancelado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter neto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto() As Variant

    neto = vNeto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter neto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter importeIva1
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva1() As Variant

    importeIva1 = vImporteIva1

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter importeIva1
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importeIva1(ByVal vNewValue As Variant)

On Error Resume Next

    vImporteIva1 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter importeIva2
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva2() As Variant

    importeIva2 = vImporteIva2

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter importeIva2
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importeIva2(ByVal vNewValue As Variant)

On Error Resume Next

    vImporteIva2 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter importeIva3
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva3() As Variant

    importeIva3 = vImporteIva3

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter importeIva3
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importeIva3(ByVal vNewValue As Variant)

On Error Resume Next

    vImporteIva3 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter percepcionIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIva() As Variant

    percepcionIva = vPercepcionIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter percepcionIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let percepcionIva(ByVal vNewValue As Variant)

On Error Resume Next

    vPercepcionIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter percepcionIIBB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIIBB() As Variant

    percepcionIIBB = vPercepcionIIBB

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter percepcionIIBB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let percepcionIIBB(ByVal vNewValue As Variant)

On Error Resume Next

    vPercepcionIIBB = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter gastosNoGravados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get gastosNoGravados() As Variant

    gastosNoGravados = vGastosNoGravados

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter gastosNoGravados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let gastosNoGravados(ByVal vNewValue As Variant)

On Error Resume Next

    vGastosNoGravados = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter ajustes
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ajustes() As Variant

    ajustes = vAjustes

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter ajustes
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ajustes(ByVal vNewValue As Variant)

On Error Resume Next

    vAjustes = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter monotributo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get monotributo() As Variant

    monotributo = vMonotributo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter monotributo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let monotributo(ByVal vNewValue As Variant)

On Error Resume Next

    vMonotributo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter cuentaMovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaMovimientoID() As Variant

    cuentaMovimientoID = vCuentaMovimientoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter cuentaMovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cuentaMovimientoID(ByVal vNewValue As Variant)

On Error Resume Next

    vCuentaMovimientoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones() As Variant

    observaciones = vObservaciones

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let observaciones(ByVal vNewValue As Variant)

On Error Resume Next

    vObservaciones = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter proveedorMovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedormovimientoID() As Variant

    proveedormovimientoID = vProveedormovimientoID

End Property

Public Property Get key() As String

    key = "k." & proveedormovimientoID
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedorMovimientoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedormovimientoID(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedormovimientoID = vNewValue

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

Public Property Get pago() As Variant

    pago = vPago

End Property

Public Property Let pago(ByVal vNewValue As Variant)

On Error Resume Next

    vPago = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vEmpresaID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vNroComprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vEmpresaID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vNroComprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOProveedorMov
Dim objObjeto As New clsDAOProveedorMov

    With objObjeto
        .proveedorID = vProveedorID
        .comprobanteID = vComprobanteID
        .prefijo = vPrefijo
        .nroComprobante = vNroComprobante
        .empresaID = vEmpresaID
        .negocioID = vNegocioID
        .fechaComprobante = vFechaComprobante
        .fechaVencimiento = vFechaVencimiento
        .total = vTotal
        .totalCancelado = vTotalCancelado
        .neto = vNeto
        .importeIva1 = vImporteIva1
        .importeIva2 = vImporteIva2
        .importeIva3 = vImporteIva3
        .percepcionIva = vPercepcionIva
        .percepcionIIBB = vPercepcionIIBB
        .gastosNoGravados = vGastosNoGravados
        .ajustes = vAjustes
        .monotributo = vMonotributo
        .cuentaMovimientoID = vCuentaMovimientoID
        .observaciones = vObservaciones
        .proveedormovimientoID = vProveedormovimientoID
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
Private Sub fillObject(pObjeto As clsDAOProveedorMov, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .proveedorID = pQuery.Fields(.ddl.proveedorID)
        .comprobanteID = pQuery.Fields(.ddl.comprobanteID)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .nroComprobante = pQuery.Fields(.ddl.nroComprobante)
        .empresaID = pQuery.Fields(.ddl.empresaID)
        .negocioID = pQuery.Fields(.ddl.negocioID)
        .fechaComprobante = pQuery.Fields(.ddl.fechaComprobante)
        .fechaVencimiento = pQuery.Fields(.ddl.fechaVencimiento)
        .total = pQuery.Fields(.ddl.total)
        .totalCancelado = pQuery.Fields(.ddl.totalCancelado)
        .neto = pQuery.Fields(.ddl.neto)
        .importeIva1 = pQuery.Fields(.ddl.importeIva1)
        .importeIva2 = pQuery.Fields(.ddl.importeIva2)
        .importeIva3 = pQuery.Fields(.ddl.importeIva3)
        .percepcionIva = pQuery.Fields(.ddl.percepcionIva)
        .percepcionIIBB = pQuery.Fields(.ddl.percepcionIIBB)
        .gastosNoGravados = pQuery.Fields(.ddl.gastosNoGravados)
        .ajustes = pQuery.Fields(.ddl.ajustes)
        .monotributo = pQuery.Fields(.ddl.monotributo)
        .cuentaMovimientoID = pQuery.Fields(.ddl.cuentaMovimientoID)
        .observaciones = pQuery.Fields(.ddl.observaciones)
        .proveedormovimientoID = pQuery.Fields(.ddl.proveedormovimientoID)
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
    sql = sql & " WHERE " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " AND " & ddl.nroComprobante & " = " & vNroComprobante

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
    sql = sql & " WHERE " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.nroComprobante & " DESC"
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
Public Sub findByProveedorMovimientoID(pDB As clsDB)
Dim sql As String

Dim lngProveedorMovimientoID As Long

    lngProveedorMovimientoID = vProveedormovimientoID

    clean

    vProveedormovimientoID = lngProveedorMovimientoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedormovimientoID & " = " & vProveedormovimientoID
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Sub findByComprobante(pDB As clsDB)
Dim sql As String

Dim nroComprobante As Long

Dim prefijo As Integer
Dim proveedorID As Integer
Dim comprobanteID As Integer


    proveedorID = vProveedorID
    comprobanteID = vComprobanteID
    nroComprobante = vNroComprobante
    prefijo = vPrefijo

    clean

    vProveedorID = proveedorID
    vComprobanteID = comprobanteID
    vPrefijo = prefijo
    vNroComprobante = nroComprobante

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " AND " & ddl.nroComprobante & " = " & vNroComprobante
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Sub findLastOPago(pDB As clsDB)
Dim sql As String

Dim c As New clsDDLComprobante
    
    clean

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & c.comprobanteID_t & " = " & ddl.comprobanteID_t
    sql = sql & " AND " & c.ordenPago_t & " <> 0"
    sql = sql & " AND " & c.tipoTransaccionID_t & " = " & cntTTPagos
    sql = sql & " ORDER BY " & ddl.nroComprobante_t & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"
    
    findSQL sql, pDB

End Sub

Public Sub findLastByComprobante(pDB As clsDB)
Dim sql As String

Dim comprobanteID As Long

    comprobanteID = vComprobanteID
    
    clean
    
    vComprobanteID = comprobanteID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.prefijo & " = " & vComprobanteID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & " ORDER BY " & ddl.nroComprobante & " DESC"
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
    sql = sql & " SET " & ddl.empresaID & " = " & vEmpresaID
    sql = sql & ", " & ddl.negocioID & " = " & vNegocioID
    sql = sql & ", " & ddl.fechaComprobante & " = " & fechaDB(vFechaComprobante)
    sql = sql & ", " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
    sql = sql & ", " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
    sql = sql & ", " & ddl.totalCancelado & " = '" & Replace(vTotalCancelado, ",", ".") & "'"
    sql = sql & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    sql = sql & ", " & ddl.importeIva1 & " = '" & Replace(vImporteIva1, ",", ".") & "'"
    sql = sql & ", " & ddl.importeIva2 & " = '" & Replace(vImporteIva2, ",", ".") & "'"
    sql = sql & ", " & ddl.importeIva3 & " = '" & Replace(vImporteIva3, ",", ".") & "'"
    sql = sql & ", " & ddl.percepcionIva & " = '" & Replace(vPercepcionIva, ",", ".") & "'"
    sql = sql & ", " & ddl.percepcionIIBB & " = '" & Replace(vPercepcionIIBB, ",", ".") & "'"
    sql = sql & ", " & ddl.gastosNoGravados & " = '" & Replace(vGastosNoGravados, ",", ".") & "'"
    sql = sql & ", " & ddl.ajustes & " = '" & Replace(vAjustes, ",", ".") & "'"
    sql = sql & ", " & ddl.monotributo & " = " & vMonotributo
    sql = sql & ", " & ddl.cuentaMovimientoID & " = '" & Replace(vCuentaMovimientoID, ",", ".") & "'"
    sql = sql & ", " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
    sql = sql & ", " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & ", " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & ", " & ddl.prefijo & " = " & vPrefijo
    sql = sql & ", " & ddl.nroComprobante & " = " & vNroComprobante
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

    If vEmpresaID <> clon.empresaID Or IsNull(vEmpresaID) <> IsNull(clon.empresaID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.empresaID & " = " & vEmpresaID
        blnChange = True
    End If
    If vNegocioID <> clon.negocioID Or IsNull(vNegocioID) <> IsNull(clon.negocioID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negocioID & " = " & vNegocioID
        blnChange = True
    End If
    If vFechaComprobante <> clon.fechaComprobante Or IsNull(vFechaComprobante) <> IsNull(clon.fechaComprobante) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaComprobante & " = " & fechaDB(vFechaComprobante)
        blnChange = True
    End If
    If vFechaVencimiento <> clon.fechaVencimiento Or IsNull(vFechaVencimiento) <> IsNull(clon.fechaVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
        blnChange = True
    End If
    If vTotal <> clon.total Or IsNull(vTotal) <> IsNull(clon.total) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
        blnChange = True
    End If
    If vTotalCancelado <> clon.totalCancelado Or IsNull(vTotalCancelado) <> IsNull(clon.totalCancelado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.totalCancelado & " = '" & Replace(vTotalCancelado, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto <> clon.neto Or IsNull(vNeto) <> IsNull(clon.neto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
        blnChange = True
    End If
    If vImporteIva1 <> clon.importeIva1 Or IsNull(vImporteIva1) <> IsNull(clon.importeIva1) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importeIva1 & " = '" & Replace(vImporteIva1, ",", ".") & "'"
        blnChange = True
    End If
    If vImporteIva2 <> clon.importeIva2 Or IsNull(vImporteIva2) <> IsNull(clon.importeIva2) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importeIva2 & " = '" & Replace(vImporteIva2, ",", ".") & "'"
        blnChange = True
    End If
    If vImporteIva3 <> clon.importeIva3 Or IsNull(vImporteIva3) <> IsNull(clon.importeIva3) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importeIva3 & " = '" & Replace(vImporteIva3, ",", ".") & "'"
        blnChange = True
    End If
    If vPercepcionIva <> clon.percepcionIva Or IsNull(vPercepcionIva) <> IsNull(clon.percepcionIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.percepcionIva & " = '" & Replace(vPercepcionIva, ",", ".") & "'"
        blnChange = True
    End If
    If vPercepcionIIBB <> clon.percepcionIIBB Or IsNull(vPercepcionIIBB) <> IsNull(clon.percepcionIIBB) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.percepcionIIBB & " = '" & Replace(vPercepcionIIBB, ",", ".") & "'"
        blnChange = True
    End If
    If vGastosNoGravados <> clon.gastosNoGravados Or IsNull(vGastosNoGravados) <> IsNull(clon.gastosNoGravados) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.gastosNoGravados & " = '" & Replace(vGastosNoGravados, ",", ".") & "'"
        blnChange = True
    End If
    If vAjustes <> clon.ajustes Or IsNull(vAjustes) <> IsNull(clon.ajustes) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ajustes & " = '" & Replace(vAjustes, ",", ".") & "'"
        blnChange = True
    End If
    If vMonotributo <> clon.monotributo Or IsNull(vMonotributo) <> IsNull(clon.monotributo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.monotributo & " = " & vMonotributo
        blnChange = True
    End If
    If vCuentaMovimientoID <> clon.cuentaMovimientoID Or IsNull(vCuentaMovimientoID) <> IsNull(clon.cuentaMovimientoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuentaMovimientoID & " = '" & Replace(vCuentaMovimientoID, ",", ".") & "'"
        blnChange = True
    End If
    If vObservaciones <> clon.observaciones Or IsNull(vObservaciones) <> IsNull(clon.observaciones) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
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

    vNroComprobante = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & " AND " & ddl.comprobanteID & " = " & vComprobanteID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.nroComprobante & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vNroComprobante = 1 + rstQuery.Fields(ddl.nroComprobante)

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

Public Function collectionPendientesByProveedorID(proveedorID As Long, pDB As clsDB) As Collection
Dim sql As String

Dim c As New clsDDLComprobante

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.comprobanteID_t & " = " & c.comprobanteID_t
    sql = sql & " WHERE " & ddl.proveedorID_t & " = " & proveedorID
    sql = sql & " AND " & c.cuentaCorriente_t & " <> 0"
    sql = sql & " AND ABS(" & ddl.total_t & ") > ABS(" & ddl.totalCancelado_t & ")"
    sql = sql & " ORDER BY " & ddl.fechaComprobante_t
    sql = sql & ";"

    Set collectionPendientesByProveedorID = collectionSQL(sql, pDB)

End Function

Public Function collectionCtaCteByProveedorID(proveedorID As Long, pDB As clsDB, Optional tope As Date = #1/1/2100#) As Collection
Dim sql As String

Dim c As New clsDDLComprobante

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.comprobanteID_t & " = " & c.comprobanteID_t
    sql = sql & " WHERE " & ddl.proveedorID_t & " = " & proveedorID
    sql = sql & " AND " & ddl.fechaComprobante_t & " < " & fechaDB(tope)
    sql = sql & " AND " & c.cuentaCorriente_t & " = 1"
    sql = sql & " ORDER BY " & ddl.fechaComprobante_t & ", " & c.ordenPago_t & " DESC"
    sql = sql & ";"

    Set collectionCtaCteByProveedorID = collectionSQL(sql, pDB)

End Function

Public Function collectionPeriodoByProveedorID(proveedorID As Long, desde As Date, hasta As Date, pDB As clsDB) As Collection
Dim sql As String

Dim c As New clsDDLComprobante

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.comprobanteID_t & " = " & c.comprobanteID_t
    sql = sql & " WHERE " & ddl.proveedorID_t & " = " & proveedorID
    sql = sql & " AND " & ddl.fechaComprobante_t & " BETWEEN " & fechaDB(desde) & " AND " & fechaDB(hasta)
    sql = sql & " ORDER BY " & ddl.fechaComprobante_t & ", " & c.ordenPago_t & " DESC"
    sql = sql & ";"

    Set collectionPeriodoByProveedorID = collectionSQL(sql, pDB)

End Function

Public Function collectionByProveedorMovimientoIDDeuda(proveedormovimientoID As Long, pDB As clsDB) As Collection
Dim sql As String

Dim p As New clsDDLProveedorPago

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & p.table
    sql = sql & " ON " & ddl.proveedormovimientoID_t & " = " & p.proveedorMovimientoIDAplicado_t
    sql = sql & " WHERE " & p.proveedorMovimientoIDDeuda_t & " = " & proveedormovimientoID
    sql = sql & " ORDER BY " & ddl.fechaComprobante_t & ", " & ddl.prefijo_t & ", " & ddl.nroComprobante_t
    sql = sql & ";"

    Set collectionByProveedorMovimientoIDDeuda = collectionSQL(sql, pDB)

End Function

Public Function collectionByProveedorMovimientoIDPago(proveedormovimientoID As Long, pDB As clsDB) As Collection
Dim sql As String

Dim p As New clsDDLProveedorPago

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & p.table
    sql = sql & " ON " & ddl.proveedormovimientoID_t & " = " & p.proveedorMovimientoIDDeuda_t
    sql = sql & " WHERE " & p.proveedorMovimientoIDAplicado_t & " = " & proveedormovimientoID
    sql = sql & " ORDER BY " & ddl.fechaComprobante_t & ", " & ddl.prefijo_t & ", " & ddl.nroComprobante_t
    sql = sql & ";"

    Set collectionByProveedorMovimientoIDPago = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la pDB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOProveedorMov

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOProveedorMov
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
Dim elemento As clsDAOProveedorMov

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
Dim elemento As clsDAOProveedorMov

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

