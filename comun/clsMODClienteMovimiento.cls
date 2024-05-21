VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODClienteMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClientemovimientoId As Variant
Private vNegocioId As Variant
Private vEmpresaId As Variant
Private vClienteId As Variant
Private vComprobanteId As Variant
Private vFechacomprobante As Variant
Private vFechavencimiento As Variant
Private vPuntoventa As Integer
Private vNumerocomprobante As Long
Private vImporte As Currency
Private vCancelado As Currency
Private vNetosindescuento As Currency
Private vNeto As Currency
Private vNetocancelado As Currency
Private vIva As Currency
Private vExento As Currency
Private vRecibo As Integer
Private vAnulada As Integer
Private vTipocomprobante As String
Private vLetras As String
Private vAlicuotaiva As Currency
Private vObservaciones As String
Private vCae As String
Private vCaevencimiento As String
Private vBarras As String

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClientemovimientoId = Null
    vNegocioId = Null
    vEmpresaId = Null
    vClienteId = Null
    vComprobanteId = Null
    vFechacomprobante = Null
    vFechavencimiento = Null
    vPuntoventa = 0
    vNumerocomprobante = 0
    vImporte = 0
    vCancelado = 0
    vNetosindescuento = 0
    vNeto = 0
    vNetocancelado = 0
    vIva = 0
    vExento = 0
    vRecibo = 0
    vAnulada = 0
    vTipocomprobante = ""
    vLetras = ""
    vAlicuotaiva = 0
    vObservaciones = ""
    vCae = ""
    vCaevencimiento = ""
    vBarras = ""

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter clientemovimientoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clientemovimientoId() As Variant

    clientemovimientoId = vClientemovimientoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter clientemovimientoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clientemovimientoId(ByVal vNewValue As Variant)

On Error Resume Next

    vClientemovimientoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter negocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioId() As Variant

    negocioId = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter negocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let negocioId(ByVal vNewValue As Variant)

On Error Resume Next

    vNegocioId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter empresaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get empresaId() As Variant

    empresaId = vEmpresaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter empresaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let empresaId(ByVal vNewValue As Variant)

On Error Resume Next

    vEmpresaId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter comprobanteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteId() As Variant

    comprobanteId = vComprobanteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter comprobanteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let comprobanteId(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter fechacomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechacomprobante() As Variant

    fechacomprobante = vFechacomprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter fechacomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechacomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vFechacomprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter fechavencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechavencimiento() As Variant

    fechavencimiento = vFechavencimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter fechavencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechavencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechavencimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter puntoventa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get puntoventa() As Variant

    puntoventa = vPuntoventa

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter puntoventa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let puntoventa(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoventa = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter numerocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerocomprobante() As Variant

    numerocomprobante = vNumerocomprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter numerocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let numerocomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerocomprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter importe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importe() As Variant

    importe = vImporte

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter importe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter cancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cancelado() As Variant

    cancelado = vCancelado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter cancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cancelado(ByVal vNewValue As Variant)

On Error Resume Next

    vCancelado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter netosindescuento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get netosindescuento() As Variant

    netosindescuento = vNetosindescuento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter netosindescuento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let netosindescuento(ByVal vNewValue As Variant)

On Error Resume Next

    vNetosindescuento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter neto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto() As Variant

    neto = vNeto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter neto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter netocancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get netocancelado() As Variant

    netocancelado = vNetocancelado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter netocancelado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let netocancelado(ByVal vNewValue As Variant)

On Error Resume Next

    vNetocancelado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva() As Variant

    iva = vIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter exento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get exento() As Variant

    exento = vExento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter exento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter tipocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipocomprobante() As Variant

    tipocomprobante = vTipocomprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter tipocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tipocomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vTipocomprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter letras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letras() As Variant

    letras = vLetras

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter letras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let letras(ByVal vNewValue As Variant)

On Error Resume Next

    vLetras = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter alicuotaiva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get alicuotaiva() As Variant

    alicuotaiva = vAlicuotaiva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter alicuotaiva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let alicuotaiva(ByVal vNewValue As Variant)

On Error Resume Next

    vAlicuotaiva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones() As Variant

    observaciones = vObservaciones

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let observaciones(ByVal vNewValue As Variant)

On Error Resume Next

    vObservaciones = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter cae
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cae() As Variant

    cae = vCae

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter cae
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter caevencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get caevencimiento() As Variant

    caevencimiento = vCaevencimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter caevencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let caevencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCaevencimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter barras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get barras() As Variant

    barras = vBarras

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Setter barras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let barras(ByVal vNewValue As Variant)

On Error Resume Next

    vBarras = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & clientemovimientoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vClientemovimientoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vClientemovimientoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODClienteMovimiento
Dim object As New clsMODClienteMovimiento

    object.clientemovimientoId = vClientemovimientoId
    object.negocioId = vNegocioId
    object.empresaId = vEmpresaId
    object.clienteId = vClienteId
    object.comprobanteId = vComprobanteId
    object.fechacomprobante = vFechacomprobante
    object.fechavencimiento = vFechavencimiento
    object.puntoventa = vPuntoventa
    object.numerocomprobante = vNumerocomprobante
    object.importe = vImporte
    object.cancelado = vCancelado
    object.netosindescuento = vNetosindescuento
    object.neto = vNeto
    object.netocancelado = vNetocancelado
    object.iva = vIva
    object.exento = vExento
    object.recibo = vRecibo
    object.anulada = vAnulada
    object.tipocomprobante = vTipocomprobante
    object.letras = vLetras
    object.alicuotaiva = vAlicuotaiva
    object.observaciones = vObservaciones
    object.cae = vCae
    object.caevencimiento = vCaevencimiento
    object.barras = vBarras

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "clientemovimientoId", clientemovimientoId
    object.add "negocioId", negocioId
    object.add "empresaId", empresaId
    object.add "clienteId", clienteId
    object.add "comprobanteId", comprobanteId
    object.add "fechacomprobante", modConv.date2JSON(fechacomprobante)
    object.add "fechavencimiento", modConv.date2JSON(fechavencimiento)
    object.add "puntoventa", puntoventa
    object.add "numerocomprobante", numerocomprobante
    object.add "importe", importe
    object.add "cancelado", cancelado
    object.add "netosindescuento", netosindescuento
    object.add "neto", neto
    object.add "netocancelado", netocancelado
    object.add "iva", iva
    object.add "exento", exento
    object.add "recibo", recibo
    object.add "anulada", anulada
    object.add "tipocomprobante", tipocomprobante
    object.add "letras", letras
    object.add "alicuotaiva", alicuotaiva
    object.add "observaciones", observaciones
    object.add "cae", cae
    object.add "caevencimiento", caevencimiento
    object.add "barras", barras

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    clientemovimientoId = json_object.item("clientemovimientoId")
    negocioId = json_object.item("negocioId")
    empresaId = json_object.item("empresaId")
    clienteId = json_object.item("clienteId")
    comprobanteId = json_object.item("comprobanteId")
    fechacomprobante = modConv.JSON2Date(json_object.item("fechacomprobante"))
    fechavencimiento = modConv.JSON2Date(json_object.item("fechavencimiento"))
    puntoventa = json_object.item("puntoventa")
    numerocomprobante = json_object.item("numerocomprobante")
    importe = json_object.item("importe")
    cancelado = json_object.item("cancelado")
    netosindescuento = json_object.item("netosindescuento")
    neto = json_object.item("neto")
    netocancelado = json_object.item("netocancelado")
    iva = json_object.item("iva")
    exento = json_object.item("exento")
    recibo = json_object.item("recibo")
    anulada = json_object.item("anulada")
    tipocomprobante = json_object.item("tipocomprobante")
    letras = json_object.item("letras")
    alicuotaiva = json_object.item("alicuotaiva")
    observaciones = json_object.item("observaciones")
    cae = json_object.item("cae")
    caevencimiento = json_object.item("caevencimiento")
    barras = json_object.item("barras")

End Sub

