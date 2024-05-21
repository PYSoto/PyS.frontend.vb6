VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODElectronico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vElectronicoId As Variant
Private vComprobanteId As Variant
Private vPuntoventa As Integer
Private vNumerocomprobante As Long
Private vClienteId As Variant
Private vTotal As Currency
Private vExento As Currency
Private vNeto21 As Currency
Private vNeto105 As Currency
Private vIva21 As Currency
Private vIva105 As Currency
Private vCae As String
Private vFecha As String
Private vCaevencimiento As String
Private vCodigobarras As String
Private vTipodocumento As Variant
Private vNumerodocumento As Variant
Private vClientemovimientoIdasociado As Variant

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vElectronicoId = Null
    vComprobanteId = Null
    vPuntoventa = 0
    vNumerocomprobante = 0
    vClienteId = Null
    vTotal = 0
    vExento = 0
    vNeto21 = 0
    vNeto105 = 0
    vIva21 = 0
    vIva105 = 0
    vCae = ""
    vFecha = ""
    vCaevencimiento = ""
    vCodigobarras = ""
    vTipodocumento = Null
    vNumerodocumento = Null
    vClientemovimientoIdasociado = Null

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter electronicoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get electronicoId() As Variant

    electronicoId = vElectronicoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter electronicoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let electronicoId(ByVal vNewValue As Variant)

On Error Resume Next

    vElectronicoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter comprobanteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteId() As Variant

    comprobanteId = vComprobanteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter comprobanteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let comprobanteId(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter puntoventa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get puntoventa() As Variant

    puntoventa = vPuntoventa

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter puntoventa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let puntoventa(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoventa = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter numerocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerocomprobante() As Variant

    numerocomprobante = vNumerocomprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter numerocomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let numerocomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerocomprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total() As Variant

    total = vTotal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter exento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get exento() As Variant

    exento = vExento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter exento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter neto21
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto21() As Variant

    neto21 = vNeto21

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter neto21
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let neto21(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto21 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter neto105
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto105() As Variant

    neto105 = vNeto105

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter neto105
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let neto105(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto105 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter iva21
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva21() As Variant

    iva21 = vIva21

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter iva21
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let iva21(ByVal vNewValue As Variant)

On Error Resume Next

    vIva21 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter iva105
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva105() As Variant

    iva105 = vIva105

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter iva105
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let iva105(ByVal vNewValue As Variant)

On Error Resume Next

    vIva105 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter cae
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cae() As Variant

    cae = vCae

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter cae
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As Variant

    fecha = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter caevencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get caevencimiento() As Variant

    caevencimiento = vCaevencimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter caevencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let caevencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCaevencimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter codigobarras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get codigobarras() As Variant

    codigobarras = vCodigobarras

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter codigobarras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let codigobarras(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigobarras = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter tipodocumento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipodocumento() As Variant

    tipodocumento = vTipodocumento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter tipodocumento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tipodocumento(ByVal vNewValue As Variant)

On Error Resume Next

    vTipodocumento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter numerodocumento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerodocumento() As Variant

    numerodocumento = vNumerodocumento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter numerodocumento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let numerodocumento(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerodocumento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter clientemovimientoIdasociado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clientemovimientoIdasociado() As Variant

    clientemovimientoIdasociado = vClientemovimientoIdasociado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter clientemovimientoIdasociado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clientemovimientoIdasociado(ByVal vNewValue As Variant)

On Error Resume Next

    vClientemovimientoIdasociado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & electronicoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vComprobanteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vElectronicoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vComprobanteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vElectronicoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODElectronico
Dim object As New clsMODElectronico

    object.electronicoId = vElectronicoId
    object.comprobanteId = vComprobanteId
    object.puntoventa = vPuntoventa
    object.numerocomprobante = vNumerocomprobante
    object.clienteId = vClienteId
    object.total = vTotal
    object.exento = vExento
    object.neto21 = vNeto21
    object.neto105 = vNeto105
    object.iva21 = vIva21
    object.iva105 = vIva105
    object.cae = vCae
    object.fecha = vFecha
    object.caevencimiento = vCaevencimiento
    object.codigobarras = vCodigobarras
    object.tipodocumento = vTipodocumento
    object.numerodocumento = vNumerodocumento
    object.clientemovimientoIdasociado = vClientemovimientoIdasociado

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "electronicoId", electronicoId
    object.add "comprobanteId", comprobanteId
    object.add "puntoventa", puntoventa
    object.add "numerocomprobante", numerocomprobante
    object.add "clienteId", clienteId
    object.add "total", total
    object.add "exento", exento
    object.add "neto21", neto21
    object.add "neto105", neto105
    object.add "iva21", iva21
    object.add "iva105", iva105
    object.add "cae", cae
    object.add "fecha", fecha
    object.add "caevencimiento", caevencimiento
    object.add "codigobarras", codigobarras
    object.add "tipodocumento", tipodocumento
    object.add "numerodocumento", numerodocumento
    object.add "clientemovimientoIdasociado", clientemovimientoIdasociado

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    electronicoId = json_object.item("electronicoId")
    comprobanteId = json_object.item("comprobanteId")
    puntoventa = json_object.item("puntoventa")
    numerocomprobante = json_object.item("numerocomprobante")
    clienteId = json_object.item("clienteId")
    total = json_object.item("total")
    exento = json_object.item("exento")
    neto21 = json_object.item("neto21")
    neto105 = json_object.item("neto105")
    iva21 = json_object.item("iva21")
    iva105 = json_object.item("iva105")
    cae = json_object.item("cae")
    fecha = json_object.item("fecha")
    caevencimiento = json_object.item("caevencimiento")
    codigobarras = json_object.item("codigobarras")
    tipodocumento = json_object.item("tipodocumento")
    numerodocumento = json_object.item("numerodocumento")
    clientemovimientoIdasociado = json_object.item("clientemovimientoIdasociado")

End Sub

