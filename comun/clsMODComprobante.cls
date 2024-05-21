VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vUniqueId As Variant
Private vNegocioId As Variant
Private vComprobanteId As Variant
Private vDescripcion As String
Private vModulo As Integer
Private vCuentacorriente As Integer
Private vDebita As Integer
Private vIva As Integer
Private vAplicapendiente As Integer
Private vAplicable As Integer
Private vAplicacion As Integer
Private vLibroiva As Integer
Private vLetracomprobante As String
Private vRecibo As Integer
Private vContado As Integer
Private vOrdenpago As Integer
Private vPuntoventa As Integer
Private vComprobanteafipId As Variant
Private vElectronico As Integer
Private vAsociado As Integer

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vUniqueId = Null
    vNegocioId = Null
    vComprobanteId = Null
    vDescripcion = ""
    vModulo = 0
    vCuentacorriente = 0
    vDebita = 0
    vIva = 0
    vAplicapendiente = 0
    vAplicable = 0
    vAplicacion = 0
    vLibroiva = 0
    vLetracomprobante = ""
    vRecibo = 0
    vContado = 0
    vOrdenpago = 0
    vPuntoventa = 0
    vComprobanteafipId = Null
    vElectronico = 0
    vAsociado = 0

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
' Description: Getter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uniqueId() As Variant

    uniqueId = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uniqueId(ByVal vNewValue As Variant)

On Error Resume Next

    vUniqueId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter negocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioId() As Variant

    negocioId = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter negocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let negocioId(ByVal vNewValue As Variant)

On Error Resume Next

    vNegocioId = vNewValue

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
' Description: Getter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter modulo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get modulo() As Variant

    modulo = vModulo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter modulo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let modulo(ByVal vNewValue As Variant)

On Error Resume Next

    vModulo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter cuentacorriente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentacorriente() As Variant

    cuentacorriente = vCuentacorriente

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter cuentacorriente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cuentacorriente(ByVal vNewValue As Variant)

On Error Resume Next

    vCuentacorriente = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter debita
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get debita() As Variant

    debita = vDebita

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter debita
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let debita(ByVal vNewValue As Variant)

On Error Resume Next

    vDebita = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva() As Variant

    iva = vIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter aplicapendiente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicapendiente() As Variant

    aplicapendiente = vAplicapendiente

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter aplicapendiente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicapendiente(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicapendiente = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter aplicable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicable() As Variant

    aplicable = vAplicable

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter aplicable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicable(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicable = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter aplicacion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicacion() As Variant

    aplicacion = vAplicacion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter aplicacion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let aplicacion(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicacion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter libroiva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get libroiva() As Variant

    libroiva = vLibroiva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter libroiva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let libroiva(ByVal vNewValue As Variant)

On Error Resume Next

    vLibroiva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter letracomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letracomprobante() As Variant

    letracomprobante = vLetracomprobante

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter letracomprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let letracomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vLetracomprobante = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter contado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contado() As Variant

    contado = vContado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter contado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let contado(ByVal vNewValue As Variant)

On Error Resume Next

    vContado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter ordenpago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ordenpago() As Variant

    ordenpago = vOrdenpago

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter ordenpago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ordenpago(ByVal vNewValue As Variant)

On Error Resume Next

    vOrdenpago = vNewValue

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
' Description: Getter comprobanteafipId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteafipId() As Variant

    comprobanteafipId = vComprobanteafipId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter comprobanteafipId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let comprobanteafipId(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteafipId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter electronico
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get electronico() As Variant

    electronico = vElectronico

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter electronico
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let electronico(ByVal vNewValue As Variant)

On Error Resume Next

    vElectronico = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter asociado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get asociado() As Variant

    asociado = vAsociado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Setter asociado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let asociado(ByVal vNewValue As Variant)

On Error Resume Next

    vAsociado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & comprobanteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vNegocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODComprobante
Dim object As New clsMODComprobante

    object.uniqueId = vUniqueId
    object.negocioId = vNegocioId
    object.comprobanteId = vComprobanteId
    object.descripcion = vDescripcion
    object.modulo = vModulo
    object.cuentacorriente = vCuentacorriente
    object.debita = vDebita
    object.iva = vIva
    object.aplicapendiente = vAplicapendiente
    object.aplicable = vAplicable
    object.aplicacion = vAplicacion
    object.libroiva = vLibroiva
    object.letracomprobante = vLetracomprobante
    object.recibo = vRecibo
    object.contado = vContado
    object.ordenpago = vOrdenpago
    object.puntoventa = vPuntoventa
    object.comprobanteafipId = vComprobanteafipId
    object.electronico = vElectronico
    object.asociado = vAsociado

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

    object.add "uniqueId", uniqueId
    object.add "negocioId", negocioId
    object.add "comprobanteId", comprobanteId
    object.add "descripcion", descripcion
    object.add "modulo", modulo
    object.add "cuentacorriente", cuentacorriente
    object.add "debita", debita
    object.add "iva", iva
    object.add "aplicapendiente", aplicapendiente
    object.add "aplicable", aplicable
    object.add "aplicacion", aplicacion
    object.add "libroiva", libroiva
    object.add "letracomprobante", letracomprobante
    object.add "recibo", recibo
    object.add "contado", contado
    object.add "ordenpago", ordenpago
    object.add "puntoventa", puntoventa
    object.add "comprobanteafipId", comprobanteafipId
    object.add "electronico", electronico
    object.add "asociado", asociado

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

    uniqueId = json_object.item("uniqueId")
    negocioId = json_object.item("negocioId")
    comprobanteId = json_object.item("comprobanteId")
    descripcion = json_object.item("descripcion")
    modulo = json_object.item("modulo")
    cuentacorriente = json_object.item("cuentacorriente")
    debita = json_object.item("debita")
    iva = json_object.item("iva")
    aplicapendiente = json_object.item("aplicapendiente")
    aplicable = json_object.item("aplicable")
    aplicacion = json_object.item("aplicacion")
    libroiva = json_object.item("libroiva")
    letracomprobante = json_object.item("letracomprobante")
    recibo = json_object.item("recibo")
    contado = json_object.item("contado")
    ordenpago = json_object.item("ordenpago")
    puntoventa = json_object.item("puntoventa")
    comprobanteafipId = json_object.item("comprobanteafipId")
    electronico = json_object.item("electronico")
    asociado = json_object.item("asociado")

End Sub

