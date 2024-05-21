VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia TABLA
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get table() As String

    table = "tiposcomprob"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID() As String

    negocioID = "tco_neg_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID_t() As String

    negocioID_t = table & ".tco_neg_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID() As String

    comprobanteID = "codigo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID_t() As String

    comprobanteID_t = table & ".codigo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia descripcion completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion_t() As String

    descripcion_t = table & ".descripcion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia tipoTransaccionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoTransaccionID() As String

    tipoTransaccionID = "modulo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia tipoTransaccionID completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoTransaccionID_t() As String

    tipoTransaccionID_t = table & ".modulo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicaPendiente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicaPendiente() As String

    aplicaPendiente = "aplicapend"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicaPendiente completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicaPendiente_t() As String

    aplicaPendiente_t = table & ".aplicapend"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuentaCorriente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaCorriente() As String

    cuentaCorriente = "ctacte"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuentaCorriente completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaCorriente_t() As String

    cuentaCorriente_t = table & ".ctacte"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia debita
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get debita() As String

    debita = "debita"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia debita completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get debita_t() As String

    debita_t = table & ".debita"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia iva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva() As String

    iva = "iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia iva completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get iva_t() As String

    iva_t = table & ".iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicable
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicable() As String

    aplicable = "aplicable"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicable completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicable_t() As String

    aplicable_t = table & ".aplicable"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicacion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicacion() As String

    aplicacion = "aplicacion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia aplicacion completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get aplicacion_t() As String

    aplicacion_t = table & ".aplicacion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia libroIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get libroIva() As String

    libroIva = "libroiva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia libroIva completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get libroIva_t() As String

    libroIva_t = table & ".libroiva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia letraComprobante
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letraComprobante() As String

    letraComprobante = "tipocomprob"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia letraComprobante completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letraComprobante_t() As String

    letraComprobante_t = table & ".tipocomprob"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia recibo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get recibo() As String

    recibo = "recibo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia recibo completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get recibo_t() As String

    recibo_t = table & ".recibo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia contado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contado() As String

    contado = "contado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia contado completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contado_t() As String

    contado_t = table & ".contado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ordenPago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ordenPago() As String

    ordenPago = "orden_pago"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ordenPago completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ordenPago_t() As String

    ordenPago_t = table & ".orden_pago"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia puntoVenta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get puntoVenta() As String

    puntoVenta = "tco_puntovta"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia puntoVenta completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get puntoVenta_t() As String

    puntoVenta_t = table & ".tco_puntovta"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia tipoAfip
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoAfip() As String

    tipoAfip = "tco_tipoafip"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia tipoAfip completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoAfip_t() As String

    tipoAfip_t = table & ".tco_tipoafip"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia facturaElectronica
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaElectronica() As String

    facturaElectronica = "tco_factelect"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia facturaElectronica completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaElectronica_t() As String

    facturaElectronica_t = table & ".tco_factelect"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As String

    autoID = "auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID_t() As String

    autoID_t = table & ".auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As String

    created = "created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia created completo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created_t() As String

    created_t = table & ".created"

End Property

