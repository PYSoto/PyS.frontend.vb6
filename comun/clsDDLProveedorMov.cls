VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLProveedorMov"
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

    table = "proveedor_movimiento"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID() As String

    proveedorID = "proveedor_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID_t() As String

    proveedorID_t = table & ".proveedor_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID() As String

    comprobanteID = "comprobante_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID_t() As String

    comprobanteID_t = table & ".comprobante_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia prefijo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijo() As String

    prefijo = "prefijo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia prefijo completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijo_t() As String

    prefijo_t = table & ".prefijo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nroComprobante
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nroComprobante() As String

    nroComprobante = "nro_comprobante"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nroComprobante completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nroComprobante_t() As String

    nroComprobante_t = table & ".nro_comprobante"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia empresaID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get empresaID() As String

    empresaID = "empresa_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia empresaID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get empresaID_t() As String

    empresaID_t = table & ".empresa_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID() As String

    negocioID = "negocio_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID_t() As String

    negocioID_t = table & ".negocio_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaComprobante
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaComprobante() As String

    fechaComprobante = "fecha_comprobante"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaComprobante completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaComprobante_t() As String

    fechaComprobante_t = table & ".fecha_comprobante"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaVencimiento
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaVencimiento() As String

    fechaVencimiento = "fecha_vencimiento"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaVencimiento completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaVencimiento_t() As String

    fechaVencimiento_t = table & ".fecha_vencimiento"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia total
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total() As String

    total = "total"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia total completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total_t() As String

    total_t = table & ".total"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia totalCancelado
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get totalCancelado() As String

    totalCancelado = "total_cancelado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia totalCancelado completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get totalCancelado_t() As String

    totalCancelado_t = table & ".total_cancelado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia neto
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto() As String

    neto = "neto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia neto completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get neto_t() As String

    neto_t = table & ".neto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva1
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva1() As String

    importeIva1 = "importe_iva1"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva1 completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva1_t() As String

    importeIva1_t = table & ".importe_iva1"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva2
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva2() As String

    importeIva2 = "importe_iva2"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva2 completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva2_t() As String

    importeIva2_t = table & ".importe_iva2"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva3
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva3() As String

    importeIva3 = "importe_iva3"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeIva3 completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeIva3_t() As String

    importeIva3_t = table & ".importe_iva3"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia percepcionIva
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIva() As String

    percepcionIva = "percepcion_iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia percepcionIva completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIva_t() As String

    percepcionIva_t = table & ".percepcion_iva"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia percepcionIIBB
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIIBB() As String

    percepcionIIBB = "percepcion_iibb"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia percepcionIIBB completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get percepcionIIBB_t() As String

    percepcionIIBB_t = table & ".percepcion_iibb"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia gastosNoGravados
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get gastosNoGravados() As String

    gastosNoGravados = "gastos_no_gravados"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia gastosNoGravados completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get gastosNoGravados_t() As String

    gastosNoGravados_t = table & ".gastos_no_gravados"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ajustes
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ajustes() As String

    ajustes = "ajustes"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia ajustes completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ajustes_t() As String

    ajustes_t = table & ".ajustes"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia monotributo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get monotributo() As String

    monotributo = "monotributo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia monotributo completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get monotributo_t() As String

    monotributo_t = table & ".monotributo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuentaMovimientoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaMovimientoID() As String

    cuentaMovimientoID = "cuenta_movimiento_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia cuentaMovimientoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuentaMovimientoID_t() As String

    cuentaMovimientoID_t = table & ".cuenta_movimiento_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia observaciones
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones() As String

    observaciones = "observaciones"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia observaciones completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones_t() As String

    observaciones_t = table & ".observaciones"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoID() As String

    proveedorMovimientoID = "proveedor_movimiento_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoID_t() As String

    proveedorMovimientoID_t = table & ".proveedor_movimiento_id"

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

