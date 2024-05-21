VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLValorMov"
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

    table = "valores"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID() As String

    negocioID = "val_neg_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID_t() As String

    negocioID_t = table & ".val_neg_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia valorID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorID() As String

    valorID = "codigo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia valorID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorID_t() As String

    valorID_t = table & ".codigo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia clienteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As String

    clienteID = "cgocli"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia clienteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID_t() As String

    clienteID_t = table & ".cgocli"

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
' Description: Asocia fechaEmision
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaEmision() As String

    fechaEmision = "fechaemi"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaEmision completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaEmision_t() As String

    fechaEmision_t = table & ".fechaemi"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaVencimiento
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaVencimiento() As String

    fechaVencimiento = "fechavto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaVencimiento completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaVencimiento_t() As String

    fechaVencimiento_t = table & ".fechavto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID() As String

    comprobanteID = "val_tco_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia comprobanteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comprobanteID_t() As String

    comprobanteID_t = table & ".val_tco_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nroComprobante
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nroComprobante() As String

    nroComprobante = "nrocomprob"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia nroComprobante completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nroComprobante_t() As String

    nroComprobante_t = table & ".nrocomprob"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importe
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importe() As String

    importe = "importe"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importe completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importe_t() As String

    importe_t = table & ".importe"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaRegistracion
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaRegistracion() As String

    fechaRegistracion = "fechareg"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fechaRegistracion completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaRegistracion_t() As String

    fechaRegistracion_t = table & ".fechareg"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia clientemovimientoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clientemovimientoID() As String

    clientemovimientoID = "clavemovv"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia clientemovimientoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clientemovimientoID_t() As String

    clientemovimientoID_t = table & ".clavemovv"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedormovimientoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedormovimientoID() As String

    proveedormovimientoID = "proveedormovimiento_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedormovimientoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedormovimientoID_t() As String

    proveedormovimientoID_t = table & ".proveedormovimiento_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia titular
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get titular() As String

    titular = "titular"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia titular completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get titular_t() As String

    titular_t = table & ".titular"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia banco
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get banco() As String

    banco = "banco"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia banco completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get banco_t() As String

    banco_t = table & ".banco"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia valormovimientoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valormovimientoID() As String

    valormovimientoID = "clave"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia valormovimientoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valormovimientoID_t() As String

    valormovimientoID_t = table & ".clave"

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

