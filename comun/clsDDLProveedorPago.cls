VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLProveedorPago"
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

    table = "proveedor_pago"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoIDDeuda
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDDeuda() As String

    proveedorMovimientoIDDeuda = "proveedor_movimiento_id_deuda"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoIDDeuda completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDDeuda_t() As String

    proveedorMovimientoIDDeuda_t = table & ".proveedor_movimiento_id_deuda"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoIDAplicado
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDAplicado() As String

    proveedorMovimientoIDAplicado = "proveedor_movimiento_id_aplicado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia proveedorMovimientoIDAplicado completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDAplicado_t() As String

    proveedorMovimientoIDAplicado_t = table & ".proveedor_movimiento_id_aplicado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeAplicado
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeAplicado() As String

    importeAplicado = "importe_aplicado"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia importeAplicado completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeAplicado_t() As String

    importeAplicado_t = table & ".importe_aplicado"

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

