VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLProveedorSaldo"
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

    table = "proveedor_saldo"

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
' Description: Asocia fecha
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As String

    fecha = "fecha"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia fecha completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha_t() As String

    fecha_t = table & ".fecha"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia saldo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get saldo() As String

    saldo = "saldo"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia saldo completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get saldo_t() As String

    saldo_t = table & ".saldo"

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

