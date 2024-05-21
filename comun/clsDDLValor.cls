VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLValor"
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

    table = "cgosvalores"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As String

    autoID = "clave"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia autoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID_t() As String

    autoID_t = table & ".clave"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID() As String

    negocioID = "cva_neg_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia negocioID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get negocioID_t() As String

    negocioID_t = table & ".cva_neg_id"

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
' Description: Asocia concepto
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get concepto() As String

    concepto = "concepto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia concepto completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get concepto_t() As String

    concepto_t = table & ".concepto"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia numerable
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerable() As String

    numerable = "numerable"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia numerable completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numerable_t() As String

    numerable_t = table & ".numerable"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia duplicados
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get duplicados() As String

    duplicados = "duplicados"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia duplicados completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get duplicados_t() As String

    duplicados_t = table & ".duplicados"

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
' Description: Asocia chequeTercero
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get chequeTercero() As String

    chequeTercero = "chtercero"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Asocia chequeTercero completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get chequeTercero_t() As String

    chequeTercero_t = table & ".chtercero"

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

