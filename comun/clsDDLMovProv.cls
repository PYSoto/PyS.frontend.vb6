VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLMovProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "movprov"

End Property

Public Property Get cgoProv() As String

    cgoProv = "cgoprov"

End Property

Public Property Get cgoProv_t() As String

    cgoProv_t = "movprov.cgoprov"

End Property

Public Property Get cgoComprob() As String

    cgoComprob = "cgocomprob"

End Property

Public Property Get cgoComprob_t() As String

    cgoComprob_t = "movprov.cgocomprob"

End Property

Public Property Get prefijo() As String

    prefijo = "prefijo"

End Property

Public Property Get prefijo_t() As String

    prefijo_t = "movprov.prefijo"

End Property

Public Property Get nroComprob() As String

    nroComprob = "nrocomprob"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "movprov.nrocomprob"

End Property

Public Property Get empID() As String

    empID = "mpr_emp_id"

End Property

Public Property Get empID_t() As String

    empID_t = "movprov.mpr_emp_id"

End Property

Public Property Get negID() As String

    negID = "mpr_neg_id"

End Property

Public Property Get negID_t() As String

    negID_t = "movprov.mpr_neg_id"

End Property

Public Property Get fechaComprob() As String

    fechaComprob = "fechacomprob"

End Property

Public Property Get fechaComprob_t() As String

    fechaComprob_t = "movprov.fechacomprob"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "movprov.importe"

End Property

Public Property Get neto() As String

    neto = "neto"

End Property

Public Property Get neto_t() As String

    neto_t = "movprov.neto"

End Property

Public Property Get montoIva() As String

    montoIva = "montoiva"

End Property

Public Property Get montoIva_t() As String

    montoIva_t = "movprov.montoiva"

End Property

Public Property Get montoIva27() As String

    montoIva27 = "montoiva27"

End Property

Public Property Get montoIva27_t() As String

    montoIva27_t = "movprov.montoiva27"

End Property

Public Property Get montoIva105() As String

    montoIva105 = "montoiva105"

End Property

Public Property Get montoIva105_t() As String

    montoIva105_t = "movprov.montoiva105"

End Property

Public Property Get percIva() As String

    percIva = "perciva"

End Property

Public Property Get percIva_t() As String

    percIva_t = "movprov.perciva"

End Property

Public Property Get percIngBrutos() As String

    percIngBrutos = "percingbrutos"

End Property

Public Property Get percIngBrutos_t() As String

    percIngBrutos_t = "movprov.percingbrutos"

End Property

Public Property Get gng() As String

    gng = "gng"

End Property

Public Property Get gng_t() As String

    gng_t = "movprov.gng"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "movprov.clave"

End Property

