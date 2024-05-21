VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLTiposComprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "tiposcomprob"

End Property

Public Property Get negID() As String

    negID = "tco_neg_id"

End Property

Public Property Get negID_t() As String

    negID_t = table & ".tco_neg_id"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = table & ".codigo"

End Property

Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = table & ".descripcion"

End Property

Public Property Get modulo() As String

    modulo = "modulo"

End Property

Public Property Get modulo_t() As String

    modulo_t = table & ".modulo"

End Property

Public Property Get aplicapend() As String

    aplicapend = "aplicapend"

End Property

Public Property Get aplicapend_t() As String

    aplicapend_t = table & ".aplicapend"

End Property

Public Property Get ctacte() As String

    ctacte = "ctacte"

End Property

Public Property Get ctacte_t() As String

    ctacte_t = table & ".ctacte"

End Property

Public Property Get debita() As String

    debita = "debita"

End Property

Public Property Get debita_t() As String

    debita_t = table & ".debita"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = table & ".iva"

End Property

Public Property Get aplicable() As String

    aplicable = "aplicable"

End Property

Public Property Get aplicable_t() As String

    aplicable_t = table & ".aplicable"

End Property

Public Property Get libroiva() As String

    libroiva = "libroiva"

End Property

Public Property Get libroiva_t() As String

    libroiva_t = table & ".libroiva"

End Property

Public Property Get tipocomprob() As String

    tipocomprob = "tipocomprob"

End Property

Public Property Get tipocomprob_t() As String

    tipocomprob_t = table & ".tipocomprob"

End Property

Public Property Get recibo() As String

    recibo = "recibo"

End Property

Public Property Get recibo_t() As String

    recibo_t = table & ".recibo"

End Property

Public Property Get contado() As String

    contado = "contado"

End Property

Public Property Get contado_t() As String

    contado_t = table & ".contado"

End Property

Public Property Get puntovta() As String

    puntovta = "tco_puntovta"

End Property

Public Property Get puntovta_t() As String

    puntovta_t = table & ".tco_puntovta"

End Property

Public Property Get tipoafip() As String

    tipoafip = "tco_tipoafip"

End Property

Public Property Get tipoafip_t() As String

    tipoafip_t = table & ".tco_tipoafip"

End Property

Public Property Get factelect() As String

    factelect = "tco_factelect"

End Property

Public Property Get factelect_t() As String

    factelect_t = table & ".tco_factelect"

End Property

Public Property Get asociado() As String

    asociado = "asociado"

End Property

Public Property Get asociado_t() As String

    asociado_t = table & ".asociado"

End Property


