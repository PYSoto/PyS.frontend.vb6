VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLCgosValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "cgosvalores"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "cgosvalores.clave"

End Property

Public Property Get negid() As String

    negid = "cva_neg_id"

End Property

Public Property Get negid_t() As String

    negid_t = "cgosvalores.cva_neg_id"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = "cgosvalores.codigo"

End Property

Public Property Get concepto() As String

    concepto = "concepto"

End Property

Public Property Get concepto_t() As String

    concepto_t = "cgosvalores.concepto"

End Property

Public Property Get numerable() As String

    numerable = "numerable"

End Property

Public Property Get numerable_t() As String

    numerable_t = "cgosvalores.numerable"

End Property

Public Property Get duplicados() As String

    duplicados = "duplicados"

End Property

Public Property Get duplicados_t() As String

    duplicados_t = "cgosvalores.duplicados"

End Property

Public Property Get fechaemi() As String

    fechaemi = "fechaemi"

End Property

Public Property Get fechaemi_t() As String

    fechaemi_t = "cgosvalores.fechaemi"

End Property

Public Property Get fechavto() As String

    fechavto = "fechavto"

End Property

Public Property Get fechavto_t() As String

    fechavto_t = "cgosvalores.fechavto"

End Property

Public Property Get titular() As String

    titular = "titular"

End Property

Public Property Get titular_t() As String

    titular_t = "cgosvalores.titular"

End Property

Public Property Get banco() As String

    banco = "banco"

End Property

Public Property Get banco_t() As String

    banco_t = "cgosvalores.banco"

End Property

Public Property Get chtercero() As String

    chtercero = "chtercero"

End Property

Public Property Get chtercero_t() As String

    chtercero_t = "cgosvalores.chtercero"

End Property

Public Property Get ctacte() As String

    ctacte = "ctacte"

End Property

Public Property Get ctacte_t() As String

    ctacte_t = "cgosvalores.ctacte"

End Property

